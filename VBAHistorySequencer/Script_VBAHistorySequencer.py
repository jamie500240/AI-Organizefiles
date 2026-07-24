# ==========================================================
# MODULE:       Script_VBAHistorySequencer
# PURPOSE:      VBA 歷史版本分析與模組血緣追蹤 (Office 跨平台版)
# EXPORTS:      Pipeline.run()
# IMPORTS:      os, sys, shutil, hashlib, pathlib, datetime, tkinter, csv, zipfile
# FORBIDDEN:    上帝物件、跨層依賴、隱性轉型、靜默覆寫、未授權命名、靜默失敗
# DEPENDENCIES: 作業系統檔案 I/O、Tkinter UI 環境
# VERSION:      1.2.0 [Stability: Experimental]
#
# [ADR-001] 關於 P0-11「全有或全無 (All-or-Nothing)」原則之豁免與取捨
# Context:  本系統處理之目標可能高達數百 GB，搬運與雜湊計算耗時極長。
# Decision: 遭遇 Ctrl+C (KeyboardInterrupt) 或單一檔案 I/O 異常時，不執行「全數退回 (Rollback)」，而是保留已處理之檔案並產出結算報表。
# Rationale:在巨量檔案處理情境下，銷毀數小時的成功進度會造成極差的 UX。保留 Partial State 並透過 CSV 報表確保資料狀態具備完全的可稽核性，為此情境下之最佳實務。
# [ADR-002] 關於 增加可處理物件的原因
# Context:  本系統處理之目標只能是 EXCEL 不夠泛用。
# Decision: 開放 WORD 跟 PPT 都可用。
# ==========================================================

import os
import sys
import shutil
import hashlib
import csv
import re
import uuid
import zipfile
from time import time, sleep
from datetime import datetime
from pathlib import Path
from dataclasses import dataclass, field
from abc import ABC, abstractmethod
from concurrent.futures import ThreadPoolExecutor, TimeoutError

# ==========================================
# 1. ENCAPSULATED CONFIGS & STRINGS (SSOT)
# ==========================================
@dataclass(frozen=True)
class AppConfig:
    SHA_LEN: int = 16
    # 【擴展】支援所有主流 Office 巨集格式
    OFFICE_EXTS: tuple = ('.xls', '.xlsx', '.xlsm', '.xlsb', '.doc', '.docm', '.ppt', '.pptm')
    ILLEGAL_CHAR_PATTERN: str = r'[\\/*?:"<>|]'
    RESERVED_NAMES_PATTERN: str = r'^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$'
    
    # 【統一管理的防禦與效能閾值】
    MAX_WORKERS: int = 4
    TIMEOUT_ANCHOR_SEC: float = 15.0
    TIMEOUT_PER_FILE_SEC: float = 10.0
    # 【防禦】泛化 OLE 舊檔記憶體保護，防止 .doc, .ppt, .xls 造成 OOM
    MAX_LEGACY_MB: float = 20.0
    MAX_VBA_CODE_LENGTH: int = 2 * 1024 * 1024  # 2MB 字元上限 (防止 OOM)

@dataclass(frozen=True)
class AppStrings:
    DIR_WORKSPACE_PREFIX: str = "抽取巨集"
    DIR_STAGE_PREFIX: str = "_STAGE_TEMP_"
    DIR_CLONES: str = "_CLONES_隔離區"
    DIR_FAILED: str = "_FAILED_解析失敗"  
    REPORT_FILE: str = "_LOG_Evolution_Tree.csv"
    
    STATUS_UNIQUE: str = "MUTATION (突變/新創)"
    STATUS_DUP: str = "INHERITED (繼承)"
    STATUS_CLONE: str = "CLONED (純複製體)"
    STATUS_FAIL: str = "PARSE_FAILED (解析失敗)"
    
    PREFIX_DUP: str = "[DUP]"

# ==========================================
# 2. 環境依賴防禦與動態引擎掛載 (Fail-Fast)
# ==========================================
try:
    from oletools.olevba import VBA_Parser
except ImportError:
    raise RuntimeError("【環境錯誤】缺少二進位解析套件。請執行: pip install oletools")

try:
    from rapidfuzz import fuzz
    HAS_RAPIDFUZZ = True
except ImportError:
    import difflib
    HAS_RAPIDFUZZ = False

# ==========================================
# 3. DOMAIN MODELS & DTO
# ==========================================
@dataclass
class EvolutionNode:
    file_path: Path
    is_valid: bool = True
    error_msg: str = ""
    module_names: list[str] = field(default_factory=list)
    module_hashes: dict = field(default_factory=dict)
    modules_code: dict = field(default_factory=dict) 
    similarity_score: float = 0.0
    generation: int = 0
    folder_name: str = ""
    is_clone: bool = False

@dataclass
class ProcessPayload:
    source_dir: Path
    output_dir: Path = None  
    anchor_file: Path = None
    config: AppConfig = field(default_factory=AppConfig)
    strings: AppStrings = field(default_factory=AppStrings)
    
    staging_dir: Path = None
    final_dir: Path = None
    
    nodes: list[EvolutionNode] = field(default_factory=list)
    anchor_modules: dict = field(default_factory=dict)
    stats: dict = field(default_factory=lambda: {"mutations": 0, "inherited": 0, "clones": 0, "failed": 0})

# ==========================================
# 4. UTILITY 
# ==========================================
class Sanitizer:
    @staticmethod
    def clean_filename(name: str, config: AppConfig) -> str:
        cleaned = re.sub(config.ILLEGAL_CHAR_PATTERN, '_', name)
        base_upper = cleaned.split('.')[0].upper()
        if re.match(config.RESERVED_NAMES_PATTERN, base_upper):
            cleaned = f"SAFE_{cleaned}"
        cleaned = cleaned.rstrip(' .')
        return cleaned if cleaned else f"UNNAMED_{uuid.uuid4().hex[:8]}"

class SimilarityEngine:
    @staticmethod
    def calculate_weighted_ratio(anchor_mods: dict, target_mods: dict) -> float:
        all_names = set(anchor_mods.keys()) | set(target_mods.keys())
        total_score = 0.0
        total_weight = 0.0
        
        for name in all_names:
            code_a = anchor_mods.get(name, "")
            code_t = target_mods.get(name, "")
            weight = max(len(code_a), len(code_t))
            
            if weight == 0: continue
                
            if code_a == code_t:
                ratio = 1.0
            else:
                if HAS_RAPIDFUZZ:
                    ratio = fuzz.ratio(code_a, code_t) / 100.0
                else:
                    ratio = difflib.SequenceMatcher(None, code_a, code_t, autojunk=False).ratio()
                
            total_score += ratio * weight
            total_weight += weight
            
        return (total_score / total_weight * 100) if total_weight > 0 else 0.0

class VbaExtractor:
    @staticmethod
    def extract(file_path: Path, config: AppConfig) -> tuple[bool, dict, dict, str]:
        modules, hashes = {}, {}
        vba_parser = None
        suffix = file_path.suffix.lower()
        
        # 【擴展】加入 .docm 與 .pptm 支援 ZIP 結構解析
        if suffix in ('.xlsm', '.xlsb', '.xlsx', '.docm', '.pptm'):
            try:
                with zipfile.ZipFile(file_path, 'r') as z:
                    vba_target = next((name for name in z.namelist() if name.lower().endswith('vbaproject.bin')), None)
                    if vba_target:
                        with z.open(vba_target) as vba_bin:
                            vba_data = vba_bin.read() 
                        
                        vba_parser = VBA_Parser(file_path.name, data=vba_data)
                        if vba_parser.detect_vba_macros():
                            for (_, _, vba_filename, vba_code) in vba_parser.extract_macros():
                                if vba_code and vba_code.strip():
                                    modules[vba_filename] = vba_code
                                    hashes[vba_filename] = hashlib.sha256(vba_code.encode('utf-8', errors='ignore')).hexdigest()
                        return True, modules, hashes, ""
            except Exception as e:
                return False, {}, {}, f"ZIP Extract Failed: {e}"
            finally:
                if vba_parser:
                    try: vba_parser.close()
                    except: pass

        try:
            file_size_mb = file_path.stat().st_size / (1024 * 1024)
            # 【防禦】泛化為對所有舊版 OLE 檔案進行記憶體保護
            if suffix in ('.xls', '.doc', '.ppt') and file_size_mb > config.MAX_LEGACY_MB:
                return False, {}, {}, f"Legacy file '{suffix}' too large ({file_size_mb:.1f}MB). Skipped to protect memory."

            with open(file_path, 'rb') as f:
                file_bytes = f.read()
            vba_parser = VBA_Parser(file_path.name, data=file_bytes)
            if vba_parser.detect_vba_macros():
                for (_, _, vba_filename, vba_code) in vba_parser.extract_macros():
                    if vba_code and vba_code.strip():
                        modules[vba_filename] = vba_code
                        hashes[vba_filename] = hashlib.sha256(vba_code.encode('utf-8', errors='ignore')).hexdigest()
            return True, modules, hashes, ""
        except Exception as e:
            return False, {}, {}, str(e)
        finally:
            if vba_parser:
                try: vba_parser.close()
                except: pass

# ==========================================
# 5. ACTIONS (PIPELINE)
# ==========================================
class IAction(ABC):
    @abstractmethod
    def execute(self, payload: ProcessPayload) -> ProcessPayload:
        pass

class ActionPreScanAndSequence(IAction):
    def execute(self, payload: ProcessPayload) -> ProcessPayload:
        engine_name = "RapidFuzz (極速模式)" if HAS_RAPIDFUZZ else "Difflib (標準模式)"
        print(f"\n🧬 [階段 1/4] 啟動模組比對與演化定序 | 載入引擎: {engine_name}")
        
        # 改採 OFFICE_EXTS 過濾器
        target_files = [f for f in payload.source_dir.rglob('*') if f.is_file() and f.suffix.lower() in payload.config.OFFICE_EXTS]
        if not target_files:
            raise ValueError("找不到任何支援的 Office 檔案。")

        print("🔍 正在解析錨點檔案...")
        with ThreadPoolExecutor(max_workers=1) as executor:
            future = executor.submit(VbaExtractor.extract, payload.anchor_file, payload.config)
            try:
                is_valid, payload.anchor_modules, _, err = future.result(timeout=payload.config.TIMEOUT_ANCHOR_SEC)
            except TimeoutError:
                raise RuntimeError("⚠️ 錨點檔案解析超時！該檔案可能損壞或過於巨大。")

        if not is_valid or not payload.anchor_modules:
            raise RuntimeError(f"⚠️ 錨點解析失敗或無巨集：{err}")

        executor = ThreadPoolExecutor(max_workers=payload.config.MAX_WORKERS)
        futures = [(f_path, executor.submit(VbaExtractor.extract, f_path, payload.config)) for f_path in target_files]

        for idx, (f_path, future) in enumerate(futures, 1):
            sys.stdout.write(f"\r🔍 基因掃描中 (併發解析): {idx} / {len(target_files)}")
            sys.stdout.flush()
            
            try:
                is_valid, modules, hashes, err = future.result(timeout=payload.config.TIMEOUT_PER_FILE_SEC)
            except TimeoutError:
                is_valid, modules, hashes, err = False, {}, {}, f"Parser Timeout (> {payload.config.TIMEOUT_PER_FILE_SEC}s)"
            except Exception as ex:
                is_valid, modules, hashes, err = False, {}, {}, f"Thread Error: {ex}"
            
            node = EvolutionNode(
                file_path=f_path, is_valid=is_valid, error_msg=err, 
                module_names=list(modules.keys()), module_hashes=hashes, modules_code=modules
            )
            
            if is_valid:
                total_char_len = sum(len(v) for v in modules.values())
                if total_char_len > payload.config.MAX_VBA_CODE_LENGTH:
                    node.is_valid = False
                    node.error_msg = f"VBA code size too massive ({total_char_len} chars), bypassed."
                    node.modules_code = {} 
                else:
                    node.similarity_score = SimilarityEngine.calculate_weighted_ratio(payload.anchor_modules, modules)
            
            payload.nodes.append(node)
        
        executor.shutdown(wait=False)
        print("\n✅ 比對完成。")

        valid_nodes = [n for n in payload.nodes if n.is_valid]
        valid_nodes.sort(key=lambda n: (n.similarity_score, len(n.module_names)))
        
        for generation, node in enumerate(valid_nodes, 1):
            node.generation = generation
            safe_file_name = Sanitizer.clean_filename(node.file_path.name, payload.config)
            node.folder_name = f"Gen{generation:03d}_{node.similarity_score:05.1f}%_{safe_file_name}"
            
        return payload

class ActionInitializeWorkspace(IAction):
    def execute(self, payload: ProcessPayload) -> ProcessPayload:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        base_dir = payload.output_dir if payload.output_dir else Path.cwd()
        
        payload.final_dir = base_dir / f"{payload.strings.DIR_WORKSPACE_PREFIX}_{timestamp}"
        payload.staging_dir = base_dir / f"{payload.strings.DIR_STAGE_PREFIX}_{timestamp}"
        
        if payload.final_dir.exists() or payload.staging_dir.exists():
            raise RuntimeError("發生資料夾名稱碰撞，請稍後再試。")
            
        payload.staging_dir.mkdir(parents=True, exist_ok=False)
        (payload.staging_dir / payload.strings.DIR_CLONES).mkdir(parents=True, exist_ok=False)
        (payload.staging_dir / payload.strings.DIR_FAILED).mkdir(parents=True, exist_ok=False)
        return payload

class ActionExtractAndStage(IAction):
    def execute(self, payload: ProcessPayload) -> ProcessPayload:
        print("\n🏗️ [階段 2/4 & 3/4] 啟動隔離區間實體寫入與流式報表生成...")
        
        global_hash_registry = set()
        global_genome_registry = set()
        sorted_nodes = sorted(payload.nodes, key=lambda n: n.generation)
        
        report_path = payload.staging_dir / payload.strings.REPORT_FILE
        
        with open(report_path, "w", newline="", encoding="utf-8-sig") as f:
            csv_writer = csv.writer(f)
            csv_writer.writerow(["世代(Gen)", "相似度", "來源檔案", "模組名稱", "DNA碼", "演化狀態", "備註"])
            
            for node in sorted_nodes:
                try:
                    safe_file_name = Sanitizer.clean_filename(node.file_path.name, payload.config)
                    
                    if not node.is_valid:
                        payload.stats["failed"] += 1
                        try:
                            shutil.copy2(node.file_path, payload.staging_dir / payload.strings.DIR_FAILED / node.file_path.name)
                        except Exception as copy_err:
                            node.error_msg += f" (I/O Error: {copy_err})"
                        
                        csv_writer.writerow(["N/A", "N/A", safe_file_name, "N/A", "N/A", payload.strings.STATUS_FAIL, node.error_msg])
                        f.flush()
                        continue
                        
                    file_genome = frozenset(node.module_hashes.values())
                    if file_genome in global_genome_registry and file_genome:
                        node.is_clone = True
                        payload.stats["clones"] += 1
                        shutil.copy2(node.file_path, payload.staging_dir / payload.strings.DIR_CLONES / node.file_path.name)
                        csv_writer.writerow([str(node.generation), f"{node.similarity_score:.1f}%", safe_file_name, "[ALL_MODULES]", "N/A", payload.strings.STATUS_CLONE, "100% 基因重疊，已下放 CLONES 隔離區"])
                        f.flush()
                        continue
                        
                    global_genome_registry.add(file_genome)
                    current_modules = node.modules_code 
                    
                    node_target_dir = payload.staging_dir / node.folder_name
                    node_target_dir.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(node.file_path, node_target_dir / node.file_path.name)
                    
                    for vba_filename, vba_code in current_modules.items():
                        dna = node.module_hashes.get(vba_filename, "")
                        if not dna: continue
                        
                        sha_short = dna[:payload.config.SHA_LEN]
                        safe_mod_name = Sanitizer.clean_filename(vba_filename, payload.config)
                        
                        is_mutation = dna not in global_hash_registry
                        global_hash_registry.add(dna)
                        
                        if is_mutation:
                            status = payload.strings.STATUS_UNIQUE
                            payload.stats["mutations"] += 1
                            final_mod_name = f"{safe_mod_name}_{sha_short}.txt"
                        else:
                            status = payload.strings.STATUS_DUP
                            payload.stats["inherited"] += 1
                            final_mod_name = f"{payload.strings.PREFIX_DUP}_{safe_mod_name}_{sha_short}.txt"
                        
                        with open(node_target_dir / final_mod_name, "w", encoding="utf-8") as out_f:
                            out_f.write(vba_code)
                            
                        csv_writer.writerow([str(node.generation), f"{node.similarity_score:.1f}%", safe_file_name, safe_mod_name, sha_short, status, ""])
                    
                    f.flush()
                
                except Exception as file_io_err:
                    payload.stats["failed"] += 1
                    fallback_file_name = str(node.file_path.name) if node.file_path else "UNKNOWN_FILE"
                    
                    err_str = f"檔案處理異常 ({type(file_io_err).__name__}): {file_io_err}"
                    csv_writer.writerow([str(node.generation), f"{node.similarity_score:.1f}%", fallback_file_name, "N/A", "N/A", payload.strings.STATUS_FAIL, err_str])
                    f.flush()
                    
                    try:
                        shutil.copy2(node.file_path, payload.staging_dir / payload.strings.DIR_FAILED / node.file_path.name)
                    except:
                        pass
                    continue
                    
        return payload

class ActionCommitAndFinalize(IAction):
    def execute(self, payload: ProcessPayload) -> ProcessPayload:
        print("\n🔒 [階段 4/4] Atomic Operation: 執行 O(1) 終極更名提交 (Two-Phase Commit)...")
        payload.staging_dir.rename(payload.final_dir)
        return payload

# ==========================================
# 6. FLOW / RUNNER
# ==========================================
class Pipeline:
    def __init__(self, actions: list[IAction]):
        self._actions = actions

    def run(self, payload: ProcessPayload) -> ProcessPayload:
        current_payload = payload
        for action in self._actions:
            current_payload = action.execute(current_payload)
        return current_payload

def force_rollback(payload: ProcessPayload):
    if not payload or not payload.staging_dir: return
    if payload.staging_dir.exists():
        print(f"\n🗑️ 系統層級嚴重異常，正在銷毀隔離暫存區 {payload.staging_dir.name} ...")
        for attempt in range(3):
            try:
                shutil.rmtree(payload.staging_dir, ignore_errors=False)
                print("✅ 殘留狀態清除完畢。")
                return
            except PermissionError:
                sleep(0.5)
            except Exception as e:
                print(f"⚠️ 清除暫存區失敗：{e}")
                return

# ==========================================
# 7. CLIENT ENTRY POINT 
# ==========================================
if __name__ == "__main__":
    import tkinter as tk
    from tkinter import filedialog

    initial_payload = None  # 先宣告，避免 finally 引用時 NameError

    try:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)

        print("\n" + "="*60)
        print(" 🧬 Office VBA 歷史定序引擎 啟動")
        print("="*60)

        target_folder = filedialog.askdirectory(title="【1. 選擇全庫來源】請選擇包含所有檔案的資料夾")
        if not target_folder:
            print("已取消操作。")
        else:
            anchor_file = filedialog.askopenfilename(
                title="【2. 指定進化終點】請選擇「最終定稿」的 Office 檔案",
                initialdir=target_folder,
                filetypes=[("Office Macros", "*.xls* *.doc* *.ppt*"), ("All Files", "*.*")]
            )
            if not anchor_file:
                print("必須指定最終定稿才能進行演化反推。")
            else:
                default_out_dir = str(Path(target_folder).parent)
                output_folder = filedialog.askdirectory(
                    title="【3. 選擇輸出位置】請選擇分析報告與提取檔案的儲存目錄",
                    initialdir=default_out_dir
                )
                if not output_folder:
                    print("未指定輸出路徑，已取消操作。")
                else:
                    root.destroy()

                    cleanup_flow = Pipeline([
                        ActionPreScanAndSequence(),
                        ActionInitializeWorkspace(),
                        ActionExtractAndStage(),
                        ActionCommitAndFinalize()
                    ])

                    initial_payload = ProcessPayload(
                        source_dir=Path(target_folder),
                        anchor_file=Path(anchor_file),
                        output_dir=Path(output_folder)
                    )

                    start_time = time()
                    final_result = cleanup_flow.run(initial_payload)

                    print("\n" + "="*60)
                    print(f"✅ 歷史時光機重建完畢 (耗時: {time() - start_time:.2f} 秒)")
                    print(f"📊 突變/首創模組 (MUTATION): {final_result.stats['mutations']}")
                    print(f"📊 沿用舊有模組 (INHERITED): {final_result.stats['inherited']}")
                    print(f"👻 隔離純複製體 (CLONES): {final_result.stats['clones']} 個檔案")
                    if final_result.stats['failed'] > 0:
                        print(f"⚠️ 解析失敗隔離 (FAILED): {final_result.stats['failed']} 個檔案")
                    print(f"📂 本次輸出父夾: {final_result.final_dir}")
                    print("="*60)

    except KeyboardInterrupt:
        print("\n\n🛑 [使用者中斷] 偵測到手動停止 (Ctrl+C)。")
        print("依照 [ADR-001] 原則，已停止執行並保留當前處理進度。")
        if initial_payload and initial_payload.staging_dir and initial_payload.staging_dir.exists():
            print(f"📂 斷點進度與報表已保留於: {initial_payload.staging_dir}")

    except Exception as e:
        force_rollback(initial_payload)
        print(f"\n💥 【系統架構級嚴重終止】原因: {e}")

    finally:
        print("\n")
        os.system('pause' if os.name == 'nt' else 'read -p "Press Enter to continue..."')
        os._exit(0)
