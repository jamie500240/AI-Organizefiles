# ==========================================================
# MODULE:       Script_VBAHistorySequencer_1.0.0 [Stability: Experimental]
# PURPOSE:      VBA 歷史版本分析與模組血緣追蹤
# EXPORTS:      Script_AppRouter.run()
# IMPORTS:      os, shutil, hashlib, pathlib, datetime, tkinter, csv
# FORBIDDEN:    上帝物件、跨層依賴、隱性轉型、靜默覆寫、未授權命名、靜默失敗
# DEPENDENCIES: 作業系統檔案 I/O、Tkinter UI 環境
# VERSION:      1.0.0 [Stability: Experimental]
#
# [ADR-001] 關於 P0-11「全有或全無 (All-or-Nothing)」原則之豁免與取捨
# Context:  本系統處理之目標可能高達數百 GB，搬運與雜湊計算耗時極長。
# Decision: 遭遇 Ctrl+C (KeyboardInterrupt) 時，不執行「全數退回 (Rollback)」，而是保留已處理之檔案並產出結算報表。
# Rationale:在巨量檔案處理情境下，銷毀數小時的成功進度會造成極差的 UX。保留 Partial State 並透過 CSV 報表確保資料狀態具備完全的可稽核性，為此情境下之最佳實務。
# ==========================================================

import os
import sys
import shutil
import hashlib
import csv
import re
import uuid
import difflib
from time import time, sleep
from datetime import datetime
from pathlib import Path
from dataclasses import dataclass, field
from abc import ABC, abstractmethod

# ==========================================
# 0. 環境依賴防禦與動態引擎掛載 (Fail-Fast & Graceful Degradation)
# ==========================================
try:
    from oletools.olevba import VBA_Parser
except ImportError:
    raise RuntimeError("【環境錯誤】缺少二進位解析套件。請執行: pip install oletools")

try:
    from rapidfuzz import fuzz
    HAS_RAPIDFUZZ = True
except ImportError:
    HAS_RAPIDFUZZ = False

# ==========================================
# 1. ENCAPSULATED CONFIGS & STRINGS (SSOT)
# ==========================================
@dataclass(frozen=True)
class AppConfig:
    SHA_LEN: int = 16
    EXCEL_EXTS: tuple = ('.xls', '.xlsx', '.xlsm', '.xlsb')
    ILLEGAL_CHAR_PATTERN: str = r'[\\/*?:"<>|]'
    RESERVED_NAMES_PATTERN: str = r'^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$'

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
# 2. DOMAIN MODELS & DTO
# ==========================================
@dataclass
class EvolutionNode:
    file_path: Path
    is_valid: bool = True
    error_msg: str = ""
    module_names: list[str] = field(default_factory=list)
    module_hashes: dict = field(default_factory=dict)
    similarity_score: float = 0.0
    generation: int = 0
    folder_name: str = ""
    is_clone: bool = False

@dataclass
class ReportRecord:
    generation: str
    source_excel: str
    similarity_score: str
    module_name: str
    dna_short: str
    status: str
    details: str = ""

@dataclass
class ProcessPayload:
    source_dir: Path
    anchor_file: Path = None
    config: AppConfig = field(default_factory=AppConfig)
    strings: AppStrings = field(default_factory=AppStrings)
    
    # 【核心重構：嚴格區分暫存區與最終定案區】
    staging_dir: Path = None
    final_dir: Path = None
    
    nodes: list[EvolutionNode] = field(default_factory=list)
    anchor_modules: dict = field(default_factory=dict)
    report_data: list[ReportRecord] = field(default_factory=list)
    stats: dict = field(default_factory=lambda: {"mutations": 0, "inherited": 0, "clones": 0, "failed": 0})

# ==========================================
# 3. UTILITY 
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
            
            if weight == 0:
                continue
                
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
    def extract(file_path: Path) -> tuple[bool, dict, dict, str]:
        modules, hashes = {}, {}
        vba_parser = None
        try:
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
# 4. ACTIONS (PIPELINE)
# ==========================================
class IAction(ABC):
    @abstractmethod
    def execute(self, payload: ProcessPayload) -> ProcessPayload:
        pass

class ActionPreScanAndSequence(IAction):
    def execute(self, payload: ProcessPayload) -> ProcessPayload:
        engine_name = "RapidFuzz (極速模式)" if HAS_RAPIDFUZZ else "Difflib (標準模式)"
        print(f"\n🧬 [階段 1/4] 啟動模組比對與演化定序 | 載入引擎: {engine_name}")
        
        target_files = [f for f in payload.source_dir.glob('*') if f.is_file() and f.suffix.lower() in payload.config.EXCEL_EXTS]
        if not target_files:
            raise ValueError("找不到任何 Excel 檔案。")

        # 解析 Anchor
        is_valid, payload.anchor_modules, _, err = VbaExtractor.extract(payload.anchor_file)
        if not is_valid or not payload.anchor_modules:
            raise RuntimeError(f"⚠️ 錨點解析失敗或無巨集：{err}")

        # 第一階段掃描：不儲存 code_string，僅存 Metadata 避免 OOM
        for idx, f_path in enumerate(target_files, 1):
            is_valid, modules, hashes, err = VbaExtractor.extract(f_path)
            node = EvolutionNode(
                file_path=f_path, is_valid=is_valid, error_msg=err, 
                module_names=list(modules.keys()), module_hashes=hashes
            )
            
            if is_valid:
                node.similarity_score = SimilarityEngine.calculate_weighted_ratio(payload.anchor_modules, modules)
            
            payload.nodes.append(node)
            sys.stdout.write(f"\r🔍 基因掃描中 (Low-Memory Mode): {idx} / {len(target_files)}")
            sys.stdout.flush()
            
        print("\n✅ 比對完成。")

        # 排序與賦名
        valid_nodes = [n for n in payload.nodes if n.is_valid]
        valid_nodes.sort(key=lambda n: (n.similarity_score, len(n.module_names)))
        
        for generation, node in enumerate(valid_nodes, 1):
            node.generation = generation
            safe_excel_name = Sanitizer.clean_filename(node.file_path.name, payload.config)
            node.folder_name = f"Gen{generation:03d}_{node.similarity_score:05.1f}%_{safe_excel_name}"
            
        return payload

class ActionInitializeWorkspace(IAction):
    def execute(self, payload: ProcessPayload) -> ProcessPayload:
        python_base_dir = Path(__file__).resolve().parent if '__file__' in globals() else Path.cwd()
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        payload.final_dir = python_base_dir / f"{payload.strings.DIR_WORKSPACE_PREFIX}_{timestamp}"
        payload.staging_dir = python_base_dir / f"{payload.strings.DIR_STAGE_PREFIX}_{timestamp}"
        
        if payload.final_dir.exists() or payload.staging_dir.exists():
            raise RuntimeError("發生資料夾名稱碰撞，請一秒後再試。")
            
        # 僅在 Staging 區建立結構，Final 區留待最後 Commit 時再生成
        payload.staging_dir.mkdir(parents=True, exist_ok=False)
        (payload.staging_dir / payload.strings.DIR_CLONES).mkdir(parents=True, exist_ok=False)
        (payload.staging_dir / payload.strings.DIR_FAILED).mkdir(parents=True, exist_ok=False)
        return payload

class ActionExtractAndStage(IAction):
    def execute(self, payload: ProcessPayload) -> ProcessPayload:
        print("\n🏗️ [階段 2/4] 啟動隔離 Staging 區間建置與實體檔案寫入...")
        
        global_hash_registry = set()
        global_genome_registry = set()
        sorted_nodes = sorted(payload.nodes, key=lambda n: n.generation)
        
        for node in sorted_nodes:
            safe_excel_name = Sanitizer.clean_filename(node.file_path.name, payload.config)
            
            # 處理失效檔案
            if not node.is_valid:
                payload.stats["failed"] += 1
                shutil.copy2(node.file_path, payload.staging_dir / payload.strings.DIR_FAILED / node.file_path.name)
                payload.report_data.append(ReportRecord(
                    generation="N/A", source_excel=safe_excel_name, similarity_score="N/A",
                    module_name="N/A", dna_short="N/A", status=payload.strings.STATUS_FAIL, details=node.error_msg
                ))
                continue
                
            # 處理純複製體
            file_genome = frozenset(node.module_hashes.values())
            if file_genome in global_genome_registry and file_genome:
                node.is_clone = True
                payload.stats["clones"] += 1
                shutil.copy2(node.file_path, payload.staging_dir / payload.strings.DIR_CLONES / node.file_path.name)
                payload.report_data.append(ReportRecord(
                    generation=str(node.generation), source_excel=safe_excel_name, similarity_score=f"{node.similarity_score:.1f}%",
                    module_name="[ALL_MODULES]", dna_short="N/A", status=payload.strings.STATUS_CLONE,
                    details="100% 基因重疊，已下放 CLONES 隔離區"
                ))
                continue
                
            global_genome_registry.add(file_genome)
            
            # 【Time-Space Tradeoff】針對有效的變異節點進行二次提取
            _, current_modules, _, _ = VbaExtractor.extract(node.file_path)
            
            node_target_dir = payload.staging_dir / node.folder_name
            node_target_dir.mkdir(parents=True, exist_ok=True)
            
            # 【效能優化】順便把原版 Excel 備份進 Staging 節點區
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
                    # 突變模組直接寫入最終檔名 (無前綴)
                    final_mod_name = f"{safe_mod_name}_{sha_short}.txt"
                else:
                    status = payload.strings.STATUS_DUP
                    payload.stats["inherited"] += 1
                    # 繼承模組加上 DUP 前綴
                    final_mod_name = f"{payload.strings.PREFIX_DUP}_{safe_mod_name}_{sha_short}.txt"
                
                with open(node_target_dir / final_mod_name, "w", encoding="utf-8") as out_f:
                    out_f.write(vba_code)
                    
                payload.report_data.append(ReportRecord(
                    generation=str(node.generation), source_excel=safe_excel_name, similarity_score=f"{node.similarity_score:.1f}%",
                    module_name=safe_mod_name, dna_short=sha_short, status=status
                ))
                
        # 產生 CSV 報表 (寫入 Staging)
        print("📄 [階段 3/4] 產生血緣演化報表...")
        report_path = payload.staging_dir / payload.strings.REPORT_FILE
        with open(report_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(["世代(Gen)", "相似度", "來源檔案", "模組名稱", "DNA碼", "演化狀態", "備註"])
            for record in reversed(payload.report_data):
                writer.writerow([
                    record.generation, record.similarity_score, record.source_excel, 
                    record.module_name, record.dna_short, record.status, record.details
                ])
                
        return payload

class ActionCommitAndFinalize(IAction):
    def execute(self, payload: ProcessPayload) -> ProcessPayload:
        print("\n🔒 [階段 4/4] Atomic Operation: 執行 O(1) 終極更名提交 (Two-Phase Commit)...")
        # 作業系統層級的指標抽換。一旦成功，就是完全成功；失敗，則由全域的 Rollback 負責清理 staging。
        payload.staging_dir.rename(payload.final_dir)
        return payload

# ==========================================
# 5. FLOW / RUNNER
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
    if not payload or not payload.staging_dir:
        return
        
    if payload.staging_dir.exists():
        print(f"\n🗑️ 偵測到中斷，正在銷毀隔離暫存區 {payload.staging_dir.name} ...")
        for attempt in range(5):
            try:
                shutil.rmtree(payload.staging_dir, ignore_errors=False)
                print("✅ 殘留狀態清除完畢，維持系統潔淨。")
                return
            except PermissionError:
                if attempt < 4: sleep(0.5)
            except Exception as e:
                print(f"⚠️ 清除暫存區失敗：{e}")

# ==========================================
# 6. CLIENT ENTRY POINT 
# ==========================================
if __name__ == "__main__":
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    
    print("\n" + "="*60)
    print(" 🧬 VBA 歷史定序引擎 (低記憶體 2PC 工業終極版) 啟動")
    print("="*60)
    
    target_folder = filedialog.askdirectory(title="【1. 選擇全庫來源】請選擇包含所有檔案的資料夾")
    if not target_folder: 
        root.destroy()
        sys.exit("已取消操作。")

    anchor_file = filedialog.askopenfilename(
        title="【2. 指定進化終點】請選擇「最終定稿」的 Excel 檔案",
        initialdir=target_folder,
        filetypes=[("Excel Files", "*.xls*")]
    )
    
    # 確保 GUI 資源第一時間釋放
    root.destroy()
    
    if not anchor_file: sys.exit("必須指定最終定稿才能進行演化反推。")

    cleanup_flow = Pipeline([
        ActionPreScanAndSequence(),
        ActionInitializeWorkspace(),
        ActionExtractAndStage(),
        ActionCommitAndFinalize()
    ])

    initial_payload = ProcessPayload(source_dir=Path(target_folder), anchor_file=Path(anchor_file))
    
    try:
        start_time = time()
        final_result = cleanup_flow.run(initial_payload)
        
        print("\n" + "="*60)
        print(f"✅ 歷史時光機重建完畢 (耗時: {time() - start_time:.2f} 秒)")
        print(f"📊 突變/首創模組 (MUTATION): {final_result.stats['mutations']}")
        print(f"📊 沿用舊有模組 (INHERITED): {final_result.stats['inherited']}")
        print(f"👻 隔離純複製體 (CLONES): {final_result.stats['clones']} 個檔案")
        if final_result.stats['failed'] > 0:
            print(f"⚠️ 解析失敗隔離 (FAILED): {final_result.stats['failed']} 個檔案")
        print(f"📂 本次輸出父夾: {final_result.final_dir.name}")
        print("="*60)
        
    except Exception as e:
        force_rollback(initial_payload)
        print(f"\n💥 【程式終止】原因: {e}")
        sys.exit(1)
