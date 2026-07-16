# ==========================================================
# MODULE:      Script_SafeRenamer
# PURPOSE:     安全批次改名工具：支援雙模式 (改名/復原)、外部 JSON 設定檔動態載入
# EXPORTS:     RenameFlow, RevertFlow, RuleFactory, NameProcessor, FileOps
# IMPORTS:     os, shutil, tkinter, pathlib, dataclasses, typing, re, datetime, json, hashlib
# FORBIDDEN:   禁止對「原始輸入檔案」(source_dir 內的檔案)執行 move 或 rename；
#              禁止靜默吞沒例外 (except: pass)。
#              例外：RevertFlow 依 ADR-007，對「成功區內部」已複製完成的檔案
#              執行 rename()，此為同資料夾內部操作，不受本條限制。
# DEPENDENCIES: 內建標準庫為主。EXIF 功能需額外安裝 `pip install pillow`
# VERSION:     3.1.0 [Stability: Experimental]
# ADR-005:     引入 RuleFactory 動態載入 JSON 建立管線，若未提供 JSON 則退回內建 PIPELINE。
# ADR-006:     引入 RevertFlow 復原機制。在 RenameFlow 實作 try...finally，確保遭逢
#              強制中斷 (Ctrl+C) 時，時光機地圖必定落地，並清除未驗證的殘檔，落實 100% 可回滾。
# ADR-007:     RevertFlow 中的還原採用 OS 原生 rename()，原子性由底層保證，不套用複製刪除。
# ADR-008:     檔案驗證全面升級為 SHA256 雜湊比對；Sanitization 升級同步攔截 Windows 保留字。
# ==========================================================

import os
import shutil
import tkinter as tk
import re
import json
import hashlib
from tkinter import filedialog
from pathlib import Path
from dataclasses import dataclass
from typing import List, Tuple, Set, Protocol, Dict
from datetime import datetime

try:
    from PIL import Image, ExifTags
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

# ==========================================
# 0. 規則積木庫 (Rules Library)
# ==========================================
class RenameRule(Protocol):
    def apply(self, current_name: str, original_path: Path, logger: "Logger | None" = None) -> str: ...

class UnderscoreStripRule:
    def __init__(self, left_strip: int = 0, right_strip: int = 0, sep: str = "_"):
        self.l_strip, self.r_strip, self.sep = left_strip, right_strip, sep
    def apply(self, current_name: str, original_path: Path, logger: "Logger | None" = None) -> str:
        name_part, ext_part = os.path.splitext(current_name)
        parts = name_part.split(self.sep)
        req = self.l_strip + self.r_strip
        if len(parts) <= req or req == 0: return current_name
        return self.sep.join(parts[self.l_strip:len(parts)-self.r_strip]) + ext_part

class RegexRule:
    def __init__(self, pattern: str, replacement: str):
        self.pattern = re.compile(pattern)
        self.replacement = replacement
    def apply(self, current_name: str, original_path: Path, logger: "Logger | None" = None) -> str:
        name_part, ext_part = os.path.splitext(current_name)
        return f"{self.pattern.sub(self.replacement, name_part)}{ext_part}"

class DictReplaceRule:
    def __init__(self, mapping: dict):
        self.mapping = mapping
    def apply(self, current_name: str, original_path: Path, logger: "Logger | None" = None) -> str:
        name_part, ext_part = os.path.splitext(current_name)
        for old_str, new_str in self.mapping.items():
            name_part = name_part.replace(old_str, new_str)
        return f"{name_part}{ext_part}"

class CasingRule:
    def __init__(self, mode: str = "lower"):
        self.mode = mode.lower()
    def apply(self, current_name: str, original_path: Path, logger: "Logger | None" = None) -> str:
        name_part, ext_part = os.path.splitext(current_name)
        if self.mode == "upper": name_part = name_part.upper()
        elif self.mode == "lower": name_part = name_part.lower()
        elif self.mode == "title": name_part = name_part.title()
        return f"{name_part}{ext_part}"

class ZeroPadRule:
    def __init__(self, pad_length: int = 3):
        self.pad_length = pad_length
    def apply(self, current_name: str, original_path: Path, logger: "Logger | None" = None) -> str:
        name_part, ext_part = os.path.splitext(current_name)
        return f"{re.sub(r'\d+', lambda x: x.group().zfill(self.pad_length), name_part)}{ext_part}"

class OSModifiedTimeRule:
    def __init__(self, time_format: str = "%Y%m%d_", prefix: str = ""):
        self.time_format, self.prefix = time_format, prefix
    def apply(self, current_name: str, original_path: Path, logger: "Logger | None" = None) -> str:
        time_str = datetime.fromtimestamp(original_path.stat().st_mtime).strftime(self.time_format)
        return f"{self.prefix}{time_str}{current_name}"

class ExifDateRule:
    def __init__(self, time_format: str = "%Y%m%d_%H%M%S_"):
        self.time_format = time_format
    def apply(self, current_name: str, original_path: Path, logger: "Logger | None" = None) -> str:
        if not HAS_PIL or original_path.suffix.lower() not in ['.jpg', '.jpeg', '.png']: return current_name
        try:
            with Image.open(original_path) as img:
                exif = img._getexif()
                if not exif: return current_name
                for k, v in exif.items():
                    if ExifTags.TAGS.get(k) == "DateTimeOriginal":
                        dt = datetime.strptime(v, "%Y:%m:%d %H:%M:%S")
                        return f"{dt.strftime(self.time_format)}{current_name}"
        except Exception as e:
            msg = f"[WARN] {original_path.name} EXIF 讀取異常: {str(e)}"
            print(msg)
            if logger:
                logger.write(msg)
        return current_name

# ==========================================
# 1. 引擎工廠與內建設定 (Rule Factory)
# ==========================================
DEFAULT_PIPELINE: List[RenameRule] = []

class RuleFactory:
    _REGISTRY = {
        "UnderscoreStripRule": UnderscoreStripRule,
        "RegexRule": RegexRule,
        "DictReplaceRule": DictReplaceRule,
        "CasingRule": CasingRule,
        "ZeroPadRule": ZeroPadRule,
        "OSModifiedTimeRule": OSModifiedTimeRule,
        "ExifDateRule": ExifDateRule
    }

    @classmethod
    def create_pipeline_from_json(cls, json_path: Path) -> List[RenameRule]:
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        pipeline = []
        for item in data.get("pipeline", []):
            rule_name = item.get("rule")
            params = item.get("params", {})
            
            if rule_name is None:
                continue
            if rule_name.startswith("X"):
                print(f"[停用] {rule_name[1:]} 已被使用者以 X 前綴停用。")
                continue
                
            if rule_name in cls._REGISTRY:
                pipeline.append(cls._REGISTRY[rule_name](**params))
            else:
                print(f"[警告] 找不到規則 '{rule_name}'，已略過。")
        return pipeline

# ==========================================
# 2. CONFIG & STRING
# ==========================================
@dataclass
class AppConfig:
    SUCCESS_DIR_NAME: str = "檔案"
    FAIL_DIR_NAME: str = "改名失敗"
    LOG_FILE_NAME: str = "rename_log.txt"
    REVERT_MAP_PREFIX: str = "revert_map_"
    SPEED_MB_PER_SEC: int = 150         

CONFIG = AppConfig()

class NameProcessor:
    RESERVED_NAMES = {"CON", "PRN", "AUX", "NUL", 
                      "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", 
                      "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"}

    @staticmethod
    def sanitize_filename(filename: str) -> str:
        cleaned = re.sub(r'[<>:"/\\|?*]', '_', filename)
        stem, ext = os.path.splitext(cleaned)
        if stem.upper() in NameProcessor.RESERVED_NAMES:
            cleaned = f"_{stem}{ext}"
        return cleaned

    @staticmethod
    def generate_new_name(original_path: Path, pipeline: List[RenameRule], logger: "Logger | None" = None) -> str:
        current_name = original_path.name.replace("\u3000", " ").strip()
        for rule in pipeline:
            current_name = rule.apply(current_name, original_path, logger)
        return NameProcessor.sanitize_filename(current_name)

class FileOps:
    @staticmethod
    def calculate_sha256(file_path: Path) -> str:
        sha256_hash = hashlib.sha256()
        with open(file_path, "rb") as f:
            while chunk := f.read(8192):
                sha256_hash.update(chunk)
        return sha256_hash.hexdigest()

    @staticmethod
    def get_safe_target_path(target_dir: Path, new_name: str, used_names: Set[str]) -> Path:
        name_part, ext_part = os.path.splitext(new_name)
        counter, final_name = 1, new_name
        while (target_dir / final_name).exists() or (final_name in used_names):
            final_name = f"{name_part}_{counter}{ext_part}"
            counter += 1
        used_names.add(final_name)
        return target_dir / final_name

class Logger:
    def __init__(self, log_path: Path):
        self.log_path = log_path
        if not self.log_path.exists():
            with open(self.log_path, 'w', encoding='utf-8') as f: f.write("=== 批次作業日誌 ===\n")
    def write(self, msg: str):
        with open(self.log_path, 'a', encoding='utf-8') as f: f.write(msg + "\n")

# ==========================================
# FLOW 1: Rename Mode (改名模式)
# ==========================================
class RenameFlow:
    def __init__(self):
        self.pipeline = DEFAULT_PIPELINE
        self.source_dir = Path()
        self.target_dir = Path()
        self.file_list: List[Path] = []

    def load_config(self):
        print("\n【步驟 1】請選擇外部 JSON 設定檔 (點擊取消以使用程式內建預設值)")
        cfg_path = filedialog.askopenfilename(title="選擇 JSON 設定檔", filetypes=[("JSON files", "*.json")])
        if cfg_path:
            try:
                self.pipeline = RuleFactory.create_pipeline_from_json(Path(cfg_path))
                print(f"[載入成功] 已套用 {Path(cfg_path).name} 的規則。")
            except Exception as e:
                print(f"[載入失敗] 讀取 JSON 錯誤 ({e})，降級使用內建預設值。")
        else:
            print("[資訊] 未選擇外部設定檔，使用內建預設值。")

    def select_directories(self):
        print("\n【步驟 2】請選擇來源與目的資料夾")
        src = filedialog.askdirectory(title="選擇「想處理區」(來源)")
        if not src: raise SystemExit("[FAIL FAST] 未選擇來源。")
        tgt = filedialog.askdirectory(title="選擇「完成區」(目的)")
        if not tgt: raise SystemExit("[FAIL FAST] 未選擇目的。")

        self.source_dir, self.target_dir = Path(src), Path(tgt)
        if self.source_dir.resolve() == self.target_dir.resolve():
            raise SystemExit("[FAIL FAST] 來源與目的不可相同！")
        
        self.file_list = [f for f in self.source_dir.iterdir() if f.is_file()]

    def execute(self):
        print("\n=== [DRY RUN] 預覽結果 ===")
        success_dir = self.target_dir / CONFIG.SUCCESS_DIR_NAME
        fail_dir = self.target_dir / CONFIG.FAIL_DIR_NAME
        
        logger = Logger(self.target_dir / CONFIG.LOG_FILE_NAME)
        
        used_names: Set[str] = set()
        plan: List[Tuple[Path, str]] = []

        for f in self.file_list:
            new_name = NameProcessor.generate_new_name(f, self.pipeline, logger)
            safe_path = FileOps.get_safe_target_path(success_dir, new_name, used_names)
            plan.append((f, safe_path.name))
            print(f"  {f.name}  ➔  {safe_path.name}")

        if input("\n確定執行複製與改名？(Y/N): ").strip().upper() != 'Y': 
            logger.write("[INFO] 使用者在預覽後取消操作。")
            return

        success_dir.mkdir(parents=True, exist_ok=True)
        fail_dir.mkdir(parents=True, exist_ok=True)
        
        revert_data = {"success_dir": str(success_dir.resolve()), "files": {}}
        success_count, fail_count = 0, 0
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        revert_map_filename = f"{CONFIG.REVERT_MAP_PREFIX}{timestamp}.json"

        try:
            for orig_path, new_name in plan:
                tgt_path = success_dir / new_name
                file_completed = False
                
                try:
                    shutil.copy2(orig_path, tgt_path)
                    
                    if FileOps.calculate_sha256(orig_path) != FileOps.calculate_sha256(tgt_path):
                        tgt_path.unlink(missing_ok=True)
                        raise IOError("驗證失敗：SHA256 雜湊值不符，已清除殘檔。")
                    
                    logger.write(f"[SUCCESS] {orig_path.name} -> {new_name}")
                    revert_data["files"][new_name] = orig_path.name 
                    success_count += 1
                    file_completed = True

                except KeyboardInterrupt:
                    if not file_completed:
                        tgt_path.unlink(missing_ok=True)
                        print(f"\n[中斷防護] 已清除未驗證的殘檔: {tgt_path.name}")
                    raise 

                except Exception as e:
                    logger.write(f"[FAIL] {orig_path.name} | Error: {e}")
                    fail_count += 1
                    try: shutil.copy2(orig_path, fail_dir / orig_path.name)
                    except Exception as ne: print(f"[CRITICAL] 備份失敗: {ne}")
        finally:
            if revert_data["files"]:
                with open(self.target_dir / revert_map_filename, 'w', encoding='utf-8') as f:
                    json.dump(revert_data, f, ensure_ascii=False, indent=2)
                print(f"\n[狀態保存] 已將 {len(revert_data['files'])} 筆時光機紀錄落地至 {revert_map_filename}")

        print(f"\n✅ 執行結束！成功: {success_count}，失敗: {fail_count}。")

# ==========================================
# FLOW 2: Revert Mode (時光機模式)
# ==========================================
class RevertFlow:
    def execute(self):
        print("\n【時光機模式】請選擇您要復原的 revert_map JSON 檔")
        map_path = filedialog.askopenfilename(title="選擇時光機地圖", filetypes=[("JSON files", "*.json")])
        if not map_path: return print("未選擇檔案，終止。")

        with open(map_path, 'r', encoding='utf-8') as f:
            data = json.load(f)

        success_dir = Path(data.get("success_dir", ""))
        files_map: Dict[str, str] = data.get("files", {})

        if not success_dir.exists():
            return print(f"[錯誤] 找不到目標資料夾 {success_dir}")

        logger = Logger(success_dir.parent / CONFIG.LOG_FILE_NAME)
        logger.write(f"\n=== 啟動時光機還原 (依據: {Path(map_path).name}) ===")

        print("\n=== [預覽] 即將還原以下檔案 ===")
        valid_reverts = []
        for new_name, old_name in files_map.items():
            current_path = success_dir / new_name
            if current_path.exists():
                valid_reverts.append((current_path, success_dir / old_name))
                print(f"  {new_name}  ➔  {old_name}")
            else:
                msg = f"  [遺失] 找不到檔案: {new_name}，無法還原。"
                print(msg)
                logger.write(f"[REVERT_SKIP] {new_name} (檔案已不在目標資料夾)")

        if not valid_reverts: return print("\n沒有可以還原的檔案。")
        
        if input("\n確定執行還原？此動作將在原資料夾將檔案重新命名 (Y/N): ").strip().upper() != 'Y': 
            logger.write("[REVERT_CANCEL] 使用者取消還原操作")
            return

        success_count = 0
        for current_path, restore_path in valid_reverts:
            try:
                safe_restore_path = restore_path
                counter = 1
                while safe_restore_path.exists():
                    safe_restore_path = success_dir / f"{restore_path.stem}_revert{counter}{restore_path.suffix}"
                    counter += 1
                
                current_path.rename(safe_restore_path)
                logger.write(f"[REVERT_SUCCESS] {current_path.name} -> {safe_restore_path.name}")
                success_count += 1
            except Exception as e:
                err_msg = f"[REVERT_FAIL] 無法還原 {current_path.name}: {e}"
                print(err_msg)
                logger.write(err_msg)

        print(f"\n✅ 時光機還原完成！共還原 {success_count} 個檔案。")

# ==========================================
# MAIN
# ==========================================
def main():
    root = tk.Tk()
    root.withdraw()
    
    print("========================================")
    print("      Script SafeRenamer v3.1.0")
    print("========================================")
    print("[1] 執行批次改名 (使用 JSON 或內建設定)")
    print("[2] 啟動時光機 (讀取 revert_map 還原檔名)")
    
    choice = input("\n請輸入選項 (1 或 2): ").strip()
    
    try:
        if choice == '1':
            flow = RenameFlow()
            flow.load_config()
            flow.select_directories()
            if flow.file_list: flow.execute()
            else: print("[資訊] 來源沒有檔案。")
        elif choice == '2':
            RevertFlow().execute()
        else:
            print("無效選項，程式結束。")
    except KeyboardInterrupt:
        print("\n\n[中斷] 偵測到強制中斷 (Ctrl+C)，系統已介入處理。")
    except Exception as e:
        print(f"\n[系統錯誤] {e}")
    finally:
        input("\n[結束] 請按 Enter 鍵關閉視窗...")

if __name__ == "__main__":
    main()
