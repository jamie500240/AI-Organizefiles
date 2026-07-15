# ==========================================================
# MODULE:      Script_SafeRenamer
# PURPOSE:     安全批次改名工具：支援正則、EXIF、大小寫、補零與字串切除等多重管線加工
# EXPORTS:     RenameFlow, NameProcessor, FileOps, Logger
# IMPORTS:     os, shutil, tkinter, pathlib, dataclasses, typing, re, datetime
# FORBIDDEN:   禁止直接對原始檔案執行 move 或 rename；禁止靜默吞沒例外 (except: pass)
# DEPENDENCIES: 內建標準庫為主。EXIF 功能需額外安裝 `pip install pillow` (若無則自動跳過)
# VERSION:     2.0.0 [Stability: Frozen]
# ADR-001:     檔案佔用與鎖定簡化處理 (遇到佔用直接轉入失敗區)。
# ADR-002:     原子性寫入原則 (All-or-Nothing)，驗證失敗強制清除殘檔。
# ADR-003:     導入 Pipeline (責任鏈) 模式，將各種改名邏輯解耦為獨立 Rule 積木。
# ADR-004:     輸出合法性防護。所有經過 Pipeline 加工的檔名，最終必須強制濾除 Windows 非法字元。
# ==========================================================

import os
import shutil
import tkinter as tk
import re
from tkinter import filedialog
from pathlib import Path
from dataclasses import dataclass
from typing import List, Tuple, Set, Protocol
from datetime import datetime

# 嘗試載入 Pillow 處理 EXIF，若使用者未安裝則溫柔降級 (Graceful Degradation)
try:
    from PIL import Image, ExifTags
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

# ==========================================
# 0. 規則積木庫 (Rules Library)
# ==========================================
class RenameRule(Protocol):
    def apply(self, current_name: str, original_path: Path) -> str: ...

class UnderscoreStripRule:
    def __init__(self, left_strip: int = 0, right_strip: int = 0, sep: str = "_"):
        self.l_strip = left_strip
        self.r_strip = right_strip
        self.sep = sep
        
    def apply(self, current_name: str, original_path: Path) -> str:
        name_part, ext_part = os.path.splitext(current_name)
        parts = name_part.split(self.sep)
        req = self.l_strip + self.r_strip
        if len(parts) <= req or req == 0:
            return current_name
        
        start_idx = self.l_strip
        end_idx = len(parts) - self.r_strip
        return self.sep.join(parts[start_idx:end_idx]) + ext_part

class RegexRule:
    def __init__(self, pattern: str, replacement: str):
        self.pattern = re.compile(pattern)
        self.replacement = replacement
        
    def apply(self, current_name: str, original_path: Path) -> str:
        name_part, ext_part = os.path.splitext(current_name)
        new_name = self.pattern.sub(self.replacement, name_part)
        return f"{new_name}{ext_part}"

class DictReplaceRule:
    def __init__(self, mapping: dict):
        self.mapping = mapping
        
    def apply(self, current_name: str, original_path: Path) -> str:
        name_part, ext_part = os.path.splitext(current_name)
        for old_str, new_str in self.mapping.items():
            name_part = name_part.replace(old_str, new_str)
        return f"{name_part}{ext_part}"

class CasingRule:
    def __init__(self, mode: str = "lower"):
        self.mode = mode.lower()
        
    def apply(self, current_name: str, original_path: Path) -> str:
        name_part, ext_part = os.path.splitext(current_name)
        if self.mode == "upper": name_part = name_part.upper()
        elif self.mode == "lower": name_part = name_part.lower()
        elif self.mode == "title": name_part = name_part.title()
        return f"{name_part}{ext_part}"

class ZeroPadRule:
    def __init__(self, pad_length: int = 3):
        self.pad_length = pad_length
        
    def apply(self, current_name: str, original_path: Path) -> str:
        name_part, ext_part = os.path.splitext(current_name)
        new_name = re.sub(r'\d+', lambda x: x.group().zfill(self.pad_length), name_part)
        return f"{new_name}{ext_part}"

class OSModifiedTimeRule:
    def __init__(self, time_format: str = "%Y%m%d_", prefix: str = ""):
        self.time_format = time_format
        self.prefix = prefix
        
    def apply(self, current_name: str, original_path: Path) -> str:
        mtime = original_path.stat().st_mtime
        time_str = datetime.fromtimestamp(mtime).strftime(self.time_format)
        return f"{self.prefix}{time_str}{current_name}"

class ExifDateRule:
    def __init__(self, time_format: str = "%Y%m%d_%H%M%S_"):
        self.time_format = time_format
        
    def apply(self, current_name: str, original_path: Path) -> str:
        if not HAS_PIL or original_path.suffix.lower() not in ['.jpg', '.jpeg', '.png']:
            return current_name
            
        try:
            with Image.open(original_path) as img:
                exif = img._getexif()
                if not exif: return current_name
                
                for k, v in exif.items():
                    if ExifTags.TAGS.get(k) == "DateTimeOriginal":
                        dt = datetime.strptime(v, "%Y:%m:%d %H:%M:%S")
                        return f"{dt.strftime(self.time_format)}{current_name}"
        except Exception as e:
            # P2: 不再靜默吞沒例外，留下警告痕跡，但不中斷 Pipeline
            print(f"[警告] {original_path.name} 圖片解析或 EXIF 讀取異常: {str(e)}")
            
        return current_name

# ==========================================
# 1. SET (開關與邏輯設定) - SSOT
# ==========================================
PIPELINE: List[RenameRule] = [
    # UnderscoreStripRule(left_strip=1, right_strip=1),
    # RegexRule(pattern=r'(\d{4})-(\d{2})-(\d{2})', replacement=r'\1\2\3'),
    # DictReplaceRule(mapping={"Draft": "v1", "Final": "v2", " ": "_"}),
    # ZeroPadRule(pad_length=3),
    # CasingRule(mode="lower"),
    # ExifDateRule(),
    # OSModifiedTimeRule(time_format="%Y%m%d_")
]

# ==========================================
# 2. CONFIG (系統配置) - SSOT
# ==========================================
@dataclass
class AppConfig:
    SUCCESS_DIR_NAME: str = "檔案"
    FAIL_DIR_NAME: str = "改名失敗"
    LOG_FILE_NAME: str = "rename_log.txt"
    SPEED_MB_PER_SEC: int = 150         

# ==========================================
# 3. STRING (介面與提示字串) - SSOT
# ==========================================
@dataclass
class AppStrings:
    SELECT_SOURCE: str = "【步驟 1】請選擇「想處理區」(來源資料夾)"
    SELECT_TARGET: str = "【步驟 2】請選擇「完成區」(目的資料夾)"
    ERR_SAME_DIR: str = "[FAIL FAST] 來源與目的資料夾不能相同！"
    ERR_NO_DIR: str = "[FAIL FAST] 未選擇資料夾，程式終止。"
    INFO_DRY_RUN: str = "\n=== [DRY RUN] 預覽結果 ==="
    INFO_EXECUTE: str = "\n=== 開始執行複製與改名 ==="
    PROMPT_CONTINUE: str = "\n請問是否要執行上述變更？(Y/N): "

CONFIG = AppConfig()
STR = AppStrings()

# ==========================================
# ACTION
# ==========================================
class NameProcessor:
    @staticmethod
    def normalize_filename(filename: str) -> str:
        return filename.replace("\u3000", " ").strip()

    @staticmethod
    def sanitize_filename(filename: str) -> str:
        """P2: 強制濾除 Windows 檔案系統不允許的非法字元，替換為底線"""
        return re.sub(r'[<>:"/\\|?*]', '_', filename)

    @staticmethod
    def generate_new_name(original_path: Path) -> str:
        current_name = NameProcessor.normalize_filename(original_path.name)
        
        for rule in PIPELINE:
            current_name = rule.apply(current_name, original_path)
            
        # Pipeline 加工完畢後，進行最終合法性檢查
        return NameProcessor.sanitize_filename(current_name)

class FileOps:
    @staticmethod
    def get_safe_target_path(target_dir: Path, new_name: str, used_names_in_run: Set[str]) -> Path:
        name_part, ext_part = os.path.splitext(new_name)
        counter = 1
        final_name = new_name
        
        while (target_dir / final_name).exists() or (final_name in used_names_in_run):
            final_name = f"{name_part}_{counter}{ext_part}"
            counter += 1
            
        used_names_in_run.add(final_name)
        return target_dir / final_name

class Logger:
    def __init__(self, log_path: Path):
        self.log_path = log_path
        with open(self.log_path, 'w', encoding='utf-8') as f:
            f.write("=== 批次改名日誌 ===\n")
            
    def write(self, msg: str):
        with open(self.log_path, 'a', encoding='utf-8') as f:
            f.write(msg + "\n")

# ==========================================
# FLOW
# ==========================================
class RenameFlow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw() 
        self.source_dir = Path()
        self.target_dir = Path()
        self.success_dir = Path()
        self.fail_dir = Path()
        self.file_list: List[Path] = []

    def select_directories(self):
        print(STR.SELECT_SOURCE)
        src = filedialog.askdirectory(title=STR.SELECT_SOURCE)
        if not src: raise SystemExit(STR.ERR_NO_DIR)
        
        print(STR.SELECT_TARGET)
        tgt = filedialog.askdirectory(title=STR.SELECT_TARGET)
        if not tgt: raise SystemExit(STR.ERR_NO_DIR)

        self.source_dir = Path(src)
        self.target_dir = Path(tgt)

        if self.source_dir.resolve() == self.target_dir.resolve():
            raise SystemExit(STR.ERR_SAME_DIR)

        self.success_dir = self.target_dir / CONFIG.SUCCESS_DIR_NAME
        self.fail_dir = self.target_dir / CONFIG.FAIL_DIR_NAME
        
        self.file_list = [f for f in self.source_dir.iterdir() if f.is_file()]

    def dry_run(self) -> List[Tuple[Path, str]]:
        print(STR.INFO_DRY_RUN)
        total_size_bytes = 0
        used_names: Set[str] = set()
        plan: List[Tuple[Path, str]] = [] 

        for file_path in self.file_list:
            total_size_bytes += file_path.stat().st_size
            
            base_new_name = NameProcessor.generate_new_name(file_path)
            safe_path = FileOps.get_safe_target_path(self.success_dir, base_new_name, used_names)
            
            plan.append((file_path, safe_path.name))
            
            status = "[變更]" if file_path.name != safe_path.name else "[原封]"
            print(f"{status} {file_path.name}  ➔  {safe_path.name}")

        total_mb = total_size_bytes / (1024 * 1024)
        est_seconds = max(1, int(total_mb / CONFIG.SPEED_MB_PER_SEC))
        
        print(f"\n[統計] 共 {len(self.file_list)} 個檔案，總大小約 {total_mb:.2f} MB。")
        print(f"[預估] 以 {CONFIG.SPEED_MB_PER_SEC} MB/s 計算，複製大約需要 {est_seconds} 秒。")
        if not HAS_PIL:
            print("[提醒] 系統未安裝 Pillow 模組，EXIF 讀取功能將自動跳過 (可使用 pip install pillow 安裝)。")
        
        return plan

    def execute(self, plan: List[Tuple[Path, str]]):
        print(STR.INFO_EXECUTE)
        self.success_dir.mkdir(parents=True, exist_ok=True)
        self.fail_dir.mkdir(parents=True, exist_ok=True)
        
        logger = Logger(self.target_dir / CONFIG.LOG_FILE_NAME)
        success_count, fail_count = 0, 0

        for original_path, new_name in plan:
            target_file_path = self.success_dir / new_name
            try:
                shutil.copy2(original_path, target_file_path)
                
                if original_path.stat().st_size != target_file_path.stat().st_size:
                    error_msg = f"檔案驗證失敗：來源 ({original_path.stat().st_size} bytes) 與目的大小不符。"
                    try:
                        target_file_path.unlink(missing_ok=True)
                        error_msg += "已成功清除殘檔。"
                    except Exception as unlink_e:
                        error_msg += f"【嚴重警告】清除殘檔失敗 ({str(unlink_e)})，成功區可能殘留不完整的髒檔案！"
                    raise IOError(error_msg)

                logger.write(f"[SUCCESS] {original_path.name} -> {new_name}")
                success_count += 1

            except Exception as e:
                logger.write(f"[FAIL] {original_path.name} | Error: {str(e)}")
                fail_count += 1
                try:
                    shutil.copy2(original_path, self.fail_dir / original_path.name)
                except Exception as nested_e:
                    critical_err = f"[CRITICAL] 檔案 {original_path.name} 寫入失敗區發生嚴重異常: {str(nested_e)}"
                    print(critical_err)
                    logger.write(critical_err)

        print(f"\n✅ 執行完畢！成功: {success_count}，失敗: {fail_count}。")
        print(f"📄 請至 {self.target_dir.resolve()} 查看完成結果與 Log。")

# ==========================================
# MAIN
# ==========================================
def main():
    flow = RenameFlow()
    try:
        flow.select_directories()
        
        if not flow.file_list:
            print("[資訊] 來源資料夾中沒有任何檔案。")
            return

        execution_plan = flow.dry_run()
        
        if len(PIPELINE) == 0:
            print("\n[警告] PIPELINE 內無任何啟用的規則，所有檔案都將原封不動複製。")

        user_input = input(STR.PROMPT_CONTINUE).strip().upper()
        if user_input == 'Y':
            flow.execute(execution_plan)
        else:
            print("\n[終止] 使用者取消操作，未進行任何變更。")

    except ValueError as ve: 
        print(f"\n{ve}")
    except SystemExit as se: 
        print(f"\n{se}")
    except Exception as e: 
        print(f"\n[未預期錯誤] {str(e)}")
    finally:
        # 確保 CMD 視窗在執行結束或報錯時都會停駐
        input("\n[結束] 請按 Enter 鍵關閉視窗...")

if __name__ == "__main__":
    main()
