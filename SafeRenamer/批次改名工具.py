# ==========================================================
# MODULE:      Script_SafeRenamer
# PURPOSE:     安全批次改名工具：支援乾跑預演、字串切除、防撞名，採「先複製後驗證」非破壞性處理
# EXPORTS:     RenameFlow, NameProcessor, FileOps, Logger
# IMPORTS:     os, shutil, tkinter, pathlib, dataclasses, typing
# FORBIDDEN:   禁止直接對原始檔案執行 move 或 rename；禁止靜默吞沒例外 (except: pass)
# DEPENDENCIES: 內建標準庫 (無第三方套件依賴)
# VERSION:      1.0.0 [Stability: Frozen]
# ADR-001:     檔案佔用與鎖定簡化處理。遇到佔用或權限錯誤，直接記錄失敗並轉入失敗區，不實作暫存區排隊(Queue)機制。
# ADR-002:     原子性寫入原則 (All-or-Nothing)。若寫入成功區的檔案未能通過後續的大體驗證，必須在拋出例外前將其清除 (unlink)，防止髒資料污染成功區。
# ==========================================================

import os
import shutil
import tkinter as tk
from tkinter import filedialog
from pathlib import Path
from dataclasses import dataclass
from typing import List, Tuple, Set

# ==========================================
# 1. SET (開關與邏輯設定) - SSOT
# ==========================================
@dataclass
class RenameSettings:
    LEFT_STRIP_ENABLED: bool = True     
    LEFT_STRIP_INDEX: int = 1           

    RIGHT_STRIP_ENABLED: bool = True    
    RIGHT_STRIP_INDEX: int = 1          

    SEPARATOR: str = "_"                

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
    ERR_SETTING_IDX: str = "[FAIL FAST] INDEX 設定必須大於等於 1。"
    INFO_DRY_RUN: str = "\n=== [DRY RUN] 預覽結果 ==="
    INFO_EXECUTE: str = "\n=== 開始執行複製與改名 ==="
    PROMPT_CONTINUE: str = "\n請問是否要執行上述變更？(Y/N): "

SET = RenameSettings()
CONFIG = AppConfig()
STR = AppStrings()

# ==========================================
# ACTION
# ==========================================
class NameProcessor:
    """負責檔名字串運算與正規化"""
    
    @staticmethod
    def normalize_filename(filename: str) -> str:
        """P3: 使用 Unicode \u3000 替換全形空白，避免 IDE 格式化陷阱，並剔除前後不可見字元"""
        return filename.replace("\u3000", " ").strip()

    @staticmethod
    def generate_new_name(original_filename: str) -> str:
        clean_filename = NameProcessor.normalize_filename(original_filename)
        name_part, ext_part = os.path.splitext(clean_filename)
        parts = name_part.split(SET.SEPARATOR)
        
        required_underscores = 0
        if SET.LEFT_STRIP_ENABLED: required_underscores += SET.LEFT_STRIP_INDEX
        if SET.RIGHT_STRIP_ENABLED: required_underscores += SET.RIGHT_STRIP_INDEX
        
        if len(parts) <= required_underscores:
            return clean_filename
        
        start_idx = SET.LEFT_STRIP_INDEX if SET.LEFT_STRIP_ENABLED else 0
        end_idx = len(parts) - (SET.RIGHT_STRIP_INDEX if SET.RIGHT_STRIP_ENABLED else 0)
        
        new_name_part = SET.SEPARATOR.join(parts[start_idx:end_idx])
        return f"{new_name_part}{ext_part}"

class FileOps:
    """負責檔案系統操作與防撞名計算"""
    
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
    """負責日誌寫入"""
    
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
    """協調來源、目的、預覽與執行流程"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw() 
        self.source_dir = Path()
        self.target_dir = Path()
        self.success_dir = Path()
        self.fail_dir = Path()
        self.file_list: List[Path] = []
        
    def validate_settings(self):
        if SET.LEFT_STRIP_ENABLED and SET.LEFT_STRIP_INDEX < 1:
            raise ValueError(STR.ERR_SETTING_IDX)
        if SET.RIGHT_STRIP_ENABLED and SET.RIGHT_STRIP_INDEX < 1:
            raise ValueError(STR.ERR_SETTING_IDX)

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
            
            base_new_name = NameProcessor.generate_new_name(file_path.name)
            safe_path = FileOps.get_safe_target_path(self.success_dir, base_new_name, used_names)
            
            plan.append((file_path, safe_path.name))
            
            status = "[變更]" if file_path.name != safe_path.name else "[原封]"
            print(f"{status} {file_path.name}  ➔  {safe_path.name}")

        total_mb = total_size_bytes / (1024 * 1024)
        est_seconds = max(1, int(total_mb / CONFIG.SPEED_MB_PER_SEC))
        
        print(f"\n[統計] 共 {len(self.file_list)} 個檔案，總大小約 {total_mb:.2f} MB。")
        print(f"[預估] 以 {CONFIG.SPEED_MB_PER_SEC} MB/s 計算，複製大約需要 {est_seconds} 秒。")
        
        return plan

    def execute(self, plan: List[Tuple[Path, str]]):
        print(STR.INFO_EXECUTE)
        self.success_dir.mkdir(parents=True, exist_ok=True)
        self.fail_dir.mkdir(parents=True, exist_ok=True)
        
        logger = Logger(self.target_dir / CONFIG.LOG_FILE_NAME)
        
        success_count = 0
        fail_count = 0

        for original_path, new_name in plan:
            target_file_path = self.success_dir / new_name
            try:
                # 複製檔案
                shutil.copy2(original_path, target_file_path)
                
                # P2: 強制驗證檔案大小，若失敗則清空目標檔案以落實原子性
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
                    # 將發生錯誤的原檔備份至失敗區
                    shutil.copy2(original_path, self.fail_dir / original_path.name)
                except Exception as nested_e:
                    # P1: 若連寫入失敗區都異常，必須大聲報警
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
        flow.validate_settings()
        flow.select_directories()
        
        if not flow.file_list:
            print("[資訊] 來源資料夾中沒有任何檔案。")
            return

        execution_plan = flow.dry_run()

        user_input = input(STR.PROMPT_CONTINUE).strip().upper()
        if user_input == 'Y':
            flow.execute(execution_plan)
        else:
            print("\n[終止] 使用者取消操作，未進行任何變更。")

    except ValueError as ve:
        print(ve)
    except SystemExit as se:
        print(se)
    except Exception as e:
        print(f"\n[未預期錯誤] {str(e)}")

if __name__ == "__main__":
    main()
