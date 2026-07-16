# ==========================================================
# MODULE:       Script_GodCleanup
# PURPOSE:      檔案清理與去重自動化核心 (嚴格遵守 ACDS 合約)
# EXPORTS:      Script_AppRouter.run()
# IMPORTS:      os, shutil, hashlib, pathlib, datetime, tkinter, csv
# FORBIDDEN:    上帝物件、跨層依賴、隱性轉型、靜默覆寫、未授權命名、靜默失敗
# DEPENDENCIES: 作業系統檔案 I/O、Tkinter UI 環境
# VERSION:      2.1.4 [Stability: Stable]
#
# [ADR-001]     關於 P0-11「全有或全無 (All-or-Nothing)」原則之豁免與取捨
# Context:      本系統處理之目標可能高達數百 GB，搬運與雜湊計算耗時極長。
# Decision:     遭遇 Ctrl+C (KeyboardInterrupt) 時，不執行「全數退回 (Rollback)」，而是保留已處理之檔案並產出結算報表。
# Rationale:    在巨量檔案處理情境下，銷毀數小時的成功進度會造成極差的 UX。保留 Partial State 並透過 CSV 報表確保資料狀態具備完全的可稽核性，為此情境下之最佳實務。
# ==========================================================

import os
import shutil
import hashlib
import time
import re
import csv
import traceback
from pathlib import Path
from datetime import datetime
from dataclasses import dataclass, field
import tkinter as tk
from tkinter import filedialog
from typing import List, Dict, Set

# =============================================================================
# [SET] 領域層：全域系統設定與狀態容器 (Domain Models & Config)
# =============================================================================
@dataclass(frozen=True)
class cls_AppConfig:
    BUFFER_SIZE: int = 1024 * 1024
    HASH_LENGTH: int = 16
    ESTIMATE_MB_PER_SEC: float = 50.0
    SAFE_SPACE_RATIO: float = 1.1
    LOCK_RETRY_LIMIT: int = 3  # I/O 驗證重試上限

    DIR_TEMP: str = "0_Temp_Sandbox"
    DIR_UNIQUE: str = "1_Unique_Archive"
    DIR_DUPLICATE: str = "2_Duplicates_Iso"
    DIR_FAILED: str = "3_Failed_Items"
    REPORT_FILENAME: str = "Process_Report.csv"
    
    ENABLE_FILENAME_RESTORE: bool = True
    ENABLE_DIR_FLATTENING: bool = True
    MIN_ITEMS_FOR_DIR: int = 2

@dataclass
class cls_AppState:
    """業務狀態聚合根：用以在互不相識的 Script 模組之間傳遞資料"""
    registry: Set[str] = field(default_factory=set)
    report_data: List[list] = field(default_factory=list)
    stats: Dict[str, int] = field(default_factory=lambda: {"total": 0, "unique": 0, "dup": 0, "fail": 0})

class cls_Strings:
    TITLE_SRC = "【步驟 1/2】請選擇來源資料夾 (唯讀)"
    TITLE_DST = "【步驟 2/2】請選擇輸出資料夾 (將在此建立整理結果)"
    ERR_NO_PATH = "未提供有效路徑，程式中斷。"
    WARN_SPACE = "空間不足警告：目標磁碟至少需要 {req:.2f} MB"
    REPORT_DRY_RUN = (
        "\n--- [DRY RUN 掃描報告] ---\n"
        "來源: {src}\n"
        "輸出: {dst}\n"
        "總數: {count} 個\n"
        "容量: {size:.2f} MB\n"
        "剩餘: {free:.2f} MB\n"
        "耗時: 約 {time:.1f} 分鐘\n"
        "--------------------------"
    )

# =============================================================================
# [UTILS] 基礎設施層：絕對隔離的 I/O 操作 (Infrastructure)
# =============================================================================
class Utils_Logger:
    @staticmethod
    def log(level: str, module: str, message: str):
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"[{level}] {now} {module} - {message}")

class Utils_FileSystem:
    @staticmethod
    def exists(path: Path) -> bool:
        return path.exists()

    @staticmethod
    def make_dir(path: Path):
        path.mkdir(parents=True, exist_ok=True)

    @staticmethod
    def safe_copy(src: Path, dst: Path) -> bool:
        try:
            shutil.copy2(src, dst)
            return True
        except Exception:
            return False

    @staticmethod
    def safe_move(src: Path, dst: Path) -> bool:
        try:
            shutil.move(str(src), str(dst))
            return True
        except Exception:
            return False

    @staticmethod
    def safe_rename(src: Path, dst: Path) -> bool:
        try:
            src.rename(dst)
            return True
        except Exception:
            return False

    @staticmethod
    def safe_unlink(path: Path):
        try:
            if path.exists():
                path.unlink()
        except Exception:
            pass

    @staticmethod
    def safe_rmdir(path: Path) -> bool:
        try:
            path.rmdir()
            return True
        except Exception:
            return False

    @staticmethod
    def safe_rmtree(path: Path):
        try:
            if path.exists():
                shutil.rmtree(path)
        except Exception:
            pass

    @staticmethod
    def calculate_sha256(path: Path, buffer_size: int) -> str:
        h = hashlib.sha256()
        try:
            with open(path, "rb") as f:
                while chunk := f.read(buffer_size):
                    h.update(chunk)
            return h.hexdigest()
        except Exception:
            return ""

    @staticmethod
    def get_mtime_string(path: Path) -> str:
        try:
            return datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y%m%d")
        except Exception:
            return "UnknownDate"

    @staticmethod
    def write_csv(path: Path, data: List[list], header: List[str]) -> bool:
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(header)
                writer.writerows(data)
            return True
        except Exception:
            return False

    @staticmethod
    def get_all_files(dir_path: Path) -> List[Path]:
        return [f for f in dir_path.rglob('*') if f.is_file()]

    @staticmethod
    def walk_bottom_up(dir_path: Path):
        return os.walk(dir_path, topdown=False)

    @staticmethod
    def get_disk_free_mb(dir_path: Path) -> float:
        return shutil.disk_usage(dir_path).free / (1024 * 1024)

    @staticmethod
    def sanitize_filename(name: str) -> str:
        return re.sub(r'[\/:*?"<>|]', '_', name)

# =============================================================================
# [SCRIPT] 邏輯層：單一職責的業務管線 (Business Logic Pipelines)
# =============================================================================
class Script_DryRun:
    def __init__(self, config: cls_AppConfig):
        self.config = config

    def evaluate(self, src: Path, dst: Path) -> dict:
        files = Utils_FileSystem.get_all_files(src)
        total_size_mb = sum(f.stat().st_size for f in files) / (1024 * 1024)
        req_mb = total_size_mb * self.config.SAFE_SPACE_RATIO
        free_mb = Utils_FileSystem.get_disk_free_mb(dst)
        est_time = total_size_mb / self.config.ESTIMATE_MB_PER_SEC / 60
        
        return {
            "files": files,
            "count": len(files),
            "size_mb": total_size_mb,
            "req_mb": req_mb,
            "free_mb": free_mb,
            "est_time_min": est_time,
            "is_safe": free_mb >= req_mb
        }

class Script_Archiver:
    def __init__(self, config: cls_AppConfig, state: cls_AppState, src_dir: Path, dst_dir: Path):
        self.config = config
        self.state = state
        self.src_dir = src_dir
        self.paths = {
            "temp": dst_dir / config.DIR_TEMP,
            "unique": dst_dir / config.DIR_UNIQUE,
            "dup": dst_dir / config.DIR_DUPLICATE,
            "failed": dst_dir / config.DIR_FAILED
        }
        for p in self.paths.values():
            Utils_FileSystem.make_dir(p)

    def execute(self, file_list: List[Path]):
        Utils_Logger.log("INFO", "Script_Archiver", "啟動 Phase 1: 檔案歸檔與雜湊分類")
        for idx, src_file in enumerate(file_list, 1):
            self.state.stats["total"] += 1
            temp_file = self.paths["temp"] / src_file.name
            
            # [P0-2] 1. 先取得真實的來源 DNA 作為基準
            src_hash = Utils_FileSystem.calculate_sha256(src_file, self.config.BUFFER_SIZE)
            if not src_hash:
                Utils_Logger.log("ERROR", "Script_Archiver", f"來源 Hash 計算失敗，檔案可能被鎖定: {src_file.name}")
                self._record_fail(src_file, "SRC_HASH_FAILED")
                continue

            # [P0-2] 2. 執行「複製與完整性驗證」的防禦性迴圈
            copy_success = False
            for attempt in range(self.config.LOCK_RETRY_LIMIT):
                if not Utils_FileSystem.safe_copy(src_file, temp_file):
                    time.sleep(0.5)
                    continue
                
                # 對複製品進行 DNA 採樣並比對
                temp_hash = Utils_FileSystem.calculate_sha256(temp_file, self.config.BUFFER_SIZE)
                if temp_hash == src_hash:
                    copy_success = True
                    break
                else:
                    Utils_Logger.log("WARN", "Script_Archiver", f"完整性驗證失敗，進行重試 ({attempt+1}/{self.config.LOCK_RETRY_LIMIT}): {src_file.name}")
                    Utils_FileSystem.safe_unlink(temp_file)
                    time.sleep(0.5)

            if not copy_success:
                Utils_Logger.log("ERROR", "Script_Archiver", f"超過重試次數，複製或驗證徹底失敗: {src_file.name}")
                self._record_fail(src_file, "COPY_VERIFY_FAILED")
                continue

            # 3. 判定分類與生成目標目錄
            hash_val = src_hash 
            is_unique = hash_val not in self.state.registry
            base_dir = self.paths["unique"] if is_unique else self.paths["dup"]
            status_tag = "UNIQUE" if is_unique else "DUP"

            try:
                rel_path = src_file.relative_to(self.src_dir)
                target_dir = base_dir / rel_path.parent
                Utils_FileSystem.make_dir(target_dir)
            except Exception:
                target_dir = base_dir

            clean_stem = Utils_FileSystem.sanitize_filename(src_file.stem)[:50]
            short_hash = hash_val[:self.config.HASH_LENGTH]
            marked_name = f"{short_hash}_{clean_stem}_{status_tag}{src_file.suffix}"
            final_dst = self._get_safe_destination(target_dir, marked_name)

            if Utils_FileSystem.safe_move(temp_file, final_dst):
                if is_unique:
                    self.state.registry.add(hash_val)
                    self.state.stats["unique"] += 1
                else:
                    self.state.stats["dup"] += 1
                self.state.report_data.append([final_dst.name, hash_val[:8], status_tag, str(src_file)])
            else:
                Utils_Logger.log("ERROR", "Script_Archiver", f"歸檔移動失敗: {src_file.name}")
                self._record_fail(src_file, "MOVE_FAILED")
                Utils_FileSystem.safe_unlink(temp_file)

    def _get_safe_destination(self, target_dir: Path, base_name: Path) -> Path:
        target = target_dir / base_name
        counter = 1
        while Utils_FileSystem.exists(target):
            target = target_dir / f"{target.stem}_REV{counter}{target.suffix}"
            counter += 1
        return target

    def _record_fail(self, src: Path, reason: str):
        self.state.stats["fail"] += 1
        safe_name = Utils_FileSystem.sanitize_filename(src.name)
        fail_path = self.paths["failed"] / f"ERR_{int(time.time())}_{safe_name}.txt"
        self.state.report_data.append([safe_name, "N/A", f"FAIL: {reason}", str(src)])
        try:
            with open(fail_path, "w", encoding="utf-8") as f:
                f.write(f"Source: {src}\nReason: {reason}")
        except Exception:
            pass

class Script_Renamer:
    def __init__(self, config: cls_AppConfig, unique_dir: Path):
        self.config = config
        self.unique_dir = unique_dir
        self.pattern = re.compile(rf'^([a-fA-F0-9]{{{self.config.HASH_LENGTH}}})_(.*)_UNIQUE(\..*)?$')

    def execute(self):
        Utils_Logger.log("INFO", "Script_Renamer", "啟動 Phase 2a: 印記剝離")
        for root, _, files in Utils_FileSystem.walk_bottom_up(self.unique_dir):
            current_dir = Path(root)
            for file in files:
                old_path = current_dir / file
                match = self.pattern.match(file)
                if not match:
                    continue
                
                base_name = match.group(2)
                ext = match.group(3) if match.group(3) else ""
                new_path = current_dir / f"{base_name}{ext}"

                if Utils_FileSystem.exists(new_path) and new_path != old_path:
                    new_path = self._resolve_collision(current_dir, old_path, base_name, ext)

                if not Utils_FileSystem.safe_rename(old_path, new_path):
                    Utils_Logger.log("WARN", "Script_Renamer", f"印記剝離重新命名失敗: {old_path.name}")

    def _resolve_collision(self, dir_path: Path, src_path: Path, base_name: str, ext: str) -> Path:
        mtime_str = Utils_FileSystem.get_mtime_string(src_path)
        counter = 1
        while True:
            candidate = dir_path / f"{base_name}_{mtime_str}_{counter:02d}{ext}"
            if not Utils_FileSystem.exists(candidate):
                return candidate
            counter += 1

class Script_Flattener:
    def __init__(self, config: cls_AppConfig, unique_dir: Path):
        self.config = config
        self.unique_dir = unique_dir

    def execute(self):
        Utils_Logger.log("INFO", "Script_Flattener", "啟動 Phase 2b: 目錄隧道坍塌")
        for root, dirs, files in Utils_FileSystem.walk_bottom_up(self.unique_dir):
            current_dir = Path(root)
            if current_dir == self.unique_dir:
                continue

            items = list(current_dir.iterdir())
            if len(items) < self.config.MIN_ITEMS_FOR_DIR:
                parent_dir = current_dir.parent
                for item in items:
                    target_path = parent_dir / item.name
                    if Utils_FileSystem.exists(target_path):
                        target_path = self._resolve_collision(parent_dir, item)
                    
                    if not Utils_FileSystem.safe_move(item, target_path):
                        Utils_Logger.log("WARN", "Script_Flattener", f"隧道搬遷失敗: {item.name}")
                
                if not Utils_FileSystem.safe_rmdir(current_dir):
                    Utils_Logger.log("WARN", "Script_Flattener", f"空目錄拆除失敗: {current_dir.name}")

    def _resolve_collision(self, parent_dir: Path, item: Path) -> Path:
        base = item.stem
        ext = item.suffix
        mtime_str = Utils_FileSystem.get_mtime_string(item)
        counter = 1
        while True:
            candidate = parent_dir / f"{base}_{mtime_str}_{counter:02d}{ext}"
            if not Utils_FileSystem.exists(candidate):
                return candidate
            counter += 1

class Script_Reporter:
    @staticmethod
    def finalize(config: cls_AppConfig, state: cls_AppState, dst_dir: Path):
        temp_dir = dst_dir / config.DIR_TEMP
        Utils_FileSystem.safe_rmtree(temp_dir)
        
        report_path = dst_dir / config.REPORT_FILENAME
        headers = ["處理後名稱", "SHA256_短碼", "分類狀態", "原始路徑"]
        if Utils_FileSystem.write_csv(report_path, state.report_data, headers):
            Utils_Logger.log("INFO", "Script_Reporter", f"報告已生成: {report_path}")
        else:
            Utils_Logger.log("ERROR", "Script_Reporter", "報告寫入失敗。")

# =============================================================================
# [ROUTER] 路由層：控制器與應用程式進入點 (Controller)
# =============================================================================
class Script_AppRouter:
    def __init__(self):
        self.config = cls_AppConfig()
        self.state = cls_AppState()

    @staticmethod
    def _get_ui_path(title: str) -> Path:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        folder = filedialog.askdirectory(title=title)
        root.destroy()
        return Path(folder) if folder else None

    def run(self):
        Utils_Logger.log("INFO", "AppRouter", "系統啟動")
        try:
            # 1. Input
            src_path = self._get_ui_path(cls_Strings.TITLE_SRC)
            dst_path = self._get_ui_path(cls_Strings.TITLE_DST)
            if not src_path or not dst_path:
                Utils_Logger.log("ERROR", "AppRouter", cls_Strings.ERR_NO_PATH)
                return

            if src_path in dst_path.parents or src_path == dst_path:
                Utils_Logger.log("CRIT", "AppRouter", "迴圈防禦：輸出不可位於來源之內。")
                return

            # 2. Evaluation
            evaluator = Script_DryRun(self.config)
            assessment = evaluator.evaluate(src_path, dst_path)
            
            print(cls_Strings.REPORT_DRY_RUN.format(
                src=src_path, dst=dst_path, count=assessment["count"], 
                size=assessment["size_mb"], free=assessment["free_mb"], time=assessment["est_time_min"]
            ))

            if not assessment["is_safe"]:
                Utils_Logger.log("CRIT", "AppRouter", cls_Strings.WARN_SPACE.format(req=assessment["req_mb"]))
                return

            if input("\n確認執行？ (輸入 Y 繼續): ").strip().upper() != "Y":
                Utils_Logger.log("INFO", "AppRouter", "使用者取消。")
                return

            # 3. Execution
            archiver = Script_Archiver(self.config, self.state, src_path, dst_path)
            
            try:
                archiver.execute(assessment["files"])
                
                unique_dir = dst_path / self.config.DIR_UNIQUE
                if self.config.ENABLE_FILENAME_RESTORE:
                    Script_Renamer(self.config, unique_dir).execute()
                    
                if self.config.ENABLE_DIR_FLATTENING:
                    Script_Flattener(self.config, unique_dir).execute()

            except KeyboardInterrupt:
                Utils_Logger.log("CRIT", "AppRouter", "偵測到中斷訊號 (Ctrl+C)。立即停止 (參閱 ADR-001 取代 Rollback)。")
            finally:
                # 4. Finalization
                Script_Reporter.finalize(self.config, self.state, dst_path)
                Utils_Logger.log("INFO", "AppRouter", "作業結束。")
                input("\n請按 Enter 鍵關閉視窗...")

        except Exception as e:
            Utils_Logger.log("CRIT", "AppRouter", f"嚴重錯誤:\n{traceback.format_exc()}")
            input("\n請按 Enter 鍵關閉視窗...")

# =============================================================================
# Bootstrap
# =============================================================================
if __name__ == "__main__":
    app = Script_AppRouter()
    app.run()
