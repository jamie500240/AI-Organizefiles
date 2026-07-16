# ==========================================================
# MODULE:       Script_MediaCollapseMigrator
# PURPOSE:      自動掃描指定資料夾，提取圖片 DNA，將檔案安全移轉至分發目錄，執行貪婪坍塌平攤無意義隧道，並生成溯源對照表
# EXPORTS:      Script_MediaCollapseMigrator
# IMPORTS:      shutil, csv, hashlib, logging, time, pathlib, re, threading, tkinter, tqdm, PIL, numpy, concurrent.futures
# FORBIDDEN:    禁止使用 open('w') 直接覆寫正式報表；禁止使用未經驗證的直接移動（shutil.move）
# DEPENDENCIES: PIL, numpy, tqdm
# VERSION:      1.0.0 [Stability: Experimental]
# ADR: 
# - ADR-005:    放棄 ACDS_ContractRegistry 外部依賴，將 CONFIG 轉為模組內部常數，以支援無相依環境之 Standalone 獨立執行。
# - ADR-006:    導入 safe_move 全面取代 shutil.move，包含例外回滾機制亦須遵守「複製+驗證+刪除」原則；同時引入 try-finally 確保中斷時必產出溯源報表。觸發 Major 版本升級。
# ==========================================================

import shutil
import time
import re
import csv
import hashlib
import logging
import threading
from pathlib import Path
from tkinter import Tk, filedialog
from tqdm import tqdm
from PIL import Image
import numpy as np
from concurrent.futures import ThreadPoolExecutor

# 配置日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- [SSOT 外部參數] ---
CONFIG = {
    "THRESH_BASE": 42, 
    "THRESH_AI": 130,
    "ENTROPY_AI": 6.2, 
    "LV1_DATE_MIN": 2,    
    "MAX_DIR_LEN": 50,
    "MAX_FILE_LEN": 60,
    "NON_IMG_DIR": "[ASSETS]",
    "FAILED_DIR": "[FAILED_TO_PROCESS]"
}

class Script_MediaCollapseMigrator:
    def __init__(self, source_path):
        self.root = Path(source_path).resolve()
        self.timestamp = time.strftime('%H%M%S')
        self.export_root = self.root.parent / f"[MEDIA_VAULT]_{self.timestamp}"
        self.img_exts = {'.jpg', '.jpeg', '.jfif', '.png', '.gif', '.webp', '.heic', '.bmp', '.tiff'}
        
        self.human_tags = set()
        self.tags_lock = threading.Lock()
        
        self.traceability_log = []
        self.log_lock = threading.Lock()

    def _calculate_md5(self, file_path: Path) -> str:
        """計算檔案 MD5 供安全移動比對使用"""
        hash_md5 = hashlib.md5()
        with open(file_path, "rb") as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hash_md5.update(chunk)
        return hash_md5.hexdigest()

    def safe_move(self, src: Path, dest: Path) -> Path:
        """[驗證移動] 複製檔案，比對大小與 MD5 一致後才刪除原檔，徹底遵守 FORBIDDEN 規範"""
        if dest.exists() and dest != src:
            dest = dest.parent / f"{src.stem}_{self.timestamp}{src.suffix}"
        
        shutil.copy2(src, dest)
        
        if dest.stat().st_size == src.stat().st_size and self._calculate_md5(dest) == self._calculate_md5(src):
            src.unlink()
            return dest
        else:
            # 若比對失敗，清理殘缺的目的地檔案，並拋出例外
            if dest.exists():
                dest.unlink()
            raise IOError(f"檔案移轉驗證失敗，原檔保留: {src}")

    def generate_traceability_report(self):
        """[溯源報表] 使用 'a' 模式附加，禁止直接覆寫"""
        report_path = self.export_root / f"Migration_Traceability_{self.timestamp}.csv"
        
        with open(report_path, 'a', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(["Original Path", "New Path", "Is Image", "AI Flag", "Error"])
            for log in self.traceability_log:
                writer.writerow([
                    log.get('src', ''), 
                    log.get('dest', ''), 
                    log.get('is_img', ''), 
                    log.get('ai_flag', ''),
                    log.get('error', '')
                ])

    def apply_greedy_collapse(self, target_root):
        """[貪婪坍塌演算法] 實作原子性回滾 (Atomicity)"""
        for _ in range(5): 
            all_dirs = sorted([d for d in target_root.rglob("*") if d.is_dir()], 
                              key=lambda x: len(x.parts), reverse=True)
            
            for d in all_dirs:
                if not d.exists() or d == target_root or d.name in (CONFIG["NON_IMG_DIR"], CONFIG["FAILED_DIR"]): 
                    continue
                
                content = list(d.iterdir())
                files = [f for f in content if f.is_file()]
                depth = len(d.relative_to(target_root).parts)
                is_human = d.name in self.human_tags and depth == 1

                has_no_files = (len(files) == 0)
                is_lonely_date = False
                if depth == 1 and not is_human:
                    if re.match(r'^\d{4,8}(_.*)?$', d.name):
                        if len(files) < CONFIG["LV1_DATE_MIN"]:
                            is_lonely_date = True

                if has_no_files or is_lonely_date:
                    moved_records = []
                    try:
                        for item in content:
                            dest_path = d.parent / item.name
                            actual_dest = self.safe_move(item, dest_path)
                            moved_records.append((actual_dest, item)) # 記錄 (新路徑, 舊路徑)
                        
                        try: 
                            d.rmdir()
                        except OSError: 
                            pass
                    
                    except Exception as e:
                        logging.error(f"坍塌過程發生例外，啟動目錄回滾機制: {d} | Error: {e}")
                        # 執行回滾：嚴格呼叫 safe_move 進行安全倒退，不再使用 shutil.move
                        for current_loc, original_loc in moved_records:
                            try:
                                if current_loc.exists():
                                    self.safe_move(current_loc, original_loc)
                            except Exception as rollback_err:
                                logging.critical(f"致命錯誤：回滾機制執行失敗，檔案狀態不一致: {current_loc} -> {rollback_err}")
                        # 揭露錯誤並中止該次坍塌操作
                        raise RuntimeError(f"原子性保護觸發，目錄 {d} 已終止處理。") from e

    def get_dna(self, fpath):
        """萃取圖片 DNA，明確記錄與揭露例外狀況"""
        try:
            stat = fpath.stat()
            rel = fpath.relative_to(self.root).parts
            raw_tag = rel[0] if len(rel) > 1 else None
            is_human = raw_tag and not re.match(r'^\d{4,8}$', raw_tag)
            
            if is_human:
                with self.tags_lock:
                    self.human_tags.add(raw_tag)
            
            date = time.strftime("%Y%m%d", time.localtime(stat.st_mtime))
            is_img = fpath.suffix.lower() in self.img_exts
            
            dna = {"file": fpath, "tag": raw_tag, "is_human": is_human, "date": date, "is_img": is_img, "error": None}
            
            if is_img:
                with Image.open(fpath) as img:
                    exif = img.getexif()
                    if exif and 306 in exif: 
                        date = exif[306].split(' ')[0].replace(':', '')
                    
                    gray = img.convert('L')
                    hist, _ = np.histogram(np.array(gray), bins=256, range=(0, 255))
                    p = hist / (hist.sum() + 1e-7)
                    ent = -np.sum(p * np.log2(p + 1e-7))
                    
                    dna.update({"code": "_A" if ent < CONFIG["ENTROPY_AI"] else "", "date": date})
            return dna

        except Exception as e:
            logging.error(f"檔案 DNA 萃取失敗: {fpath} - {str(e)}")
            return {"file": fpath, "is_img": False, "error": str(e), "failed": True}

    def execute(self):
        logging.info("啟動 Script_MediaCollapseMigrator (V1.0.0)。DNA 分析與分發模式運轉中...")
        all_files = [f for f in self.root.rglob("*") if f.is_file()]
        
        with ThreadPoolExecutor() as exc:
            pool = list(tqdm(exc.map(self.get_dna, all_files), total=len(all_files), desc="DNA 採集"))
        
        self.export_root.mkdir(parents=True, exist_ok=True)
        
        # 1. 隔離非圖片與錯誤檔案
        for x in [p for p in pool if not p.get('is_img')]:
            target_dir_name = CONFIG["FAILED_DIR"] if x.get('failed') else CONFIG["NON_IMG_DIR"]
            q_dir = self.export_root / target_dir_name
            q_dir.mkdir(parents=True, exist_ok=True)
            
            dest = q_dir / x['file'].name
            shutil.copy2(x['file'], dest)
            
            with self.log_lock:
                self.traceability_log.append({
                    "src": str(x['file']), 
                    "dest": str(dest), 
                    "is_img": False, 
                    "error": x.get('error', '')
                })

        # 2. 分發與密碼化命名
        umbrella_bins = {}
        for item in [p for p in pool if p.get('is_img') and not p.get('failed')]:
            u_key = item['tag'] if item['is_human'] else item['date']
            umbrella_bins.setdefault(u_key, []).append(item)

        for u_name, items in umbrella_bins.items():
            u_dir = self.export_root / self.sanitize(u_name[:CONFIG["MAX_DIR_LEN"]])
            u_dir.mkdir(parents=True, exist_ok=True)
            for it in items:
                clean = "".join(re.findall(r'[\u4e00-\u9fa5a-zA-Z0-9]+', it['file'].stem))
                max_len = CONFIG["MAX_FILE_LEN"]
                new_name = f"[{it['date']}]{it.get('code','')}{clean[:max_len]}{it['file'].suffix}"
                
                dest = u_dir / new_name
                shutil.copy2(it['file'], dest)
                
                with self.log_lock:
                    self.traceability_log.append({
                        "src": str(it['file']), 
                        "dest": str(dest), 
                        "is_img": True, 
                        "ai_flag": it.get('code', '')
                    })

        # 3. 執行貪婪坍塌 (加上 try-finally 防護)
        try:
            logging.info("分發完成，開始執行目錄坍塌...")
            self.apply_greedy_collapse(self.export_root)
        except Exception as e:
            logging.error(f"貪婪坍塌過程遭遇致命中斷，部分目錄可能未平攤: {e}")
        finally:
            # 4. 生成溯源對照表 (確保必定執行，留下紀錄)
            self.generate_traceability_report()
            logging.info("移轉流程結束。溯源報表已生成，請核對 FAILED_TO_PROCESS 及遷移日誌。")

    def sanitize(self, name):
        return re.sub(r'[\\/:*?"<>|]', '_', str(name)).strip()

if __name__ == "__main__":
    Tk().withdraw()
    p = filedialog.askdirectory(title="選擇要進行坍塌重構的來源資料夾")
    if p: 
        Script_MediaCollapseMigrator(p).execute()
    input("\n[程序終點] 按下 Enter 結束並檢核最終成果。")
