# ==========================================================
# MODULE: AegisArchiver
# PURPOSE: 自動掃描指定資料夾，執行資安檢測。安全檔案予以去毒引渡；風險檔案移至隔離區，並於安全區留下溯源佔位符。
# EXPORTS: AegisArchiver
# IMPORTS: os, csv, shutil, sys, pathlib, datetime, pikepdf, tqdm, logging, hashlib
# FORBIDDEN: 禁止使用 open('w') 直接覆寫正式報表；禁止使用未經驗證的直接移動
# DEPENDENCIES: pikepdf, tqdm
# VERSION: 1.1.0 [Stability: Stable]
# ==========================================================

import os
import csv
import shutil
import logging
import hashlib
from pathlib import Path
from datetime import datetime
import pikepdf
from tqdm import tqdm

CONFIG = {
    "SAFE_ZONE_NAME": "SafeZone_Archived",
    "INFECTED_DIR": "!_Infected_Originals",
    "FAILED_DIR": "!_Failed_Processing",
    "REPORT": "00_Security_Audit_Report.csv",
    "LOG_FILE": "00_Execution_Log.log"
}

class AuditLogFormatter(logging.Formatter):
    def format(self, record):
        mapping = {'WARNING': 'WARN', 'CRITICAL': 'CRIT'}
        original_levelname = record.levelname
        record.levelname = mapping.get(original_levelname, original_levelname)
        result = super().format(record)
        record.levelname = original_levelname
        return result

class AegisArchiver:
    def __init__(self, src_path):
        self.src = Path(src_path).resolve()
        ts = datetime.now().strftime('%m%d_%H%M%S')
        
        self.safe_zone = self.src.parent / f"{CONFIG['SAFE_ZONE_NAME']}_{ts}"
        self.safe_zone.mkdir(parents=True, exist_ok=True)
        
        self.infected_zone = self.src / CONFIG["INFECTED_DIR"]
        self.failed_zone = self.src / CONFIG["FAILED_DIR"]
        
        self.log_data = []
        self.stats = {"TOTAL": 0, "PDF": 0, "SAFE": 0, "RISK": 0, "FAILED": 0}
        
        self._setup_logger()

    def _setup_logger(self):
        self.logger = logging.getLogger("AegisArchiver")
        self.logger.setLevel(logging.DEBUG)
        
        fh = logging.FileHandler(self.safe_zone / CONFIG["LOG_FILE"], encoding="utf-8")
        fh.setLevel(logging.DEBUG)
        
        formatter = AuditLogFormatter('[%(levelname)s] %(asctime)s %(name)s %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
        fh.setFormatter(formatter)
        self.logger.addHandler(fh)
        self.logger.info("AegisArchiver initialized. Safe zone created.")

    def _get_unique_path(self, target_dir, filename):
        base = target_dir / filename
        if not base.exists(): 
            return base
        counter = 1
        while (target_dir / f"{base.stem}({counter}){base.suffix}").exists():
            counter += 1
        return target_dir / f"{base.stem}({counter}){base.suffix}"

    def _calculate_sha256(self, filepath):
        sha256_hash = hashlib.sha256()
        with open(filepath, "rb") as f:
            for byte_block in iter(lambda: f.read(4096), b""):
                sha256_hash.update(byte_block)
        return sha256_hash.hexdigest()

    def _verified_move(self, src, dst):
        try:
            shutil.copy2(src, dst)
            src_hash = self._calculate_sha256(src)
            dst_hash = self._calculate_sha256(dst)
            
            if src_hash == dst_hash:
                os.remove(src)
            else:
                raise IOError(f"SHA-256 mismatch during move. Aborting deletion of original file: {src.name}")
        except Exception as e:
            if dst.exists():
                try:
                    os.remove(dst)
                    self.logger.info(f"Cleaned up partial/corrupted target file at: {dst}")
                except Exception as clean_e:
                    self.logger.warning(f"Failed to clean target file {dst}: {clean_e}")
            raise e

    def handle_pdf(self, path):
        is_risky = False
        safe_path = self._get_unique_path(self.safe_zone, path.name)
        placeholder_path = self._get_unique_path(self.safe_zone, f"{path.name}.quarantined.txt")
        
        try:
            with pikepdf.open(path) as pdf:
                risk_tags = ["/JS", "/JavaScript", "/OpenAction", "/Launch"]
                is_risky = any(tag in pdf.root for tag in risk_tags)
                if "/Names" in pdf.root and "/JavaScript" in pdf.root.Names: 
                    is_risky = True
                
                if is_risky:
                    self.logger.warning(f"Malicious tags detected in: {path.name}")
                else:
                    for page in pdf.pages:
                        if "/Annots" in page:
                            for annot in page.Annots:
                                if annot.get("/Subtype") == "/Link":
                                    if "/A" in annot: del annot["/A"]
                                    if "/AA" in annot: del annot["/AA"]
                    pdf.save(safe_path)

            if is_risky:
                self.infected_zone.mkdir(exist_ok=True)
                inf_path = self._get_unique_path(self.infected_zone, path.name)
                self._verified_move(path, inf_path)
                self.logger.info(f"Original risky file verified and isolated: {inf_path}")
                
                with open(placeholder_path, "w", encoding="utf-8") as f:
                    f.write("[FILE QUARANTINED] 檔案已被隔離\n")
                    f.write("====================================\n")
                    f.write(f"原始檔名: {path.name}\n")
                    f.write(f"隔離原因: 偵測到高風險標籤 (/JS, /Launch, /OpenAction 等)\n")
                    f.write(f"隔離位置: {inf_path}\n")
                    f.write(f"處理時間: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                
                return "RISK", f"Isolated at: {inf_path.parent.name} (Placeholder created)", placeholder_path
            else:
                self.logger.info(f"PDF processed and safely archived: {path.name}")
                return "SUCCESS", "Links stripped, safely archived", safe_path

        except Exception as e:
            self.logger.error(f"Processing failed for {path.name}: {str(e)}")
            
            if safe_path.exists():
                try:
                    os.remove(safe_path)
                    self.logger.info(f"Cleaned up orphaned safe file for {path.name} after processing failure.")
                except Exception as clean_e:
                    self.logger.warning(f"Failed to clean orphaned safe file {safe_path}: {clean_e}")
                    
            if placeholder_path.exists():
                try:
                    os.remove(placeholder_path)
                    self.logger.info(f"Cleaned up orphaned placeholder file for {path.name}.")
                except Exception as clean_e:
                    self.logger.warning(f"Failed to clean orphaned placeholder {placeholder_path}: {clean_e}")

            self.failed_zone.mkdir(exist_ok=True)
            
            if path.exists():
                fail_path = self._get_unique_path(self.failed_zone, path.name)
                try:
                    shutil.copy2(path, fail_path)
                    final_record_path = fail_path
                except Exception as copy_e:
                    final_record_path = "[LOCKED_OR_UNREADABLE]"
                    self.logger.error(f"Failed to copy locked/unreadable file {path.name}: {copy_e}")
            else:
                final_record_path = "[FILE_LOST_DURING_PROCESS]"
                self.logger.error(f"Source file {path.name} lost before failure fallback could execute.")
                
            return "FAILED", f"Corrupted/Failed: {str(e)}", final_record_path

    def execute(self):
        exts = [".pdf", ".txt", ".jpg", ".jpeg", ".png", ".gif", ".xlsx", ".docx", ".pptx"]
        files = [f for f in self.src.rglob("*") if f.is_file() and f.suffix.lower() in exts]
        files = [f for f in files if all(d not in str(f) for d in [CONFIG["INFECTED_DIR"], CONFIG["FAILED_DIR"], CONFIG["SAFE_ZONE_NAME"]])]
        
        self.stats["TOTAL"] = len(files)
        print("Starting deep scan and archival process...")
        self.logger.info(f"Scan complete. Total files to process: {self.stats['TOTAL']}")

        for f in tqdm(files, desc="Processing Files", ascii=True):
            ext = f.suffix.lower()
            size_mb = 0.0
            
            # 【P2 修復】I/O 前置檢查：阻斷因 getsize 引發的全域崩潰
            try:
                size_mb = round(os.path.getsize(f) / (1024 * 1024), 2)
            except Exception as e:
                self.logger.error(f"Pre-process I/O failed for {f.name}: {str(e)}")
                self.stats["FAILED"] += 1
                self.log_data.append([f.name, f.parent, size_mb, "FAILED", f"I/O Probe Failed: {str(e)}", "[LOCKED_OR_LOST_PRE_PROCESS]"])
                continue
            
            if ext == ".pdf":
                res, det, final_p = self.handle_pdf(f)
                if res == "RISK": self.stats["RISK"] += 1
                elif res == "FAILED": self.stats["FAILED"] += 1
                else: self.stats["PDF"] += 1
            else:
                final_p = self._get_unique_path(self.safe_zone, f.name)
                try:
                    shutil.copy2(f, final_p)
                    res, det = "SUCCESS", "Standard copy"
                    self.stats["SAFE"] += 1
                    self.logger.info(f"File safely archived: {f.name}")
                except Exception as e:
                    self.logger.error(f"Copy failed for {f.name}: {str(e)}")
                    self.failed_zone.mkdir(exist_ok=True)
                    
                    if final_p.exists():
                        try:
                            os.remove(final_p)
                            self.logger.info(f"Cleaned up partial safe file for {f.name}.")
                        except Exception as clean_e:
                            self.logger.warning(f"Failed to clean partial safe file {final_p}: {clean_e}")
                    
                    if f.exists():
                        fail_path = self._get_unique_path(self.failed_zone, f.name)
                        try:
                            shutil.copy2(f, fail_path)
                            final_p = fail_path
                            det = f"Copy Failed (Backed up): {str(e)}"
                        except Exception as copy_e:
                            final_p = "[LOCKED_OR_UNREADABLE]"
                            det = f"Critical Failure (Cannot copy): {str(copy_e)}"
                    else:
                        final_p = "[FILE_LOST_DURING_PROCESS]"
                        det = f"Source lost: {str(e)}"
                        
                    res = "FAILED"
                    self.stats["FAILED"] += 1
            
            self.log_data.append([f.name, f.parent, size_mb, res, det, final_p])

        self._finalize()

    def _finalize(self):
        report_path = self.safe_zone / CONFIG["REPORT"]
        temp_report_path = report_path.with_suffix(".tmp")
        
        try:
            with open(temp_report_path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(["FileName", "OriginalPath", "Size(MB)", "Status", "Details", "FinalLocation"])
                writer.writerows(self.log_data)
            
            os.replace(temp_report_path, report_path)
            self.logger.info("Audit report successfully generated.")
            
        except Exception as e:
            self.logger.critical(f"Failed to write audit report: {e}")

        print("\n" + "="*70)
        print("Execution Completed.")
        print(f"Total: {self.stats['TOTAL']} | Isolated: {self.stats['RISK']} | Failed: {self.stats['FAILED']}")
        print(f"Audit Report: {report_path}")
        print("="*70)
        
        self.logger.info("Process terminated normally.")

if __name__ == "__main__":
    p = input("Enter source directory path: ").strip().strip('"')
    if os.path.isdir(p):
        AegisArchiver(p).execute()
    else:
        print("Error: Invalid directory path.")
