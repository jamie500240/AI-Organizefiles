# ==========================================================
# MODULE:       Script_DirectoryFlattener
# PURPOSE:      自動扁平化目標資料夾，具備兩階段驗證搬移、內容雜湊比對與安全復原機制
# EXPORTS:      flatten_and_report
# IMPORTS:      os, shutil, csv, datetime, hashlib, logging
# FORBIDDEN:    禁止使用未驗證的 shutil.move 進行跨目錄轉移；禁止忽略可能的檔案權限或系統檔案錯誤；禁止缺乏 Rollback 機制的批次修改。
# DEPENDENCIES: 僅依賴 Python 標準庫 (Standard Library)
# VERSION:      1.0.0 [Stability: Experimental]
# ==========================================================

import os
import shutil
import csv
import hashlib
import logging
from datetime import datetime

def get_file_hash(filepath, chunk_size=8192):
    """計算檔案的 SHA-256 雜湊值"""
    hasher = hashlib.sha256()
    with open(filepath, 'rb') as f:
        for chunk in iter(lambda: f.read(chunk_size), b''):
            hasher.update(chunk)
    return hasher.hexdigest()

def setup_logger(log_filepath):
    """設定結構化日誌 (P3 規範)"""
    logger = logging.getLogger("FlattenerLogger")
    logger.setLevel(logging.INFO)
    
    # 避免重複 addHandler
    if not logger.handlers:
        fh = logging.FileHandler(log_filepath, encoding='utf-8')
        formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
        fh.setFormatter(formatter)
        logger.addHandler(fh)
    return logger

def flatten_and_report(target_dir):
    ts_str = datetime.now().strftime('%Y%m%d_%H%M%S')
    csv_filename = f"flatten_report_{ts_str}.csv"
    log_filename = f"flatten_syslog_{ts_str}.log"
    
    csv_path = os.path.join(target_dir, csv_filename)
    log_path = os.path.join(target_dir, log_filename)
    
    logger = setup_logger(log_path)
    logger.info(f"啟動扁平化任務，目標目錄: {target_dir}")

    if not os.path.exists(target_dir):
        logger.error(f"錯誤：找不到路徑 '{target_dir}'")
        print(f"錯誤：找不到路徑 '{target_dir}'")
        return

    print(f"開始處理資料夾: {target_dir}")
    print("正在執行 Phase 1: 複製與雜湊驗證 (若遇錯將自動復原)...")

    report_data = []
    transactions = []     # 記錄已成功複製的檔案，供 Rollback 使用
    sources_to_delete = [] # 記錄 Phase 1 成功後，待刪除的來源檔 (含真重複檔案)
    dirs_to_check = []    # 記錄處理過的目錄，待清空

    has_error = False
    
    # =========================================
    # Phase 1: 規劃、複製與驗證 (不刪除任何來源)
    # =========================================
    try:
        for root, dirs, files in os.walk(target_dir, topdown=False):
            if root == target_dir:
                continue
                
            dirs_to_check.append(root)

            for filename in files:
                source_path = os.path.join(root, filename)
                base, extension = os.path.splitext(filename)
                new_filename = filename
                destination_path = os.path.join(target_dir, new_filename)
                
                source_hash = get_file_hash(source_path)
                action_status = ""
                is_duplicate = False

                # 處理撞名與雜湊比對 (P4 規範)
                if os.path.exists(destination_path):
                    dest_hash = get_file_hash(destination_path)
                    if source_hash == dest_hash:
                        # 內容完全一致，視為真重複，不複製
                        logger.info(f"發現真重複檔案: {source_path} 與 {destination_path} 內容一致，標記為略過複製。")
                        action_status = "成功 (內容重複，略過複製並清除來源)"
                        is_duplicate = True
                    else:
                        # 檔名相同但內容不同，加流水號
                        counter = 1
                        while os.path.exists(destination_path):
                            new_filename = f"{base}_{counter}{extension}"
                            destination_path = os.path.join(target_dir, new_filename)
                            counter += 1
                        logger.info(f"檔名衝突但內容不同，重新命名為: {new_filename}")

                if not is_duplicate:
                    # 執行複製 (保留 metadata)
                    shutil.copy2(source_path, destination_path)
                    
                    # 複製後驗證 (P0 規範)
                    verify_hash = get_file_hash(destination_path)
                    if source_hash != verify_hash:
                        raise ValueError(f"雜湊驗證失敗: {source_path} -> {destination_path}")
                    
                    transactions.append(destination_path)
                    action_status = "成功 (複製且驗證通過)"
                    logger.info(f"複製並驗證成功: {source_path} -> {destination_path}")

                sources_to_delete.append(source_path)
                report_data.append({
                    "原始路徑": source_path,
                    "新檔名": new_filename if not is_duplicate else "[重複略過]",
                    "狀態": action_status
                })
                print(f"[檢驗與複製] {filename} -> {'[真重複略過]' if is_duplicate else new_filename}")

    except Exception as e:
        logger.error(f"Phase 1 發生嚴重錯誤: {str(e)}。觸發全批次復原 (Rollback)！")
        print(f"\n[錯誤] {str(e)}\n正在啟動 Rollback 復原機制...")
        has_error = True
        
        # 執行 Rollback (P0 規範)
        for dest_file in transactions:
            try:
                if os.path.exists(dest_file):
                    os.remove(dest_file)
                    logger.info(f"Rollback 刪除檔案: {dest_file}")
            except Exception as rb_e:
                logger.error(f"Rollback 失敗 (需人工介入): 無法刪除 {dest_file}, 原因: {rb_e}")
                
        print("Rollback 完成，來源資料未受損。任務終止。")

    # =========================================
    # Phase 2: 來源刪除與清理 (僅當 Phase 1 無錯時執行)
    # =========================================
    if not has_error:
        print("\nPhase 1 驗證通過，正在執行 Phase 2: 刪除原始檔案與空資料夾...")
        logger.info("啟動 Phase 2: 刪除原始檔案")
        
        for src_file in sources_to_delete:
            try:
                os.remove(src_file)
                logger.info(f"成功刪除來源檔: {src_file}")
            except Exception as e:
                logger.warning(f"來源檔刪除失敗: {src_file}, 原因: {e}")
                
        # 清理空目錄 (由下而上)
        for d in dirs_to_check:
            try:
                if not os.listdir(d):
                    os.rmdir(d)
                    logger.info(f"成功刪除空目錄: {d}")
            except Exception as e:
                # 解決靜默吞例外 (P1 規範)
                logger.warning(f"嘗試刪除目錄失敗 (可能非空或權限不足): {d}, 原因: {e}")

        # 產生 CSV 報表
        try:
            with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.DictWriter(f, fieldnames=["原始路徑", "新檔名", "狀態"])
                writer.writeheader()
                writer.writerows(report_data)
            logger.info(f"報表已產生: {csv_filename}")
            print("-" * 30)
            print(f"處理完成！\n報表已產生：{csv_filename}\n日誌已產生：{log_filename}")
        except Exception as e:
            logger.error(f"CSV 報表產生失敗: {e}")
            print(f"報表產生失敗: {e}")

if __name__ == "__main__":
    print("=== Script_DirectoryFlattener v1.0.1 (Experimental) ===")
    
    path_input = input("請輸入或拖入要拍扁的資料夾路徑: ").strip('"').strip("'")
    
    if os.path.isdir(path_input):
        confirm = input(f"確定要拍扁 '{path_input}' 嗎？此操作不可逆！(y/n): ")
        if confirm.lower() == 'y':
            flatten_and_report(path_input)
        else:
            print("操作已取消。")
    else:
        print("路徑無效，請確保它是一個實體資料夾。")

    print("\n" + "="*40)
    input("任務結束。請檢查資料夾與報表，確認無誤後按 Enter 鍵關閉視窗...")
