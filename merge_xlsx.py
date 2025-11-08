#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
XLSX檔案去重整合工具
功能：
1. 讀取指定資料夾中的所有xlsx檔案
2. 逐個整合檔案並顯示重複的筆數及詳細行位
3. 最終輸出去重後的CSV檔案
"""

import pandas as pd
import os
from pathlib import Path


def create_input_folder(folder_name='xlsx_files'):
    """建立輸入資料夾"""
    folder_path = Path(folder_name)
    if not folder_path.exists():
        folder_path.mkdir(parents=True, exist_ok=True)
        print(f"✓ 已建立資料夾: {folder_path.absolute()}")
    else:
        print(f"✓ 資料夾已存在: {folder_path.absolute()}")
    return folder_path


def merge_xlsx_files(input_folder='xlsx_files', output_file='merged_output.csv'):
    """
    整合xlsx檔案並去重
    
    參數:
        input_folder: 包含xlsx檔案的資料夾路徑
        output_file: 輸出的CSV檔案名稱
    """
    folder_path = Path(input_folder)
    
    # 檢查資料夾是否存在
    if not folder_path.exists():
        print(f"✗ 錯誤: 資料夾 '{input_folder}' 不存在")
        return
    
    # 取得所有xlsx檔案
    xlsx_files = list(folder_path.glob('*.xlsx'))
    
    if not xlsx_files:
        print(f"✗ 警告: 在 '{input_folder}' 資料夾中沒有找到xlsx檔案")
        print(f"請將xlsx檔案放入 '{folder_path.absolute()}' 資料夾後再執行")
        return
    
    print(f"\n找到 {len(xlsx_files)} 個xlsx檔案")
    print("=" * 60)
    
    # 用於儲存所有資料的DataFrame
    all_data = pd.DataFrame()
    total_rows_before = 0
    total_duplicates = 0
    
    # 逐個讀取並整合xlsx檔案
    for i, file_path in enumerate(xlsx_files, 1):
        print(f"\n[{i}/{len(xlsx_files)}] 正在處理: {file_path.name}")
        
        try:
            # 讀取xlsx檔案
            df = pd.read_excel(file_path)
            rows_in_file = len(df)
            print(f"  - 檔案包含 {rows_in_file} 筆資料（原始欄位：{len(df.columns)} 欄）")
            
            # 只保留"題目"和"答案"欄位
            required_columns = ['題目', '答案']
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                print(f"  ✗ 警告：檔案缺少必要欄位 {missing_columns}，跳過此檔案")
                print(f"      現有欄位：{list(df.columns)}")
                continue
            
            # 只選取需要的欄位
            df = df[required_columns]
            print(f"  - 已篩選保留 2 個欄位：題目、答案")
            
            total_rows_before += rows_in_file
            
            # 合併前的總行數
            before_merge = len(all_data)
            
            # 將當前檔案資料新增到總資料中
            all_data = pd.concat([all_data, df], ignore_index=True)
            
            # 找出與已合併資料重複的行
            after_concat = len(all_data)
            
            # 找出本次新增檔案中與已合併資料重複的行
            duplicate_count = 0
            
            if before_merge > 0:  # 如果已有合併的資料
                for idx in range(before_merge, after_concat):
                    current_row = all_data.iloc[idx]
                    # 檢查是否與已合併的資料（before_merge之前的資料）重複
                    is_duplicate = any(
                        all_data.iloc[:before_merge].apply(
                            lambda x: x.equals(current_row), axis=1
                        )
                    )
                    if is_duplicate:
                        duplicate_count += 1
            
            # 去重
            all_data = all_data.drop_duplicates()
            after_dedup = len(all_data)
            
            total_duplicates += duplicate_count
            
            print(f"  - 本次整合發現 {duplicate_count} 筆與已合併資料重複")
            print(f"  - 目前累計總資料: {after_dedup} 筆")
            
        except Exception as e:
            print(f"  ✗ 讀取檔案時發生錯誤: {e}")
            continue
    
    # 輸出統計資訊
    print("\n" + "=" * 60)
    print("整合完成！統計資訊：")
    print(f"  - 總共讀取: {total_rows_before} 筆資料")
    print(f"  - 重複資料: {total_duplicates} 筆")
    print(f"  - 去重後資料: {len(all_data)} 筆")
    print("=" * 60)
    
    # 匯出為CSV
    if not all_data.empty:
        output_path = Path(output_file)
        all_data.to_csv(output_path, index=False, encoding='utf-8-sig')
        print(f"\n✓ 成功匯出CSV檔案: {output_path.absolute()}")
        print(f"  檔案大小: {output_path.stat().st_size / 1024:.2f} KB")
        print(f"  資料欄位: {list(all_data.columns)}")
        print(f"  資料列數: {len(all_data)}")
    else:
        print("\n✗ 沒有資料可匯出")


def main():
    """主函式"""
    print("=" * 60)
    print("XLSX檔案去重整合工具")
    print("=" * 60)
    
    # 建立輸入資料夾
    input_folder = 'xlsx_files'
    create_input_folder(input_folder)
    
    # 整合xlsx檔案並匯出csv
    print("\n開始整合處理...")
    merge_xlsx_files(input_folder=input_folder, output_file='merged_output.csv')
    
    print("\n" + "=" * 60)
    print("程式執行完畢")
    print("=" * 60)


if __name__ == '__main__':
    main()

