# pip install PyMuPDF pillow imagehash
# pip install imagehash

import fitz  # PyMuPDF
import pandas as pd
import json
import os
import datetime
import hashlib
import tabula
from collections import defaultdict
import io
import shutil
from pathlib import Path
from PIL import Image

# Define directories
pdf_dir = r'C:\Users\0127043\OneDrive - ENEOSグループ\練習チャネル\大西テスト\PowerAutomateで画像説明\Pathがある程度固まったので、こちらを実験用に\PDF'  # PDFが配置されているフォルダ
root_output_dir = os.path.dirname(pdf_dir)  # PDFフォルダと同じ階層
table_and_json_dir = os.path.join(root_output_dir, "ImageAndJSON")  # メイン出力フォルダ

# Create main output directory if it doesn't exist
os.makedirs(table_and_json_dir, exist_ok=True)

print("Starting PDF table extraction process...")

# 画像のハッシュを計算する関数
def get_image_hash(image_path):
    # バイナリハッシュ（完全に同じバイナリデータの場合のみ一致）
    with open(image_path, 'rb') as f:
        binary_hash = hashlib.md5(f.read()).hexdigest()
    
    return binary_hash

# フォルダ構造を作成する関数
def create_folder_structure(pdf_filename):
    """
    PDFファイル名に基づいてフォルダ構造を作成する
    
    Args:
        pdf_filename (str): PDFファイル名（拡張子を含む）
    
    Returns:
        tuple: (ドキュメントフォルダパス, 表フォルダパス, JSONフォルダパス)
    """
    # PDFファイル名から拡張子を除去
    doc_name = os.path.splitext(pdf_filename)[0]
    
    # フォルダパスを作成
    doc_folder = os.path.join(table_and_json_dir, doc_name)
    table_folder = os.path.join(doc_folder, "Image")
    json_folder = os.path.join(doc_folder, "JSON")
    
    # フォルダを作成
    os.makedirs(doc_folder, exist_ok=True)
    os.makedirs(table_folder, exist_ok=True)
    os.makedirs(json_folder, exist_ok=True)
    
    return doc_folder, table_folder, json_folder

# 表領域を画像として保存する関数
def save_table_as_image(page, rect, image_path, dpi=300):
    """
    表領域をレンダリングして画像として保存する
    
    Args:
        page: PDF Page object
        rect: 表の領域を示す矩形 (fitz.Rect)
        image_path: 保存先パス
        dpi: 解像度 (デフォルト300dpi)
    
    Returns:
        bool: 成功したかどうか
    """
    try:
        # 高解像度でレンダリングするための行列（拡大率設定）
        zoom = dpi / 72  # 72 DPIがデフォルト
        matrix = fitz.Matrix(zoom, zoom)
        
        # 表領域の周りに少し余白を追加
        padding = 5
        rect.x0 = max(0, rect.x0 - padding)
        rect.y0 = max(0, rect.y0 - padding)
        rect.x1 = min(page.rect.width, rect.x1 + padding)
        rect.y1 = min(page.rect.height, rect.y1 + padding)
        
        # ページの表領域をクリップしてレンダリング
        pixmap = page.get_pixmap(matrix=matrix, clip=rect)
        
        # 画像として保存
        pixmap.save(image_path)
        return True
        
    except Exception as e:
        print(f"Error saving table as image: {e}")
        return False

# PyMuPDFを使用して表を抽出し画像として保存する関数
def extract_tables_with_pymupdf(pdf_path, table_folder):
    pdf_document = fitz.open(pdf_path)
    pdf_filename = os.path.basename(pdf_path)
    table_data = []
    
    for page_num in range(len(pdf_document)):
        page = pdf_document[page_num]
        
        # 表を検出する
        table_finder = page.find_tables()
        tables = table_finder.tables if hasattr(table_finder, "tables") else []
        
        for table_index, table in enumerate(tables):
            try:
                # テーブル領域を取得
                if hasattr(table, "rect"):
                    rect = table.rect
                else:
                    # 矩形座標から手動で作成
                    rect = fitz.Rect(table.bbox[:4])  # bbox は [x0, y0, x1, y1] の形式
                
                # 表のデータをPandasデータフレームとして取得（メタデータ用）
                if hasattr(table, "to_pandas"):
                    df = table.to_pandas()
                else:
                    # DataFrameが直接取得できない場合、代わりに行数と列数を記録
                    rows = len(table.cells) if hasattr(table, "cells") else 0
                    cols = len(table.cells[0]) if rows > 0 and hasattr(table, "cells") else 0
                    df = pd.DataFrame(index=range(rows), columns=range(cols))
                
                if df.empty and (not hasattr(table, "cells") or len(table.cells) == 0):
                    continue
                
                # 画像ファイル名を生成
                table_image_filename = f"【{os.path.splitext(pdf_filename)[0]}】page{page_num+1}_table{table_index+1}.png"
                image_path = os.path.join(table_folder, table_image_filename)
                
                # 表を画像として保存
                if save_table_as_image(page, rect, image_path):
                    print(f"Saved table {table_index+1} from page {page_num+1} as image")
                    
                    # 表データを記録
                    column_names = df.columns.tolist() if not df.empty else []
                    
                    table_data.append({
                        "image_path": image_path,
                        "filename": table_image_filename,
                        "page_number": page_num + 1,
                        "rows": len(df) if not df.empty else 0,
                        "columns": len(df.columns) if not df.empty else 0,
                        "position": {
                            "x0": rect.x0,
                            "y0": rect.y0,
                            "x1": rect.x1,
                            "y1": rect.y1
                        },
                        "column_names": column_names,
                        "extraction_method": "pymupdf"
                    })
            except Exception as e:
                print(f"Error processing table {table_index} on page {page_num+1}: {e}")
    
    # 別のアプローチを試す：テーブル検出のバックアップ方法
    if len(table_data) == 0:
        print("No tables found with standard method, trying another approach...")
        
        for page_num in range(len(pdf_document)):
            page = pdf_document[page_num]
            
            # 表を検出する別の方法
            try:
                # テーブル構造認識のためのオプション設定
                tab = fitz.TableFinder(page)
                tab.extract()
                
                for table_index, table_rect in enumerate(tab.tables):
                    # 画像ファイル名を生成
                    table_image_filename = f"【{os.path.splitext(pdf_filename)[0]}】page{page_num+1}_alt_table{table_index+1}.png"
                    image_path = os.path.join(table_folder, table_image_filename)
                    
                    # 表を画像として保存
                    if save_table_as_image(page, table_rect, image_path):
                        print(f"Saved table {table_index+1} from page {page_num+1} as image (alternative method)")
                        
                        # 簡易的なメタデータを記録
                        table_data.append({
                            "image_path": image_path,
                            "filename": table_image_filename,
                            "page_number": page_num + 1,
                            "position": {
                                "x0": table_rect.x0,
                                "y0": table_rect.y0,
                                "x1": table_rect.x1,
                                "y1": table_rect.y1
                            },
                            "extraction_method": "pymupdf_alternative"
                        })
            except Exception as e:
                print(f"Error with alternative table detection on page {page_num+1}: {e}")
    
    # それでも表が見つからない場合は、pdfplumberまたはtabulaを使用する
    if len(table_data) == 0:
        print("Still no tables found, trying tabula...")
        try:
            # ページ全体の画像を生成し、表領域を保存
            for page_num in range(len(pdf_document)):
                # ページ全体を高解像度画像としてレンダリング
                page = pdf_document[page_num]
                pixmap = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72))
                temp_img_path = f"temp_page_{page_num}.png"
                pixmap.save(temp_img_path)
                
                # ページ全体を表画像として保存
                table_image_filename = f"【{os.path.splitext(pdf_filename)[0]}】page{page_num+1}_full_page_table.png"
                image_path = os.path.join(table_folder, table_image_filename)
                
                shutil.copy(temp_img_path, image_path)
                
                # メタデータを記録
                table_data.append({
                    "image_path": image_path,
                    "filename": table_image_filename,
                    "page_number": page_num + 1,
                    "position": {
                        "x0": 0,
                        "y0": 0,
                        "x1": page.rect.width,
                        "y1": page.rect.height
                    },
                    "extraction_method": "full_page"
                })
                
                # 一時ファイルを削除
                if os.path.exists(temp_img_path):
                    os.remove(temp_img_path)
                    
        except Exception as e:
            print(f"Error using fallback method: {e}")
    
    return table_data

# 重複テーブル画像をチェックする関数
def process_duplicate_tables(table_data):
    print("Checking for duplicate tables...")
    unique_tables = []
    duplicate_info = {}
    
    # テーブルハッシュ用のディクショナリ
    table_hashes = {}
    
    for i, table_info in enumerate(table_data):
        image_path = table_info["image_path"]
        
        try:
            # 画像のハッシュを計算
            image_hash = get_image_hash(image_path)
            
            # 重複チェック
            if image_hash in table_hashes:
                # 重複テーブルの情報を記録
                duplicate_reference = table_hashes[image_hash]
                if duplicate_reference not in duplicate_info:
                    duplicate_info[duplicate_reference] = []
                
                duplicate_info[duplicate_reference].append({
                    "path": image_path,
                    "page_number": table_info["page_number"]
                })
                
                # 重複テーブルの画像を削除
                if os.path.exists(image_path):
                    os.remove(image_path)
                print(f"Removed duplicate table: {os.path.basename(image_path)}")
            else:
                # イメージハッシュを登録
                table_hashes[image_hash] = image_path
                
                # 一意のテーブル情報を保存
                table_info["table_hash"] = image_hash
                unique_tables.append(table_info)
                
        except Exception as e:
            print(f"Error processing {image_path} for duplication: {e}")
            table_info["table_hash"] = "error_hash"
            unique_tables.append(table_info)
    
    # 重複情報をユニークテーブルに関連付ける
    for i, table_info in enumerate(unique_tables):
        image_path = table_info["image_path"]
        if image_path in duplicate_info:
            unique_tables[i]["duplicates"] = duplicate_info[image_path]
        else:
            unique_tables[i]["duplicates"] = []
    
    print(f"Kept {len(unique_tables)} unique tables, removed {len(table_data) - len(unique_tables)} duplicates")
    return unique_tables

# 表メタデータを保存する関数
def save_table_metadata(table_data_list, json_folder, source_pdf):
    for i, table_data in enumerate(table_data_list):
        try:
            table_filename = table_data["filename"]
            image_path = table_data["image_path"]
            page_number = table_data["page_number"]
            
            print(f"Processing table {i+1}/{len(table_data_list)}: {os.path.basename(image_path)} from page {page_number}")
            
            # 画像情報を取得
            with Image.open(image_path) as img:
                width, height = img.size
                format_type = img.format
                mode = img.mode
            
            # ファイル情報を取得
            file_size = os.path.getsize(image_path)
            file_name = os.path.basename(image_path)
            
            # 抽出日時
            extraction_date = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            # 重複ページ情報があれば追加
            duplicate_pages = []
            if "duplicates" in table_data and table_data["duplicates"]:
                duplicate_pages = [dup["page_number"] for dup in table_data["duplicates"]]
                print(f"  This table also appears on pages: {', '.join(map(str, duplicate_pages))}")
            
            # 相対パスを作成（JSONからの相対パス）
            rel_image_path = os.path.join("..", "Image", table_filename)
            
            # JSONデータを作成
            json_data = {
                "source_pdf": os.path.splitext(source_pdf)[0],
                "file_name": file_name,
                "pdf_page_number": page_number,
                #"duplicate_appearances": duplicate_pages,
                #"image_path": rel_image_path,
                #"width": width,
                #"height": height,
                #"format": format_type,
                #"color_mode": mode,
                #"extraction_method": table_data.get("extraction_method", "unknown"),
                #"extraction_date": extraction_date,
                #"file_size_bytes": file_size,
                #"table_hash": table_data["table_hash"],
                #"estimated_rows": table_data.get("rows", 0),
                #"estimated_columns": table_data.get("columns", 0),
                #"column_names": table_data.get("column_names", []),
                #"position": table_data.get("position", {}),
                "Summary": "",
                "LinkToSP": ""
            }
            
            # 表ファイル名と同じ名前でJSONファイルを作成
            json_filename = f"{os.path.splitext(table_filename)[0]}.json"
            json_path = os.path.join(json_folder, json_filename)
            with open(json_path, "w", encoding='utf-8') as json_file:
                json.dump(json_data, json_file, ensure_ascii=False, indent=4)

                print(f"  Created JSON metadata: {json_filename}")

            # JSONファイルの拡張子をtxtに変更
            #txt_filename = f"{os.path.splitext(table_filename)[0]}.txt"
            #txt_path = os.path.join(json_folder, txt_filename)
            #if os.path.exists(json_path):
            #    shutil.copy(json_path, txt_path)  # JSONファイルをTXT拡張子でコピー
            #    os.remove(json_path)  # 元のJSONファイルを削除
            #    print(f"  Renamed JSON to TXT: {txt_filename}") 
 
        except Exception as e:
            print(f"Error creating metadata for {table_data['filename']}: {e}")

# フォルダ内の全PDFファイルを処理
def process_pdf_folder(pdf_dir):
    # PDFフォルダ内のPDFファイルを検索
    pdf_files = [f for f in os.listdir(pdf_dir) if f.lower().endswith('.pdf')]
    
    if not pdf_files:
        print(f"No PDF files found in {pdf_dir}")
        return
    
    print(f"Found {len(pdf_files)} PDF files to process")
    
    # 各PDFファイルを処理
    for pdf_file in pdf_files:
        try:
            pdf_path = os.path.join(pdf_dir, pdf_file)
            print(f"\n{'='*60}\nProcessing PDF: {pdf_file}")
            
            # フォルダ構造を作成
            doc_folder, table_folder, json_folder = create_folder_structure(pdf_file)
            print(f"Created folders:\n  Document: {doc_folder}\n  Table Images: {table_folder}\n  JSON: {json_folder}")
            
            # PDFから表を抽出
            table_data_list = extract_tables_with_pymupdf(pdf_path, table_folder)
            print(f"Extracted {len(table_data_list)} tables from {pdf_file}")
            
            if len(table_data_list) > 0:
                # 重複チェックと処理
                unique_tables = process_duplicate_tables(table_data_list)
                
                # 各表のメタデータをJSONファイルとして保存
                save_table_metadata(unique_tables, json_folder, pdf_file)
            else:
                print(f"No tables found in {pdf_file}")
            
        except Exception as e:
            print(f"Error processing PDF {pdf_file}: {str(e)}")

# メイン処理
if __name__ == "__main__":
    try:
        # フォルダ内の全PDFを処理
        process_pdf_folder(pdf_dir)
        print("\nProcessing complete. All PDFs have been processed.")
        print(f"Results are saved in: {table_and_json_dir}")
    except Exception as e:
        print(f"Error during processing: {e}")