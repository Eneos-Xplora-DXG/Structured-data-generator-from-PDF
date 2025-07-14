# pip install PyMuPDF pillow imagehash
# pip install imagehash
import fitz  # PyMuPDF
from PIL import Image, ImageMath
import imagehash
import json
import os
import datetime
import hashlib
from collections import defaultdict
import io
import numpy as np
import shutil

# Define directories
pdf_dir = r'C:\Users\0127043\OneDrive - ENEOSグループ\練習チャネル\大西テスト\PowerAutomateで画像説明\Pathがある程度固まったので、こちらを実験用に\PDF'  # PDFが配置されているフォルダ
root_output_dir = os.path.dirname(pdf_dir)  # PDFフォルダと同じ階層
image_and_json_dir = os.path.join(root_output_dir, "ImageAndJSON")  # メイン出力フォルダ

# Create main output directory if it doesn't exist
os.makedirs(image_and_json_dir, exist_ok=True)

print("Starting PDF image extraction process...")

# 画像の重複をチェックするためのディクショナリ
image_hashes = {}
binary_hashes = {}

# 画像ハッシュを計算する関数
def get_image_hash(image_path):
    # バイナリハッシュ（完全に同じバイナリデータの場合のみ一致）
    with open(image_path, 'rb') as f:
        binary_hash = hashlib.md5(f.read()).hexdigest()
    
    # 画像内容のパーセプチュアルハッシュ（見た目が似ている場合も検出）
    try:
        img = Image.open(image_path)
        # 複数のハッシュアルゴリズムを組み合わせて精度を向上
        phash = str(imagehash.phash(img))
        dhash = str(imagehash.dhash(img))
        combined_hash = f"{phash}_{dhash}"
        return binary_hash, combined_hash
    except Exception:
        # 画像が開けない場合はバイナリハッシュのみ
        return binary_hash, None

# 画像とマスクを適切に合成する関数
def process_image_with_mask(image_bytes, mask_bytes=None, smask_bytes=None):
    """画像とマスク情報を適切に合成する"""
    try:
        # 元の画像をPILで開く
        img = Image.open(io.BytesIO(image_bytes))
        
        # マスクがない場合は元の画像をそのまま返す
        if mask_bytes is None and smask_bytes is None:
            return image_bytes, img.format.lower() if img.format else "png"
        
        # 画像フォーマットを決定
        img_format = img.format.lower() if img.format else "png"
        
        # SMaskがある場合（透明度情報）
        if smask_bytes:
            try:
                # SMaskをPILで開く
                smask = Image.open(io.BytesIO(smask_bytes)).convert("L")
                
                # 画像のモードを確認し、必要に応じて変換
                if img.mode not in ("RGBA", "LA"):
                    if img.mode in ("RGB", "L"):
                        img = img.convert("RGBA" if img.mode == "RGB" else "LA")
                    else:
                        # その他のモードはRGBAに変換
                        img = img.convert("RGBA")
                
                # 画像とマスクのサイズを合わせる
                if img.size != smask.size:
                    smask = smask.resize(img.size)
                
                # アルファチャンネルをSMaskから設定
                if img.mode == "RGBA":
                    r, g, b, _ = img.split()
                    img = Image.merge("RGBA", (r, g, b, smask))
                elif img.mode == "LA":
                    l, _ = img.split()
                    img = Image.merge("LA", (l, smask))
            except Exception as e:
                print(f"Error processing SMask: {e}")
                # SMaskの処理が失敗した場合は元の画像を使用
        
        # 通常のマスクがある場合
        elif mask_bytes:
            try:
                # マスクをPILで開く
                mask = Image.open(io.BytesIO(mask_bytes)).convert("1")  # 2値マスク
                
                # 画像のモードを確認し、必要に応じて変換
                if img.mode not in ("RGBA", "LA"):
                    if img.mode in ("RGB", "L"):
                        img = img.convert("RGBA" if img.mode == "RGB" else "LA")
                    else:
                        img = img.convert("RGBA")
                
                # 画像とマスクのサイズを合わせる
                if img.size != mask.size:
                    mask = mask.resize(img.size)
                
                # マスクを適用
                if img.mode == "RGBA":
                    r, g, b, a = img.split()
                    new_a = ImageMath.eval("convert(a & mask, 'L')", a=a, mask=mask)
                    img = Image.merge("RGBA", (r, g, b, new_a))
                elif img.mode == "LA":
                    l, a = img.split()
                    new_a = ImageMath.eval("convert(a & mask, 'L')", a=a, mask=mask)
                    img = Image.merge("LA", (l, new_a))
            except Exception as e:
                print(f"Error processing mask: {e}")
                # マスクの処理が失敗した場合は元の画像を使用
        
        # 処理後の画像をバイト配列に変換
        output = io.BytesIO()
        # PNGで保存するとアルファチャンネル（透明度）は保持される
        img.save(output, format="PNG")
        return output.getvalue(), "png"
        
    except Exception as e:
        print(f"Error in image processing: {e}")
        # エラーが発生した場合は元の画像を返す
        return image_bytes, "png"

# フォルダ構造を作成する関数
def create_folder_structure(pdf_filename):
    """
    PDFファイル名に基づいてフォルダ構造を作成する
    
    Args:
        pdf_filename (str): PDFファイル名（拡張子を含む）
    
    Returns:
        tuple: (ドキュメントフォルダパス, 画像フォルダパス, JSONフォルダパス)
    """
    # PDFファイル名から拡張子を除去
    doc_name = os.path.splitext(pdf_filename)[0]
    
    # フォルダパスを作成
    doc_folder = os.path.join(image_and_json_dir, doc_name)
    image_folder = os.path.join(doc_folder, "Image")
    json_folder = os.path.join(doc_folder, "JSON")
    
    # フォルダを作成
    os.makedirs(doc_folder, exist_ok=True)
    os.makedirs(image_folder, exist_ok=True)
    os.makedirs(json_folder, exist_ok=True)
    
    return doc_folder, image_folder, json_folder

# Function to extract images from PDF with proper mask handling
def extract_images_from_pdf(pdf_path, image_folder):
    pdf_document = fitz.open(pdf_path)
    pdf_filename = os.path.basename(pdf_path)
    image_data = []  # List to store image data with page numbers
    
    # 各ページから画像を抽出
    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        image_list = page.get_images(full=True)
        
        for img_index, img in enumerate(image_list):
            try:
                xref = img[0]
                
                # 画像を抽出
                base_image = pdf_document.extract_image(xref)
                image_bytes = base_image["image"]
                image_ext = base_image["ext"]
                
                # マスク情報があるか確認
                mask_xref = base_image.get("mask", 0)
                smask_xref = base_image.get("smask", 0)
                
                mask_bytes = None
                smask_bytes = None
                
                # マスクがある場合、抽出
                if mask_xref:
                    try:
                        mask_img = pdf_document.extract_image(mask_xref)
                        mask_bytes = mask_img["image"]
                    except Exception as e:
                        print(f"Error extracting mask: {e}")
                
                # SMaskがある場合、抽出
                if smask_xref:
                    try:
                        smask_img = pdf_document.extract_image(smask_xref)
                        smask_bytes = smask_img["image"]
                    except Exception as e:
                        print(f"Error extracting smask: {e}")
                
                # 画像とマスクを適切に処理
                processed_image_bytes, processed_ext = process_image_with_mask(
                    image_bytes, mask_bytes, smask_bytes
                )
                
                # 画像ファイル名を生成
                image_filename =  "【"+ os.path.splitext(pdf_filename)[0]+"】" + f"page{page_num+1}_img{img_index}.{processed_ext}"
                image_path = os.path.join(image_folder, image_filename)
                
                # 画像を保存
                with open(image_path, "wb") as image_file:
                    image_file.write(processed_image_bytes)
                
                # 画像情報を記録
                image_data.append({
                    "path": image_path,
                    "filename": image_filename,
                    "page_number": page_num + 1,
                    "xref": xref,
                    "has_mask": bool(mask_bytes),
                    "has_smask": bool(smask_bytes)
                })
                
            except Exception as e:
                print(f"Error extracting image {img_index} from page {page_num+1}: {e}")
    
    return image_data

# Function to check for duplicates and update metadata
def process_duplicates(image_data, pdf_filename):
    print("Checking for duplicate images...")
    unique_images = []
    duplicate_info = {}
    
    # ドキュメントごとに重複検出用の辞書をクリア
    local_binary_hashes = {}
    local_image_hashes = {}
    
    for i, img_data in enumerate(image_data):
        image_path = img_data["path"]
        page_number = img_data["page_number"]
        
        try:
            binary_hash, perceptual_hash = get_image_hash(image_path)
            
            # 重複チェック（同一PDFファイル内のみ）
            is_duplicate = False
            duplicate_reference = None
            
            # バイナリが完全一致する場合
            if binary_hash in local_binary_hashes:
                is_duplicate = True
                duplicate_reference = local_binary_hashes[binary_hash]
            # 知覚的ハッシュが一致する場合（画像内容が似ている）
            elif perceptual_hash and perceptual_hash in local_image_hashes:
                is_duplicate = True
                duplicate_reference = local_image_hashes[perceptual_hash]
            
            if is_duplicate:
                # 重複画像の情報を記録
                if duplicate_reference not in duplicate_info:
                    duplicate_info[duplicate_reference] = []
                duplicate_info[duplicate_reference].append({
                    "path": image_path,
                    "page_number": page_number
                })
                
                # 重複画像を削除
                os.remove(image_path)
                print(f"Removed duplicate image: {os.path.basename(image_path)}")
            else:
                # ユニークな画像として記録
                local_binary_hashes[binary_hash] = image_path
                if perceptual_hash:
                    local_image_hashes[perceptual_hash] = image_path
                
                # ユニーク画像のデータを保存
                unique_images.append({
                    "path": image_path,
                    "filename": img_data["filename"],
                    "page_number": page_number,
                    "binary_hash": binary_hash,
                    "has_mask": img_data.get("has_mask", False),
                    "has_smask": img_data.get("has_smask", False),
                    "duplicates": []
                })
        except Exception as e:
            print(f"Error processing {image_path} for duplication: {e}")
            # エラーが発生した場合でも、画像をユニークとして扱う
            unique_images.append({
                "path": image_path,
                "filename": img_data["filename"],
                "page_number": page_number,
                "binary_hash": "error_hash",
                "has_mask": img_data.get("has_mask", False),
                "has_smask": img_data.get("has_smask", False),
                "duplicates": []
            })
    
    # 重複情報を元の画像に紐づける
    for i, img_data in enumerate(unique_images):
        img_path = img_data["path"]
        if img_path in duplicate_info:
            unique_images[i]["duplicates"] = duplicate_info[img_path]
    
    print(f"Kept {len(unique_images)} unique images, removed {len(image_data) - len(unique_images)} duplicates")
    return unique_images

# Function to collect image metadata and save as JSON
def save_image_metadata(image_data_list, json_folder, source_pdf):
    for i, image_data in enumerate(image_data_list):
        try:
            image_path = image_data["path"]
            image_filename = image_data["filename"]
            page_number = image_data["page_number"]
            duplicate_info = image_data.get("duplicates", [])
            
            print(f"Processing image {i+1}/{len(image_data_list)}: {os.path.basename(image_path)} from page {page_number}")
            
            # 画像を開いてメタデータを取得
            with Image.open(image_path) as img:
                width, height = img.size
                format_type = img.format
                mode = img.mode
                has_transparency = "transparency" in img.info or mode in ("RGBA", "LA")
            
            # ファイル情報を取得
            file_size = os.path.getsize(image_path)
            file_name = os.path.basename(image_path)
            
            # 抽出日時
            extraction_date = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            # 重複ページ情報があれば追加
            duplicate_pages = []
            if duplicate_info:
                duplicate_pages = [dup["page_number"] for dup in duplicate_info]
                print(f"  This image also appears on pages: {', '.join(map(str, duplicate_pages))}")
            
            # 相対パスを作成（JSONからの相対パス）
            rel_path = os.path.join("..", "Image", image_filename)
            
            # JSONデータを作成（ページ情報と重複情報を含む）
            json_data = {
                #"image_path": rel_path,
                "source_pdf": os.path.splitext(source_pdf)[0],
                "file_name": file_name,
                "pdf_page_number": page_number,
                "Summary":"",
                "LinkToSP":""
                #"duplicate_appearances": duplicate_pages,
                #"extraction_date": extraction_date,
                #"file_size_bytes": file_size,
                #"width": width,
                #"height": height,
                #"format": format_type,
                #"color_mode": mode,
                #"has_transparency": has_transparency,
                #"has_mask": image_data.get("has_mask", False),
                #"has_smask": image_data.get("has_smask", False),
                #"image_hash": image_data["binary_hash"]
            }
            
            # 画像ファイル名と同じ名前でJSONファイルを作成
            json_filename = f"{os.path.splitext(image_filename)[0]}.json"
            json_path = os.path.join(json_folder, json_filename)
            with open(json_path, "w", encoding='utf-8') as json_file:
                json.dump(json_data, json_file, ensure_ascii=False, indent=4)
            
            print(f"  Created JSON metadata: {json_filename}")

            # JSONファイルの拡張子をtxtに変更
            #txt_filename = f"{os.path.splitext(image_filename)[0]}.txt"
            #txt_path = os.path.join(json_folder, txt_filename)
            #if os.path.exists(json_path):
            #    shutil.copy(json_path, txt_path)  # JSONファイルをTXT拡張子でコピー
            #    os.remove(json_path)  # 元のJSONファイルを削除
            #    print(f"  Renamed JSON to TXT: {txt_filename}") 
                                
        except Exception as e:
            print(f"Error processing {image_path}: {str(e)}")

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
            doc_folder, image_folder, json_folder = create_folder_structure(pdf_file)
            print(f"Created folders:\n  Document: {doc_folder}\n  Image: {image_folder}\n  JSON: {json_folder}")
            
            # PDFから画像を抽出
            image_data_list = extract_images_from_pdf(pdf_path, image_folder)
            print(f"Extracted {len(image_data_list)} images from {pdf_file}")
            
            # 重複チェックと処理
            unique_images = process_duplicates(image_data_list, pdf_file)
            
            # 各画像のメタデータをJSONファイルとして保存
            save_image_metadata(unique_images, json_folder, pdf_file)
            
        except Exception as e:
            print(f"Error processing PDF {pdf_file}: {str(e)}")

# メイン処理
if __name__ == "__main__":
    try:
        # フォルダ内の全PDFを処理
        process_pdf_folder(pdf_dir)
        print("\nProcessing complete. All PDFs have been processed.")
        print(f"Results are saved in: {image_and_json_dir}")
    except Exception as e:
        print(f"Error during processing: {e}")