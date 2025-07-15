# Structured data generator from PDF
PythonとPowerAutomateを用いて、PDFから非構造化データ(図と表)を抽出し、構造化データに変換するためのソース

# 背景・目的
生成AIのRAG機能では、図や表といった非構造化データの内容を参照しても、その回答精度が上がらないという問題がある。
そこで、RAGで参照するドキュメントの中から、このような非構造化データを抽出し、構造化データに変換する処理を実装した。
具体的には、2025/6段階で、技術管理部からこのような要望が挙がったことに端を発する。

# 使用するツール
1. Python
  - PDFから、図表を画像形式で抜き出すとともに、それらのメタデータが格納されたJSONファイルを出力（JSONファイルについては後述）
  - PyMuPDF pillow imagehashとimagehashのpip installが必要。その他のimportはソース内冒頭に記載
2. PowerAutomate
  - 1.で抽出した図表の画像に対し、AI Builderで説明と、画像が配置されたSharePointパス情報をJSONに追記
  - AI Builderを用いるため、AIクレジットが必要

# 処理の流れ
1. Pythonコードを実行すると、PDFファイル内の図や表を抽出し、画像としてフォルダに保存される。  
   この時同時に、各画像についてのメタデータ（ページ数や、生成された画像のファイル名）が格納される。  
   生成されるフォルダは以下の通り。

    **[フォルダ構成(処理前)]**  
   <img width="155" height="151" alt="image" src="https://github.com/user-attachments/assets/ee9a1aa1-05f0-402c-92f0-df05670543aa" />
   
    **[フォルダ構成(処理後)]**  
    <img width="385" height="798" alt="image" src="https://github.com/user-attachments/assets/41731bd8-416d-4e6e-b85b-543b27b35df9" />

3. PowerAutomateフローを実行すると、各PDF名フォルダ内の「Image」のフォルダ内の画像を1つ1つ参照し、AI Builderによる画像説明の文章を作成する。  
その説明文と、該当の画像が配置されたSharePointURLを、対応する「JSON」に書き込む。  
このようにして完成したJSONの中身は以下の通り。  

    **[JSONの中身]**  
    <img width="943" height="131" alt="image" src="https://github.com/user-attachments/assets/382fa85e-d467-48a1-9a75-2f7b6f40a82f" />

    これらのJSONファイルは画像ごとに分かれてしまっているため、最後に1つのテキストファイルに合体し、txtファイル化する（GENEOSにはjsonファイルはアップロードできないため）。  
    この合体版txtファイル(【PDF名】図表の構造化データ.txt)を、GENEOSのナレッジフォルダに格納して用いる。

    **[全処理完了時のフォルダとファイル]**  
   <img width="424" height="863" alt="image" src="https://github.com/user-attachments/assets/088e6fd0-f105-4669-94f3-42dbb4b46873" />


# デポジトリに保存された各種ファイルの説明
1. PDFからimage抽出.py  
  - PDFから図を抽出して画像として保存し、対応するJSONも作成するソース。
2. PDFからtable抽出.py
  - PDFから表を抽出して画像として保存し、対応するJSONも作成するソース。
3. PowerAutomateフロー内に記載のプロンプト
  - PowerAutomateフローのAIビルダーに記載のプロンプト。

※1.と2.は将来的に1つのソースにしてもよい。

# 使用場所
  - 2025/7/14現在、PythonはVS Code上でデバッグをして使用している。  
    今後、PyinstallerでPythonソースをコンパイルし、exe化して展開する想定。  
    exe実行の際には、引数としてファイルパス指定が必要（詳細は後述）。
  - 対象となるPDFは、SharePoint上の任意のフォルダに配置する（詳細は後述）。

# 使い方説明
1. Pythonについて
  - 実行にはまず、**PDFが配置されたフォルダ**を引数として準備する必要がある。  
    上記のフォルダ構成図の「…/PDF」までのフォルダを指す。  
    このPDFフォルダは、OneDrive上にショートカット作成したSharePoint上で行うことを推奨する  
    （ただし、**SharePointのフォルダ階層が深すぎると、後のPowerAutomateフローでエラーになるため注意。**）

    **[デバッグで実行する場合]**  
    以下の、**pdf_dir**の中身のフォルダパスを変更する
    <img width="1236" height="45" alt="image" src="https://github.com/user-attachments/assets/73fde262-9d94-4f70-9a72-3032485927b2" />

    **[exeで実行する場合]**  
    Windows + Rを押し、**"exeの配置パス" "フォルダのパス"**を入力してエンターする  
    <img width="414" height="225" alt="image" src="https://github.com/user-attachments/assets/a551ce7e-a863-4686-90c6-e3c8e303fc02" />

2. PowerAutomateについて
  - PowerAutomateフロー内の**要変更**となっている部分は変更が必要。  
    SharePointのURLから、必要事項を入力する

    **[【★要変更】SharePointサイトのURL]**  
    SharePointのURLの冒頭部分を入力する  
    <img width="487" height="149" alt="image" src="https://github.com/user-attachments/assets/071dec26-f0b6-498f-8b11-94aae0afaec2" />  
    <img width="664" height="44" alt="image" src="https://github.com/user-attachments/assets/bcb9fd1b-761e-42bf-8ee5-497c95c7f92a" />

    **[【★要変更】SharePointサイトのファイル識別子(ImageAndJSON直下ドキュメント名フォルダまで)]**  
    上記の冒頭部分を除いたURL～「ImageAndJSON」のフォルダまでを入力する  
    <img width="497" height="169" alt="image" src="https://github.com/user-attachments/assets/bce330f6-f1e9-4d2e-a607-c472f6496a38" />
