import os
import re
import shutil
import urllib.parse
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass

# WindowsシステムでANSIサポートを有効にする
if os.name == "nt":  # WindowsシステムではANSIサポートを有効にする必要があります
    from ctypes import windll, byref
    from ctypes.wintypes import DWORD

    kernel32 = windll.kernel32
    kernel32.GetConsoleMode.restype = DWORD
    kernel32.SetConsoleMode.argtypes = (DWORD, DWORD)

    # 現在のコンソールモードを取得
    hStdout = kernel32.GetStdHandle(-11)
    mode = DWORD()
    kernel32.GetConsoleMode(hStdout, byref(mode))

    # 仮想端末処理を有効にする
    kernel32.SetConsoleMode(hStdout, mode.value | 0x0004)


@dataclass
class Color:
    red = "\033[91m"
    green = "\033[92m"
    yellow = "\033[93m"
    cyan = "\033[96m"
    reset = "\033[0m"


def print_banner():
    """
    バナーを表示する
    """
    print(
        '''
 _____  __  __ _____  _____  __  __  
|  __ \|  \/  |  __ \|  __ \|  \/  | 
| |__) | \  / | |  | | |__) | \  / | 
|  _  /| |\/| | |  | |  _  /| |\/| | 
| | \ \| |  | | |__| | | \ \| |  | | 
|_|  \_\_|  |_|_____/|_|  \_\_|  |_| 

   EPUB電子書籍から偽のDRM暗号化を削除するツール
'''
    )


def copy_with_time(filename, date_time, new_zip, file_content, encode=""):
    """
    指定された時間でZIPをコピーする
    """
    new_info = zipfile.ZipInfo(filename)
    new_info.date_time = date_time
    if encode:
        new_zip.writestr(new_info, file_content.encode(encode))
    else:
        new_zip.writestr(new_info, file_content)


def parse_xhtml():
    """
    マッピングを構築する
    """
    print(f"[{Color.yellow}*{Color.reset}] ファイルの解析を開始します")
    items = {}
    with zipfile.ZipFile("./cache/input.zip", "r") as z:
        with z.open("OEBPS/content.opf") as f:
            content = f.read()
            root = ET.fromstring(content)
            namespaces = {"ns": root.tag.split("}")[0].strip("{")} if "}" in root.tag else {}
            for item in root.findall(".//ns:item", namespaces):
                if "%" in item.get("href"):
                    item_id = item.get("id")
                    item_href = f"OEBPS/{item.get('href')}"
                    if item_id != "toc" and os.path.splitext(item_id)[1] == "":  # 完全なファイル名、'toc'は大文字小文字を区別しないソフトウェアの問題を避けるため
                        item_id = item_id + os.path.splitext(os.path.basename(item_href))[1]
                    items[item_href] = item_id
    print(f"[{Color.green}+{Color.reset}] ファイルの解析が成功しました\n")
    return items


def rename_files_in_zip(items):
    """
    ファイル名を変更する
    """
    print(f"[{Color.yellow}*{Color.reset}] ファイル名の変更を開始します")
    name_set = {urllib.parse.unquote(item) for item in items.keys()}
    with zipfile.ZipFile("./cache/input.zip", "r") as original_zip:
        with zipfile.ZipFile("./cache/output.zip", "w") as new_zip:
            for item in original_zip.infolist():
                if item.filename in name_set:
                    file_data = original_zip.read(item.filename)
                    dir_path = os.path.dirname(item.filename)
                    new_filename = os.path.join(dir_path, items[urllib.parse.quote(item.filename)])
                    copy_with_time(new_filename, item.date_time, new_zip, file_data)
                else:
                    file_data = original_zip.read(item.filename)
                    copy_with_time(item.filename, item.date_time, new_zip, file_data)
    print(f"[{Color.green}+{Color.reset}] ファイル名の変更が成功しました\n")


def is_text_file(zipname, file):
    """
    ファイルがテキストファイルかどうかを判定する
    ファイルの内容を直接バイトとして読み取り、UTF-8としてデコードを試みる
    """
    try:
        raw_data = zipname.read(file)[:38]  # この方法は十分に正確ではないため、精度を向上させるために38と1024を別々にカットする
        try:
            raw_data.decode("utf-8")
            return True
        except UnicodeDecodeError:
            try:
                raw_data = zipname.read(file)[:1024]
                raw_data.decode("utf-8")
                return True
            except UnicodeDecodeError:
                return False
    except IOError:
        return False


def check_file_quote(items):
    """
    内部ファイル参照を修正する
    """
    print(f"[{Color.yellow}*{Color.reset}] 内部参照の修正を開始します")
    new_dic = {os.path.basename(k): items[k] for k in items.keys()}
    pattern = re.compile(r'(?:%[0-9A-Fa-f]{2})+(?:\.[A-Za-z0-9]+)?')
    with zipfile.ZipFile("./cache/output.zip", "r") as original_zip:
        with zipfile.ZipFile("./cache/output2.zip", "w") as new_zip:
            for item in original_zip.infolist():
                if item.filename[:5] == "OEBPS" and is_text_file(original_zip, item):  # OEBPSディレクトリ下のファイルのみがコンテンツ関連
                    file_content = original_zip.read(item.filename).decode("utf-8")
                    matches = pattern.findall(file_content)
                    for match in matches:
                        if match in new_dic:
                            file_content = file_content.replace(match, new_dic[match])
                    copy_with_time(item.filename, item.date_time, new_zip, file_content, encode="utf-8")
                else:
                    copy_with_time(item.filename, item.date_time, new_zip, original_zip.read(item.filename))
    print(f"[{Color.green}+{Color.reset}] 修正が成功しました\n")


def remove_encryption():
    """
    META-INF内の暗号化関連のXMLを削除する
    """
    print(f"[{Color.yellow}*{Color.reset}] 暗号化情報の削除を開始します")
    with zipfile.ZipFile("./cache/output2.zip", "r") as original_zip:
        with zipfile.ZipFile("./cache/output3.zip", "w") as new_zip:
            for item in original_zip.infolist():
                if item.filename != "META-INF/encryption.xml":
                    copy_with_time(item.filename, item.date_time, new_zip, original_zip.read(item.filename))
    print(f"[{Color.green}+{Color.reset}] 削除が成功しました\n")


def check_toc():
    """
    TOCナビゲーションの問題を修正する
    """
    print(f"[{Color.yellow}*{Color.reset}] TOCの自己チェックを開始します")
    pattern = re.compile(r'(?:%[0-9A-Fa-f]{2})+(?:\.[A-Za-z0-9]+)?')
    with zipfile.ZipFile("./cache/output3.zip", "r") as original_zip:
        file_content = original_zip.read("OEBPS/Text/TOC.xhtml").decode("utf-8")
        matches = pattern.findall(file_content)
        if matches:
            print("    修正を開始します")
            toc_dic = {}
            with original_zip.open("OEBPS/Text/TOC.xhtml") as f:
                content = f.read()
                root = ET.fromstring(content)
                namespaces = {"ns": root.tag.split("}")[0].strip("{")} if "}" in root.tag else {}
                # すべての<div>タグを走査
                for div in root.findall(".//ns:div", namespaces):
                    a = div.find("ns:a", namespaces)
                    p = a.find("ns:p", namespaces)
                    match = re.search(pattern, a.get("href"))
                    if match:
                        toc_dic[match[0]] = p.text

            pattern = "chapter\d+.xhtml"  # 通常、標準的な命名は'chapter'で始まります
            real_file = {}
            for item in original_zip.infolist():
                match = re.search(pattern, item.filename)
                if match:
                    with original_zip.open(item.filename) as f:
                        content = f.read()
                        root = ET.fromstring(content)
                        namespaces = {"ns": root.tag.split("}")[0].strip("{")} if "}" in root.tag else {}
                        for num in range(1, 5):  # 見出しレベルが不明なため、h1からh5まで試す
                            for i in root.findall(f".//ns:h{num}", namespaces):
                                real_file[i.text] = os.path.basename(item.filename)

            with zipfile.ZipFile("./cache/output4.zip", "w") as new_zip:
                for item in original_zip.infolist():
                    if item.filename == "OEBPS/Text/TOC.xhtml":
                        for m in matches:
                            file_content = file_content.replace(m, real_file[toc_dic[m]])
                        copy_with_time(item.filename, item.date_time, new_zip, file_content, encode="utf-8")
                    else:
                        copy_with_time(item.filename, item.date_time, new_zip, original_zip.read(item.filename))
            pattern = re.compile(r'(?:%[0-9A-Fa-f]{2})+(?:\.[A-Za-z0-9]+)?')
            matches = pattern.findall(file_content)
            if matches:
                print(f"[{Color.red}-{Color.reset}] 修正に失敗しました")
            else:
                print(f"[{Color.green}+{Color.reset}] 修正が完了しました")
        else:
            print("    TOCは正常です")
        print(f"[{Color.green}+{Color.reset}] TOCの自己チェックが完了しました\n")
        return


def self_check():
    """
    修正が完全かどうかを自己チェックする、マッチしない名前が存在する可能性がある
    """
    print(f"[{Color.yellow}*{Color.reset}] 自己チェックを開始します\n")
    dic_match = {
        "css": "スタイルファイル、読み取りに影響しません",
        "xhtml": f"{Color.red}コンテンツファイル、読み取りに影響します{Color.reset}",
        "opf": f"{Color.red}メタデータファイル、書籍の開封に影響します{Color.reset}",
        "js": "JSコード（ほとんどは注釈用）、読み取りに影響しません",
        "ncx": f"{Color.yellow}TOC関連、読み取りに影響しませんが、ナビゲーションに問題を引き起こす可能性があります{Color.reset}",
        "ttf": "フォントファイル、読み取りに影響しません",
        "png": f"{Color.yellow}画像ファイル、一部の画像が正しく表示されない可能性があります{Color.reset}",
        "jpg": f"{Color.yellow}画像ファイル、一部の画像が正しく表示されない可能性があります{Color.reset}",
        "jpeg": f"{Color.yellow}画像ファイル、一部の画像が正しく表示されない可能性があります{Color.reset}",
        "webp": f"{Color.yellow}画像ファイル、一部の画像が正しく表示されない可能性があります{Color.reset}",
    }
    pattern = re.compile(r'(?:%[0-9A-Fa-f]{2})+(?:\.[A-Za-z0-9]+)?')
    with zipfile.ZipFile("./cache/output4.zip", "r") as zip:
        # オリジナルZIP内のすべてのファイルを走査
        for item in zip.infolist():
            if item.filename[:5] == "OEBPS" and is_text_file(zip, item):
                file_content = zip.read(item.filename).decode("utf-8")
                matches = pattern.findall(file_content)
                dic = {}
                name = os.path.basename(item.filename)
                for match in matches:
                    suf = match.split(".")[1]
                    if dic.get(suf):
                        dic[suf] += 1
                    else:
                        dic[suf] = 1
                for k in dic.keys():
                    print(
                        f"    {Color.yellow}{name}{Color.reset}内に、{Color.yellow}{k}{Color.reset}ファイルへの{Color.yellow}{dic[k]}{Color.reset}個の参照がマッチしませんでした")
                    print(f"    {name}は{dic_match[name.split('.')[1]]}です")
                    print(f"    マッチしなかった参照は{dic_match[k]}です", end="\n\n")
    print(f"[{Color.green}+{Color.reset}] 自己チェックが完了しました\n")


def main():
    """
    メイン関数
    """
    print_banner()
    epub_path=input("EPUBのパスを入力するか、EPUBファイルをウィンドウにドラッグしてください:")
    epub_name = os.path.basename(epub_path)
    os.makedirs("./cache", exist_ok=True)
    shutil.copy2(epub_path, "./cache/input.zip")
    items = parse_xhtml()
    if not items:
        print(f"[{Color.red}-{Color.reset}] 暗号化を識別できませんでした、偽のDRM暗号化がない可能性があります")
        a=input("任意のキーを押して終了します")
    rename_files_in_zip(items)
    check_file_quote(items)
    remove_encryption()
    check_toc()
    self_check()
    new_epub_name = f"./[fixed]{epub_name}"
    shutil.copy2("./cache/output4.zip", new_epub_name)
    stat = os.stat(epub_path)
    os.utime(new_epub_name,(stat.st_atime,stat.st_mtime))
    shutil.rmtree("./cache")
    a = input(f"変換が完了しました、任意のキーを押して終了します")

if __name__ == "__main__":
    main()