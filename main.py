import os
import re
import shutil
import urllib.parse
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
import locale
import warnings

warnings.filterwarnings("ignore", category=DeprecationWarning, message=".*locale.getdefaultlocale.*")

# Add support for displaying colors
if os.name == "nt":  # Windows system needs to enable ANSI support
    from ctypes import windll, byref
    from ctypes.wintypes import DWORD

    kernel32 = windll.kernel32
    kernel32.GetConsoleMode.restype = DWORD
    kernel32.SetConsoleMode.argtypes = (DWORD, DWORD)

    # Get current console mode
    hStdout = kernel32.GetStdHandle(-11)
    mode = DWORD()
    kernel32.GetConsoleMode(hStdout, byref(mode))

    # Enable virtual terminal processing
    kernel32.SetConsoleMode(hStdout, mode.value | 0x0004)


@dataclass
class Color:
    red = "\033[91m"
    green = "\033[92m"
    yellow = "\033[93m"
    cyan = "\033[96m"
    reset = "\033[0m"


def get_system_language():
    """
    Get system language to determine which language to use
    """
    try:
        # Try to get system language
        lang, _ = locale.getdefaultlocale()
        if lang:
            lang = lang.lower()
            if lang.startswith('zh'):
                return 'zh'
            elif lang.startswith('ja') or lang.startswith('jp'):
                return 'ja'
            else:
                return 'en'
    except:
        pass
    return 'en'  # Default to English


def get_translations():
    """
    Get translation dictionary based on system language
    """
    lang = get_system_language()

    translations = {
        'zh': {
            'banner': '''
 _____  __  __ _____  _____  __  __  
|  __ \|  \/  |  __ \|  __ \|  \/  | 
| |__) | \  / | |  | | |__) | \  / | 
|  _  /| |\/| | |  | |  _  /| |\/| | 
| | \ \| |  | | |__| | | \ \| |  | | 
|_|  \_\_|  |_|_____/|_|  \_\_|  |_| 
   去除一些伪DRM加密的EPUB电子书的工具
''',
            'input_prompt': "输入EPUB路径或者直接拖动EPUB文件到窗口:",
            'parsing_files': "开始解析文件",
            'parsing_success': "解析文件成功",
            'renaming_files': "开始处理文件名",
            'rename_success': "处理成功",
            'modifying_references': "开始修改内部引用",
            'modification_success': "修改成功",
            'removing_encryption': "开始删除加密信息",
            'removal_success': "处理成功",
            'checking_toc': "开始自检目录",
            'fixing_toc': "    开始修复",
            'toc_ok': "    目录无问题",
            'fix_failed': "修复失败",
            'fix_completed': "修复完成",
            'toc_check_complete': "目录自检结束",
            'self_check': "开始自检",
            'self_check_complete': "自检完成",
            'no_encryption': "无法识别加密信息，可能不存在伪DRM加密",
            'conversion_complete': "转换完成，按任意键退出",
            'css_desc': "样式文件，不影响阅读",
            'xhtml_desc': f"{Color.red}书籍内容文件，会影响阅读{Color.reset}",
            'opf_desc': f"{Color.red}书籍属性相关文件，影响书籍打开{Color.reset}",
            'js_desc': "js代码(多用于注解)，不影响阅读",
            'ncx_desc': f"{Color.yellow}与目录相关，不影响阅读，但可能导致目录无法正确跳转{Color.reset}",
            'ttf_desc': "字体文件，不影响阅读",
            'png_desc': f"{Color.yellow}图片文件，会导致部分图片无法正常显示{Color.reset}",
            'jpg_desc': f"{Color.yellow}图片文件，会导致部分图片无法正常显示{Color.reset}",
            'jpeg_desc': f"{Color.yellow}图片文件，会导致部分图片无法正常显示{Color.reset}",
            'webp_desc': f"{Color.yellow}图片文件，会导致部分图片无法正常显示{Color.reset}",
            'unmatched_reference': "    在{filename}中有{count}项引用的{ext}文件未匹配成功",
            'file_type_info': "    {filename}为{desc}",
            'reference_type_info': "    未能匹配到的引用的文件为{desc}"
        },
        'ja': {
            'banner': '''
 _____  __  __ _____  _____  __  __  
|  __ \|  \/  |  __ \|  __ \|  \/  | 
| |__) | \  / | |  | | |__) | \  / | 
|  _  /| |\/| | |  | |  _  /| |\/| | 
| | \ \| |  | | |__| | | \ \| |  | | 
|_|  \_\_|  |_|_____/|_|  \_\_|  |_| 

   EPUB電子書籍から偽のDRM暗号化を削除するツール
''',
            'input_prompt': "EPUBのパスを入力するか、EPUBファイルをウィンドウにドラッグしてください:",
            'parsing_files': "ファイルの解析を開始します",
            'parsing_success': "ファイルの解析が成功しました",
            'renaming_files': "ファイル名の変更を開始します",
            'rename_success': "ファイル名の変更が成功しました",
            'modifying_references': "内部参照の修正を開始します",
            'modification_success': "修正が成功しました",
            'removing_encryption': "暗号化情報の削除を開始します",
            'removal_success': "削除が成功しました",
            'checking_toc': "TOCの自己チェックを開始します",
            'fixing_toc': "    修正を開始します",
            'toc_ok': "    TOCは正常です",
            'fix_failed': "修正に失敗しました",
            'fix_completed': "修正が完了しました",
            'toc_check_complete': "TOCの自己チェックが完了しました",
            'self_check': "自己チェックを開始します",
            'self_check_complete': "自己チェックが完了しました",
            'no_encryption': "暗号化を識別できませんでした、偽のDRM暗号化がない可能性があります",
            'conversion_complete': "変換が完了しました、任意のキーを押して終了します",
            'css_desc': "スタイルファイル、読み取りに影響しません",
            'xhtml_desc': f"{Color.red}コンテンツファイル、読み取りに影響します{Color.reset}",
            'opf_desc': f"{Color.red}メタデータファイル、書籍の開封に影響します{Color.reset}",
            'js_desc': "JSコード（ほとんどは注釈用）、読み取りに影響しません",
            'ncx_desc': f"{Color.yellow}TOC関連、読み取りに影響しませんが、ナビゲーションに問題を引き起こす可能性があります{Color.reset}",
            'ttf_desc': "フォントファイル、読み取りに影響しません",
            'png_desc': f"{Color.yellow}画像ファイル、一部の画像が正しく表示されない可能性があります{Color.reset}",
            'jpg_desc': f"{Color.yellow}画像ファイル、一部の画像が正しく表示されない可能性があります{Color.reset}",
            'jpeg_desc': f"{Color.yellow}画像ファイル、一部の画像が正しく表示されない可能性があります{Color.reset}",
            'webp_desc': f"{Color.yellow}画像ファイル、一部の画像が正しく表示されない可能性があります{Color.reset}",
            'unmatched_reference': "    {filename}内に、{ext}ファイルへの{count}個の参照がマッチしませんでした",
            'file_type_info': "{filename}は{desc}です",
            'reference_type_info': "    マッチしなかった参照は{desc}です"
        },
        'en': {
            'banner': '''
 _____  __  __ _____  _____  __  __  
|  __ \|  \/  |  __ \|  __ \|  \/  | 
| |__) | \  / | |  | | |__) | \  / | 
|  _  /| |\/| | |  | |  _  /| |\/| | 
| | \ \| |  | | |__| | | \ \| |  | | 
|_|  \_\_|  |_|_____/|_|  \_\_|  |_| 

   A tool to remove fake DRM encryption from EPUB ebooks
''',
            'input_prompt': "Enter EPUB path or drag EPUB file to the window:",
            'parsing_files': "Starting file parsing",
            'parsing_success': "File parsing successful",
            'renaming_files': "Starting file renaming",
            'rename_success': "Renaming successful",
            'modifying_references': "Starting to modify internal references",
            'modification_success': "Modification successful",
            'removing_encryption': "Starting to remove encryption information",
            'removal_success': "Removal successful",
            'checking_toc': "Starting TOC self-check",
            'fixing_toc': "    Starting fix",
            'toc_ok': "    TOC is fine",
            'fix_failed': "Fix failed",
            'fix_completed': "Fix completed",
            'toc_check_complete': "TOC self-check completed",
            'self_check': "Starting self-check",
            'self_check_complete': "Self-check completed",
            'no_encryption': "Unable to identify encryption, possibly no fake DRM encryption",
            'conversion_complete': "Conversion completed, press any key to exit",
            'css_desc': "Style file, does not affect reading",
            'xhtml_desc': f"{Color.red}Content file, affects reading{Color.reset}",
            'opf_desc': f"{Color.red}Metadata file, affects book opening{Color.reset}",
            'js_desc': "JS code (mostly for annotations), does not affect reading",
            'ncx_desc': f"{Color.yellow}Related to TOC, does not affect reading, but may cause navigation issues{Color.reset}",
            'ttf_desc': "Font file, does not affect reading",
            'png_desc': f"{Color.yellow}Image file, may cause some images to not display properly{Color.reset}",
            'jpg_desc': f"{Color.yellow}Image file, may cause some images to not display properly{Color.reset}",
            'jpeg_desc': f"{Color.yellow}Image file, may cause some images to not display properly{Color.reset}",
            'webp_desc': f"{Color.yellow}Image file, may cause some images to not display properly{Color.reset}",
            'unmatched_reference': "    In {filename}, there are {count} references to {ext} files that failed to match",
            'file_type_info': "    {filename} is a {desc}",
            'reference_type_info': "    Unmatched references are {desc}"
        }
    }

    return translations.get(lang, translations['en'])


def print_banner(t):
    """
    Display banner
    """
    print(t['banner'])


def copy_with_time(filename, date_time, new_zip, file_content, encode=""):
    """
    Copy zip with specified time
    """
    new_info = zipfile.ZipInfo(filename)
    new_info.date_time = date_time
    if encode:
        new_zip.writestr(new_info, file_content.encode(encode))
    else:
        new_zip.writestr(new_info, file_content)


def parse_xhtml(t):
    """
    Build mapping
    """
    print(f"[{Color.yellow}*{Color.reset}] {t['parsing_files']}")
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
                    if item_id != "toc" and os.path.splitext(item_id)[
                        1] == "":  # Complete filename, 'toc' is to avoid case-insensitive software issues
                        item_id = item_id + os.path.splitext(os.path.basename(item_href))[1]
                    items[item_href] = item_id

    print(f"[{Color.green}+{Color.reset}] {t['parsing_success']}\n")
    return items


def rename_files_in_zip(items, t):
    """
    Rename files
    """
    print(f"[{Color.yellow}*{Color.reset}] {t['renaming_files']}")
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
    print(f"[{Color.green}+{Color.reset}] {t['rename_success']}\n")


def is_text_file(zipname, file):
    """
    Determine if a file is a text file
    Directly read the file content as bytes and try to decode it as UTF-8
    """
    try:
        raw_data = zipname.read(file)[
                   :38]  # This method may not be accurate enough, to improve accuracy, separately cut 38 and 1024
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


def check_file_quote(items, t):
    """
    Modify internal file references
    """
    print(f"[{Color.yellow}*{Color.reset}] {t['modifying_references']}")
    new_dic = {os.path.basename(k): items[k] for k in items.keys()}
    pattern = re.compile(r'(?:%[0-9A-Fa-f]{2})+(?:\.[A-Za-z0-9]+)?')
    with zipfile.ZipFile("./cache/output.zip", "r") as original_zip:
        with zipfile.ZipFile("./cache/output2.zip", "w") as new_zip:
            for item in original_zip.infolist():
                if item.filename[:5] == "OEBPS" and is_text_file(original_zip,
                                                                 item):  # Only files under OEBPS directory are content-related
                    file_content = original_zip.read(item.filename).decode("utf-8")
                    matches = pattern.findall(file_content)
                    for match in matches:
                        if match in new_dic:
                            file_content = file_content.replace(match, new_dic[match])
                    copy_with_time(item.filename, item.date_time, new_zip, file_content, encode="utf-8")
                else:
                    copy_with_time(item.filename, item.date_time, new_zip, original_zip.read(item.filename))
    print(f"[{Color.green}+{Color.reset}] {t['modification_success']}\n")


def remove_encryption(t):
    """
    Remove encryption-related XML in META-INF
    """
    print(f"[{Color.yellow}*{Color.reset}] {t['removing_encryption']}")
    with zipfile.ZipFile("./cache/output2.zip", "r") as original_zip:
        with zipfile.ZipFile("./cache/output3.zip", "w") as new_zip:
            for item in original_zip.infolist():
                if item.filename != "META-INF/encryption.xml":
                    copy_with_time(item.filename, item.date_time, new_zip, original_zip.read(item.filename))
    print(f"[{Color.green}+{Color.reset}] {t['removal_success']}\n")


def check_toc(t):
    """
    Fix potential TOC navigation issues
    """
    print(f"[{Color.yellow}*{Color.reset}] {t['checking_toc']}")
    pattern = re.compile(r'(?:%[0-9A-Fa-f]{2})+(?:\.[A-Za-z0-9]+)?')
    with zipfile.ZipFile("./cache/output3.zip", "r") as original_zip:
        file_content = original_zip.read("OEBPS/Text/TOC.xhtml").decode("utf-8")
        matches = pattern.findall(file_content)
        if matches:
            print(t['fixing_toc'])
            toc_dic = {}
            with original_zip.open("OEBPS/Text/TOC.xhtml") as f:
                content = f.read()
                root = ET.fromstring(content)
                namespaces = {"ns": root.tag.split("}")[0].strip("{")} if "}" in root.tag else {}
                # Traverse all <div> tags
                for div in root.findall(".//ns:div", namespaces):
                    a = div.find("ns:a", namespaces)
                    p = a.find("ns:p", namespaces)
                    match = re.search(pattern, a.get("href"))
                    if match:
                        toc_dic[match[0]] = p.text

            pattern_str = "chapter\d+.xhtml"  # Normally, standard naming starts with 'chapter'
            real_file = {}
            for item in original_zip.infolist():
                match = re.search(pattern_str, item.filename)
                if match:
                    with original_zip.open(item.filename) as f:
                        content = f.read()
                        root = ET.fromstring(content)
                        namespaces = {"ns": root.tag.split("}")[0].strip("{")} if "}" in root.tag else {}
                        for num in range(1, 5):  # Since the heading level is unknown, try from h1 to h5
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
                print(f"[{Color.red}-{Color.reset}] {t['fix_failed']}")
            else:
                print(f"[{Color.green}+{Color.reset}] {t['fix_completed']}")
        else:
            shutil.copy2("./cache/output3.zip", "./cache/output4.zip")
            print(t['toc_ok'])
        print(f"[{Color.green}+{Color.reset}] {t['toc_check_complete']}\n")
        return


def self_check(t):
    """
    Self-check if modifications are complete, there may be unmatched names
    """
    print(f"[{Color.yellow}*{Color.reset}] {t['self_check']}\n")
    dic_match = {
        "css": t['css_desc'],
        "xhtml": t['xhtml_desc'],
        "opf": t['opf_desc'],
        "js": t['js_desc'],
        "ncx": t['ncx_desc'],
        "ttf": t['ttf_desc'],
        "png": t['png_desc'],
        "jpg": t['jpg_desc'],
        "jpeg": t['jpeg_desc'],
        "webp": t['webp_desc'],
    }
    pattern = re.compile(r'(?:%[0-9A-Fa-f]{2})+(?:\.[A-Za-z0-9]+)?')
    with zipfile.ZipFile("./cache/output4.zip", "r") as zip:
        # Traverse all files in the original ZIP
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
                    print(t['unmatched_reference'].format(filename=Color.yellow + name + Color.reset,
                                                          count=Color.yellow + str(dic[k]) + Color.reset,
                                                          ext=Color.yellow + k + Color.reset))
                    print(t['file_type_info'].format(filename=name, desc=dic_match[name.split('.')[1]]))
                    print(t['reference_type_info'].format(desc=dic_match[k]), end="\n\n")
    print(f"[{Color.green}+{Color.reset}] {t['self_check_complete']}\n")


def main():
    """
    Main function
    """
    t = get_translations()
    print_banner(t)
    epub_path = input(t['input_prompt'])
    epub_path = epub_path.strip().strip('"\'')
    epub_name = os.path.basename(epub_path)
    os.makedirs("./cache", exist_ok=True)
    shutil.copy2(epub_path, "./cache/input.zip")
    items = parse_xhtml(t)
    if not items:  #an easy check,maybe not correct
        print(f"[{Color.red}-{Color.reset}] {t['no_encryption']}")
        shutil.rmtree("./cache")
        a = input("Press any key to exit")
    rename_files_in_zip(items, t)
    check_file_quote(items, t)
    remove_encryption(t)
    check_toc(t)
    self_check(t)
    new_epub_name = f"./[fixed]{epub_name}"
    shutil.copy2("./cache/output4.zip", new_epub_name)
    stat = os.stat(epub_path)
    os.utime(new_epub_name, (stat.st_atime, stat.st_mtime))
    shutil.rmtree("./cache")
    a = input(f"{t['conversion_complete']}")


if __name__ == "__main__":
    main()
