import os
import re
import shutil
import urllib.parse
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass

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


def print_banner():
    """
    Display banner
    """
    print(
        '''
 _____  __  __ _____  _____  __  __  
|  __ \|  \/  |  __ \|  __ \|  \/  | 
| |__) | \  / | |  | | |__) | \  / | 
|  _  /| |\/| | |  | |  _  /| |\/| | 
| | \ \| |  | | |__| | | \ \| |  | | 
|_|  \_\_|  |_|_____/|_|  \_\_|  |_| 

   A tool to remove fake DRM encryption from EPUB ebooks
'''
    )


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


def parse_xhtml():
    """
    Build mapping
    """
    print(f"[{Color.yellow}*{Color.reset}] Starting file parsing")
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
                    if item_id != "toc" and os.path.splitext(item_id)[1] == "":  # Complete filename, 'toc' is to avoid case-insensitive software issues
                        item_id = item_id + os.path.splitext(os.path.basename(item_href))[1]
                    items[item_href] = item_id
    print(f"[{Color.green}+{Color.reset}] File parsing successful\n")
    return items


def rename_files_in_zip(items):
    """
    Rename files
    """
    print(f"[{Color.yellow}*{Color.reset}] Starting file renaming")
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
    print(f"[{Color.green}+{Color.reset}] Renaming successful\n")


def is_text_file(zipname, file):
    """
    Determine if a file is a text file
    Directly read the file content as bytes and try to decode it as UTF-8
    """
    try:
        raw_data = zipname.read(file)[:38]  # This method may not be accurate enough, to improve accuracy, separately cut 38 and 1024
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
    Modify internal file references
    """
    print(f"[{Color.yellow}*{Color.reset}] Starting to modify internal references")
    new_dic = {os.path.basename(k): items[k] for k in items.keys()}
    pattern = re.compile(r'(?:%[0-9A-Fa-f]{2})+(?:\.[A-Za-z0-9]+)?')
    with zipfile.ZipFile("./cache/output.zip", "r") as original_zip:
        with zipfile.ZipFile("./cache/output2.zip", "w") as new_zip:
            for item in original_zip.infolist():
                if item.filename[:5] == "OEBPS" and is_text_file(original_zip, item):  # Only files under OEBPS directory are content-related
                    file_content = original_zip.read(item.filename).decode("utf-8")
                    matches = pattern.findall(file_content)
                    for match in matches:
                        if match in new_dic:
                            file_content = file_content.replace(match, new_dic[match])
                    copy_with_time(item.filename, item.date_time, new_zip, file_content, encode="utf-8")
                else:
                    copy_with_time(item.filename, item.date_time, new_zip, original_zip.read(item.filename))
    print(f"[{Color.green}+{Color.reset}] Modification successful\n")


def remove_encryption():
    """
    Remove encryption-related XML in META-INF
    """
    print(f"[{Color.yellow}*{Color.reset}] Starting to remove encryption information")
    with zipfile.ZipFile("./cache/output2.zip", "r") as original_zip:
        with zipfile.ZipFile("./cache/output3.zip", "w") as new_zip:
            for item in original_zip.infolist():
                if item.filename != "META-INF/encryption.xml":
                    copy_with_time(item.filename, item.date_time, new_zip, original_zip.read(item.filename))
    print(f"[{Color.green}+{Color.reset}] Removal successful\n")


def check_toc():
    """
    Fix potential TOC navigation issues
    """
    print(f"[{Color.yellow}*{Color.reset}] Starting TOC self-check")
    pattern = re.compile(r'(?:%[0-9A-Fa-f]{2})+(?:\.[A-Za-z0-9]+)?')
    with zipfile.ZipFile("./cache/output3.zip", "r") as original_zip:
        file_content = original_zip.read("OEBPS/Text/TOC.xhtml").decode("utf-8")
        matches = pattern.findall(file_content)
        if matches:
            print("    Starting fix")
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

            pattern = "chapter\d+.xhtml"  # Normally, standard naming starts with 'chapter'
            real_file = {}
            for item in original_zip.infolist():
                match = re.search(pattern, item.filename)
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
                print(f"[{Color.red}-{Color.reset}] Fix failed")
            else:
                print(f"[{Color.green}+{Color.reset}] Fix completed")
        else:
            print("    TOC is fine")
        print(f"[{Color.green}+{Color.reset}] TOC self-check completed\n")
        return


def self_check():
    """
    Self-check if modifications are complete, there may be unmatched names
    """
    print(f"[{Color.yellow}*{Color.reset}] Starting self-check\n")
    dic_match = {
        "css": "Style file, does not affect reading",
        "xhtml": f"{Color.red}Content file, affects reading{Color.reset}",
        "opf": f"{Color.red}Metadata file, affects book opening{Color.reset}",
        "js": "JS code (mostly for annotations), does not affect reading",
        "ncx": f"{Color.yellow}Related to TOC, does not affect reading, but may cause navigation issues{Color.reset}",
        "ttf": "Font file, does not affect reading",
        "png": f"{Color.yellow}Image file, may cause some images to not display properly{Color.reset}",
        "jpg": f"{Color.yellow}Image file, may cause some images to not display properly{Color.reset}",
        "jpeg": f"{Color.yellow}Image file, may cause some images to not display properly{Color.reset}",
        "webp": f"{Color.yellow}Image file, may cause some images to not display properly{Color.reset}",
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
                    print(
                        f"    In {Color.yellow}{name}{Color.reset}, there are {Color.yellow}{dic[k]}{Color.reset} references to {Color.yellow}{k}{Color.reset} files that failed to match")
                    print(f"    {name} is a {dic_match[name.split('.')[1]]}")
                    print(f"    Unmatched references are {dic_match[k]}", end="\n\n")
    print(f"[{Color.green}+{Color.reset}] Self-check completed\n")


def main():
    """
    Main function
    """
    print_banner()
    epub_path=input("Enter EPUB path or drag EPUB file to the window:")
    epub_name = os.path.basename(epub_path)
    os.makedirs("./cache", exist_ok=True)
    shutil.copy2(epub_path, "./cache/input.zip")
    items = parse_xhtml()
    if not items:
        print(f"[{Color.red}-{Color.reset}] Unable to identify encryption, possibly no fake DRM encryption")
        a=input("Press any key to exit")
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
    a = input(f"Conversion completed, press any key to exit")

if __name__ == "__main__":
    main()