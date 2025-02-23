import os
import re
import shutil
import urllib.parse
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass

# 为显示颜色添加支持
if os.name == "nt":  # Windows系统需要启用ANSI支持
    from ctypes import windll, byref
    from ctypes.wintypes import DWORD

    kernel32 = windll.kernel32
    kernel32.GetConsoleMode.restype = DWORD
    kernel32.SetConsoleMode.argtypes = (DWORD, DWORD)

    # 获取当前控制台模式
    hStdout = kernel32.GetStdHandle(-11)
    mode = DWORD()
    kernel32.GetConsoleMode(hStdout, byref(mode))

    # 启用虚拟终端处理
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
    显示banner
    """
    print(
        '''
 _____  __  __ _____  _____  __  __  
|  __ \|  \/  |  __ \|  __ \|  \/  | 
| |__) | \  / | |  | | |__) | \  / | 
|  _  /| |\/| | |  | |  _  /| |\/| | 
| | \ \| |  | | |__| | | \ \| |  | | 
|_|  \_\_|  |_|_____/|_|  \_\_|  |_| 

   去除一些伪DRM加密的EPUB电子书的工具
'''
    )


def copy_with_time(filename, date_time, new_zip, file_content, encode=""):
    """
    指定时间复制zip
    """
    new_info = zipfile.ZipInfo(filename)
    new_info.date_time = date_time
    if encode:
        new_zip.writestr(new_info, file_content.encode(encode))
    else:
        new_zip.writestr(new_info, file_content)


def parse_xhtml():
    """
    构建映射
    """
    print(f"[{Color.yellow}*{Color.reset}] 开始解析文件")
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
                    if item_id != "toc" and os.path.splitext(item_id)[1] == "":  # 补全文件名,toc那个是为了避免部分对文件名大小写不敏感的软件无法识别
                        item_id = item_id + os.path.splitext(os.path.basename(item_href))[1]
                    items[item_href] = item_id
    print(f"[{Color.green}+{Color.reset}] 解析文件成功\n")
    return items


def rename_files_in_zip(items):
    """
    重命名文件
    """
    print(f"[{Color.yellow}*{Color.reset}] 开始处理文件名")
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
    print(f"[{Color.green}+{Color.reset}] 处理成功\n")


def is_text_file(zipname, file):
    """
    判断文件是否为文本文件
    直接读取文件内容为 bytes 对象，尝试将读取的数据解码为UTF-8
    """
    try:
        raw_data = zipname.read(file)[:38]  # 这样的方式可能不够准确，为了准确性高一点分别截取了38和1024
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
    修改内部文件的引用
    """
    print(f"[{Color.yellow}*{Color.reset}] 开始修改内部引用")
    new_dic = {os.path.basename(k): items[k] for k in items.keys()}
    pattern = re.compile(r'(?:%[0-9A-Fa-f]{2})+(?:\.[A-Za-z0-9]+)?')
    with zipfile.ZipFile("./cache/output.zip", "r") as original_zip:
        with zipfile.ZipFile("./cache/output2.zip", "w") as new_zip:
            for item in original_zip.infolist():
                if item.filename[:5] == "OEBPS" and is_text_file(original_zip, item):  # 只有OEBPS目录下的才是和内容相关的
                    file_content = original_zip.read(item.filename).decode("utf-8")
                    matches = pattern.findall(file_content)
                    for match in matches:
                        if match in new_dic:
                            file_content = file_content.replace(match, new_dic[match])
                    copy_with_time(item.filename, item.date_time, new_zip, file_content, encode="utf-8")
                else:
                    copy_with_time(item.filename, item.date_time, new_zip, original_zip.read(item.filename))
    print(f"[{Color.green}+{Color.reset}] 修改成功\n")


def remove_encryption():
    """
    相关META-INF中的内容，删除有关加密的xml
    """
    print(f"[{Color.yellow}*{Color.reset}] 开始删除加密信息")
    with zipfile.ZipFile("./cache/output2.zip", "r") as original_zip:
        with zipfile.ZipFile("./cache/output3.zip", "w") as new_zip:
            for item in original_zip.infolist():
                if item.filename != "META-INF/encryption.xml":
                    copy_with_time(item.filename, item.date_time, new_zip, original_zip.read(item.filename))
    print(f"[{Color.green}+{Color.reset}] 处理成功\n")


def check_toc():
    """
    经测试可能会出现目录无法正确跳转的原因，所以这里修复一下
    """
    print(f"[{Color.yellow}*{Color.reset}] 开始自检目录")
    pattern = re.compile(r'(?:%[0-9A-Fa-f]{2})+(?:\.[A-Za-z0-9]+)?')
    with zipfile.ZipFile("./cache/output3.zip", "r") as original_zip:
        file_content = original_zip.read("OEBPS/Text/TOC.xhtml").decode("utf-8")
        matches = pattern.findall(file_content)
        if matches:
            print("    开始修复")
            toc_dic = {}
            with original_zip.open("OEBPS/Text/TOC.xhtml") as f:
                content = f.read()
                root = ET.fromstring(content)
                namespaces = {"ns": root.tag.split("}")[0].strip("{")} if "}" in root.tag else {}
                # 遍历所有的<div>标签
                for div in root.findall(".//ns:div", namespaces):
                    a = div.find("ns:a", namespaces)
                    p = a.find("ns:p", namespaces)
                    match = re.search(pattern, a.get("href"))
                    if match:
                        toc_dic[match[0]] = p.text

            pattern = "chapter\d+.xhtml"  # 正常的话标准命名通常是以chapter开头
            real_file = {}
            for item in original_zip.infolist():
                match = re.search(pattern, item.filename)
                if match:
                    with original_zip.open(item.filename) as f:
                        content = f.read()
                        root = ET.fromstring(content)
                        namespaces = {"ns": root.tag.split("}")[0].strip("{")} if "}" in root.tag else {}
                        for num in range(1, 5):  # 由于不知道标题是h几，所以从1到5都试一下
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
                print(f"[{Color.red}-{Color.reset}] 修复失败")
            else:
                print(f"[{Color.green}+{Color.reset}] 修复完成")
        else:
            print("    目录无问题")
        print(f"[{Color.green}+{Color.reset}] 目录自检结束\n")
        return


def self_check():
    """
    自检是否修改完全，可能存在无法匹配到的命名
    """
    print(f"[{Color.yellow}*{Color.reset}] 开始自检\n")
    dic_match = {
        "css": "样式文件，不影响阅读",
        "xhtml": f"{Color.red}书籍内容文件，会影响阅读{Color.reset}",
        "opf": f"{Color.red}书籍属性相关文件，影响书籍打开{Color.reset}",
        "js": "js代码(多用于注解)，不影响阅读",
        "ncx": f"{Color.yellow}与目录相关，不影响阅读，但可能导致目录无法正确跳转{Color.reset}",
        "ttf": "字体文件，不影响阅读",
        "png": f"{Color.yellow}图片文件，会导致部分图片无法正常显示{Color.reset}",
        "jpg": f"{Color.yellow}图片文件，会导致部分图片无法正常显示{Color.reset}",
        "jpeg": f"{Color.yellow}图片文件，会导致部分图片无法正常显示{Color.reset}",
        "webp": f"{Color.yellow}图片文件，会导致部分图片无法正常显示{Color.reset}",
    }
    pattern = re.compile(r'(?:%[0-9A-Fa-f]{2})+(?:\.[A-Za-z0-9]+)?')
    with zipfile.ZipFile("./cache/output4.zip", "r") as zip:
        # 遍历原始 ZIP 文件中的所有文件
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
                        f"    在{Color.yellow}{name}{Color.reset}中有{Color.yellow}{dic[k]}{Color.reset}项引用的{Color.yellow}{k}{Color.reset}文件未匹配成功")
                    print(f"    {name}为{dic_match[name.split('.')[1]]}")
                    print(f"    未能匹配到的引用的文件为{dic_match[k]}", end="\n\n")
    print(f"[{Color.green}+{Color.reset}] 自检完成\n")


def main():
    """
    主函数
    """
    print_banner()
    epub_path=input("输入EPUB路径或者直接拖动EPUB文件到窗口:")
    epub_name = os.path.basename(epub_path)
    os.makedirs("./cache", exist_ok=True)
    shutil.copy2(epub_path, "./cache/input.zip")
    items = parse_xhtml()
    if not items:
        print(f"[{Color.red}-{Color.reset}] 无法识别加密信息，可能不存在伪DRM加密")
        a=input("按任意键退出")
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
    a = input(f"转换完成，按任意键退出")


if __name__ == "__main__":
    main()
