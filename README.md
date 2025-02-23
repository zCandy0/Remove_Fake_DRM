# Remove Fake DRM

Engilsh | [中文](./README_zh.md) | [日本语](./README_ja.md)

This is a tool for removing **fake DRM encryption** from EPUB files.  
**Fake DRM encryption** refers to cases where the original author claims that DRM encryption is used, and the EPUB can only be opened with specific software (e.g. iBook). However, in reality, there is no true DRM encryption, instead, special characters are used o protect the EPUB and prevent it from being opened by software like Calibre, which has editing capabilities.  
**Fake DRM encryption** is not an official term,it's just how I refer to it.

---

# Usage

Download the exe file or source code from the release section and run it directly.  
If you download the source code, no third-party libraries are required, as this project uses only Python standard libraries.

![image](https://github.com/user-attachments/assets/0307c054-34ec-4497-889f-fa1b1047c475)


This tool guarantees the modification of files but maintains the modification time unchanged.

---

# Notes

+ For users who download the source code, please note that this project is developed using Python 3.11 and may not be compatible with other versions.
+ For Windows users, the tool works best in the Windows 11 terminal. If you see strange characters instead of colors, it means ANSI escape codes are not working, which might be because your Windows terminal doesn't have escape functionality enabled. However, this won't affect the functionality of the tool.

---

# Q&A

+ What is fake DRM encryption? How can I determine if my EPUB has fake DRM encryption?  
  The definition is provided earlier in the article.  
  Here are some methods to help you determine if your EPUB has fake DRM encryption ,though they may not always be accurate:
    - The original author mentions that only non-editable EPUB readers or specific software can open the file.
    - Rename the file extension to .zip , unzip it, and check the OEBPS\Text directory. If the filenames are strange characters, this might be a sign.

      ![image](https://github.com/user-attachments/assets/68271d86-25b0-4abd-9342-592cfd486799)


These are just possibilities and not definitive indicators.

+ What if self-inspection causes issues that affect reading?  
  A semi-automatic correction feature will be added later ~~(if I have time)~~. Users can work with the program to make corrections.
+ What if the tool doesn't work at all?  
  If you encounter issues with the program, please submit an issue. Don't just say it doesn't work without providing specific details.

