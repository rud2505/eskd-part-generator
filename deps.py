"""Подключает библиотеки из папки lib рядом с программой (pywin32, pymupdf, openpyxl).
Так программа не зависит от того, что установлено в системном Python,
и её можно переносить на другой компьютер вместе с папкой lib."""
import os
import site
import sys

LIB = os.path.join(os.path.dirname(os.path.abspath(__file__)), "lib")
if os.path.isdir(LIB) and LIB not in sys.path:
    site.addsitedir(LIB)          # обрабатывает pywin32.pth — подключает DLL pywin32
