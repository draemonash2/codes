"""
指定 .drawio ファイル内のページを GUI で選択し、
選択以外を削除して「<選択名>.drawio」として保存する。
"""

from __future__ import annotations

import os
import gzip
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import List

import tkinter as tk
from tkinter import messagebox


def _read_raw_xml(path: Path) -> str:
    if zipfile.is_zipfile(path):
        with zipfile.ZipFile(path) as zf:
            for info in zf.infolist():
                if info.filename.endswith((".xml", ".drawio")):
                    return zf.read(info).decode("utf-8")
        raise ValueError("ZIP 内に XML が見つかりません。")
    else:
        with open(path, "rb") as f:
            head = f.read(2)
            f.seek(0)
            if head == b"\x1f\x8b":                  # gzip
                with gzip.open(f) as gz:
                    return gz.read().decode("utf-8")
            return f.read().decode("utf-8")


def _write_xml(path: Path, xml_text: str) -> None:
    path.write_text(xml_text, encoding="utf-8")


def get_diagram_names(root: ET.Element) -> List[str]:
    return [e.get("name", "") for e in root.iter("diagram") if e.get("name")]


def keep_only_diagram(root: ET.Element, keep_name: str) -> None:
    for elem in list(root):
        if elem.tag == "diagram" and elem.get("name") != keep_name:
            root.remove(elem)


def choose_diagram(names: List[str]) -> str | None:
    """ダイアグラム名を 1 つ選ばせて返す（キャンセル時は None）"""
    root = tk.Tk()
    root.title("drawioダイアグラム抽出")

    tk.Label(root, text="抽出するダイアグラムを選択してください:").pack(pady=(10, 0))

    lb = tk.Listbox(root, height=min(15, len(names)), width=50)
    for n in names:
        lb.insert(tk.END, n)
    lb.pack(padx=10, pady=10)

    lb.focus_set()
    lb.selection_set(0)

    chosen: list[str] = []

    def _on_ok(event=None) -> None:
        sel = lb.curselection()
        if not sel:
            messagebox.showwarning("未選択", "ダイアグラムを選択してください。")
            return
        chosen.append(lb.get(sel[0]))
        root.destroy()

    def _on_cancel() -> None:
        root.destroy()

    btn_frame = tk.Frame(root)
    btn_frame.pack(pady=(0, 10))
    tk.Button(btn_frame, text="OK", width=10, command=_on_ok).grid(row=0, column=0, padx=5)
    tk.Button(btn_frame, text="キャンセル", width=10, command=_on_cancel).grid(row=0, column=1, padx=5)

    lb.bind("<Return>", _on_ok)

    root.mainloop()
    return chosen[0] if chosen else None


def sanitize_filename(name: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', "_", name)



def extract_drawio_diagram(src_file_path: str, dst_dir_path: str) -> None:
    src_path = Path(src_file_path)

    try:
        xml_text = _read_raw_xml(src_path)
        root = ET.fromstring(xml_text)
        names = get_diagram_names(root)
        if not names:
            print("ダイアグラムが見つかりません。", file=sys.stderr)
            sys.exit(2)

        selected = choose_diagram(names)
        if not selected:
            print("キャンセルされました。")
            sys.exit(0)

        keep_only_diagram(root, selected)
        new_xml = ET.tostring(root, encoding="unicode")

        dst_name = sanitize_filename(selected) + ".drawio"
        dst_path = Path(dst_dir_path) / dst_name
        _write_xml(dst_path, new_xml)

        messagebox.showinfo("drawioダイアグラム抽出", f"ダイアグラム「{selected}」を抽出しました:\n  {dst_path}")
        print(f"Saved: {dst_path}")

    except Exception as e:
        messagebox.showerror("エラー", str(e))
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(3)


def main() -> None:
    src_file_path = os.environ.get('USERPROFILE') + "/Dropbox/100_Documents/#temp.drawio"
    dst_dir_path = os.environ.get('MYDIRPATH_DESKTOP')
    extract_drawio_diagram(src_file_path, dst_dir_path)

if __name__ == "__main__":
    main()
