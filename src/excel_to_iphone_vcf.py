import argparse
import csv
import datetime as dt
import os
import re
import sys
import tkinter as tk
import unicodedata
import zipfile
from decimal import Decimal, InvalidOperation
from tkinter import filedialog, messagebox, ttk
from xml.etree import ElementTree as ET

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    DND_FILES = None
    TkBase = tk.Tk
else:
    TkBase = TkinterDnD.Tk


APP_TITLE = "Excel 转 iPhone 通讯录 VCF"
APP_VERSION = "1.1.0"
NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
VALID_STATUS = "可导出"
NUMERIC_PHONE_STATUS = "电话列需设为文本"
INVALID_PHONE_STATUS = "电话格式异常"


def local_name(tag):
    return tag.rsplit("}", 1)[-1]


def child_text(node, name):
    for child in node:
        if local_name(child.tag) == name:
            return child.text or ""
    return ""


def all_text(node):
    parts = []
    for item in node.iter():
        if local_name(item.tag) == "t" and item.text:
            parts.append(item.text)
    if parts:
        return "".join(parts)
    return "".join(node.itertext())


def column_number(cell_ref):
    letters = re.sub(r"\d+", "", cell_ref or "").upper()
    result = 0
    for char in letters:
        if "A" <= char <= "Z":
            result = result * 26 + ord(char) - ord("A") + 1
    return result


def normalize_number_text(value):
    text = str(value or "").strip()
    if not text:
        return ""

    if re.match(r"^[+-]?\d+(\.\d+)?[eE][+-]?\d+$", text):
        try:
            expanded = format(Decimal(text), "f")
            if "." in expanded:
                expanded = expanded.rstrip("0").rstrip(".")
            return expanded
        except InvalidOperation:
            return text

    if re.match(r"^\d+\.0+$", text):
        return re.sub(r"\.0+$", "", text)

    return text


def normalize_header_text(value):
    return unicodedata.normalize("NFKC", str(value or "")).strip()


def validate_headers(name_header, phone_header):
    actual_name = normalize_header_text(name_header)
    actual_phone = normalize_header_text(phone_header)
    if actual_name == "姓名" and actual_phone == "电话":
        return

    shown_name = actual_name or "（空）"
    shown_phone = actual_phone or "（空）"
    raise ValueError(
        "表头不正确。第一列必须是“姓名”，第二列必须是“电话”。\n\n"
        f"当前第一列：{shown_name}\n当前第二列：{shown_phone}"
    )


def zero_mask_width(number_format):
    text = str(number_format or "").split(";", 1)[0]
    text = re.sub(r'"[^"]*"', "", text)
    text = re.sub(r"\\.", "", text)
    text = re.sub(r"\[[^\]]*\]", "", text)
    if re.search(r"[a-zA-Z@%]", text):
        return None
    width = text.count("0")
    return width if width > 1 else None


def excel_phone_value(value, number_format=""):
    if value is None:
        return "", None

    if isinstance(value, bool):
        return str(value), INVALID_PHONE_STATUS

    if isinstance(value, (int, float, Decimal)):
        text = normalize_number_text(value)
        if not re.fullmatch(r"\d+", text):
            return text, INVALID_PHONE_STATUS

        width = zero_mask_width(number_format)
        if width is not None and len(text) <= width:
            return text.zfill(width), None
        return text, NUMERIC_PHONE_STATUS

    return str(value).strip(), None


def normalize_phone_text(value):
    text = unicodedata.normalize("NFKC", str(value or "")).strip()
    if not text:
        return "", False

    if not re.fullmatch(r"\+?[0-9\s().-]+", text):
        return text, False

    compact = re.sub(r"[\s().-]", "", text)
    if not re.fullmatch(r"\+?\d+", compact):
        return text, False

    digit_count = len(compact.lstrip("+"))
    if not 3 <= digit_count <= 20:
        return compact, False
    return compact, True


def read_zip_text(zf, name):
    try:
        with zf.open(name) as file:
            return file.read().decode("utf-8")
    except KeyError:
        return None


def first_worksheet_path(zf):
    workbook_text = read_zip_text(zf, "xl/workbook.xml")
    rels_text = read_zip_text(zf, "xl/_rels/workbook.xml.rels")
    if not workbook_text or not rels_text:
        raise ValueError("找不到 Excel 工作簿信息，请确认文件是 .xlsx 格式。")

    workbook = ET.fromstring(workbook_text)
    rels = ET.fromstring(rels_text)

    first_sheet = None
    for node in workbook.iter():
        if local_name(node.tag) == "sheet":
            first_sheet = node
            break
    if first_sheet is None:
        raise ValueError("Excel 中没有找到工作表。")

    rel_id = first_sheet.attrib.get(NS_REL + "id") or first_sheet.attrib.get("r:id")
    for rel in rels:
        if rel.attrib.get("Id") == rel_id:
            target = rel.attrib.get("Target", "")
            if target.startswith("/"):
                return target.lstrip("/")
            return "xl/" + target

    raise ValueError("无法定位第一张工作表内容。")


def read_shared_strings(zf):
    shared_text = read_zip_text(zf, "xl/sharedStrings.xml")
    if not shared_text:
        return []

    shared_xml = ET.fromstring(shared_text)
    strings = []
    for node in shared_xml:
        if local_name(node.tag) == "si":
            strings.append(all_text(node))
    return strings


def cell_value(cell, shared_strings):
    cell_type = cell.attrib.get("t", "")
    if cell_type == "inlineStr":
        for child in cell:
            if local_name(child.tag) == "is":
                return all_text(child).strip()
        return ""

    value = ""
    for child in cell:
        if local_name(child.tag) == "v":
            value = child.text or ""
            break

    if cell_type == "s":
        try:
            index = int(value)
            return shared_strings[index].strip()
        except (ValueError, IndexError):
            return ""

    return normalize_number_text(value)


def read_contacts_with_builtin_xlsx(path):
    if not zipfile.is_zipfile(path):
        raise ValueError("当前内置读取器只支持 .xlsx 文件。请另存为 .xlsx 后再试。")

    contacts = []
    with zipfile.ZipFile(path) as zf:
        shared_strings = read_shared_strings(zf)
        sheet_path = first_worksheet_path(zf)
        sheet_text = read_zip_text(zf, sheet_path)
        if not sheet_text:
            raise ValueError("无法读取第一张工作表。")

        sheet_xml = ET.fromstring(sheet_text)
        header_found = False
        for row in sheet_xml.iter():
            if local_name(row.tag) != "row":
                continue

            row_no = int(row.attrib.get("r", "0") or "0")
            name = ""
            phone = ""
            phone_issue = None
            for cell in row:
                if local_name(cell.tag) != "c":
                    continue

                col = column_number(cell.attrib.get("r", ""))
                if col == 1:
                    name = cell_value(cell, shared_strings)
                elif col == 2:
                    phone = cell_value(cell, shared_strings)
                    if phone and cell.attrib.get("t", "") in {"", "n"}:
                        phone_issue = NUMERIC_PHONE_STATUS

            if row_no == 1:
                validate_headers(name, phone)
                header_found = True
                continue

            add_contact_row(contacts, row_no, name, phone, phone_issue)

        if not header_found:
            validate_headers("", "")

    return contacts


def read_contacts_with_openpyxl(path):
    try:
        from openpyxl import load_workbook
    except ImportError:
        return None

    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook.worksheets[0]
        header = next(sheet.iter_rows(min_row=1, max_row=1, max_col=2), None)
        if header is None:
            validate_headers("", "")
        validate_headers(header[0].value, header[1].value)

        contacts = []
        for index, row in enumerate(sheet.iter_rows(min_row=2, max_col=2), start=2):
            name = "" if row[0].value is None else str(row[0].value).strip()
            phone, phone_issue = excel_phone_value(row[1].value, row[1].number_format)
            add_contact_row(contacts, index, name, phone, phone_issue)
        return contacts
    finally:
        workbook.close()


def add_contact_row(contacts, row_no, name, phone, phone_issue=None):
    name = str(name or "").strip()
    phone = str(phone or "").strip()
    if not name and not phone:
        return

    status = VALID_STATUS
    if not name:
        status = "缺少姓名"
    elif not phone:
        status = "缺少电话"
    elif phone_issue:
        status = phone_issue
    else:
        phone, is_valid = normalize_phone_text(phone)
        if not is_valid:
            status = INVALID_PHONE_STATUS

    contacts.append({
        "row": row_no,
        "name": name,
        "phone": phone,
        "status": status,
    })


def read_contacts(path):
    extension = os.path.splitext(path)[1].lower()
    if extension == ".csv":
        return read_contacts_csv(path)
    if extension != ".xlsx":
        raise ValueError("只支持 .xlsx 或 .csv 文件。")

    contacts = read_contacts_with_openpyxl(path)
    if contacts is not None:
        return contacts

    return read_contacts_with_builtin_xlsx(path)


def read_contacts_csv(path):
    contacts = []
    encodings = ["utf-8-sig", "gb18030"]
    last_error = None

    for encoding in encodings:
        try:
            with open(path, newline="", encoding=encoding) as file:
                reader = csv.reader(file)
                header = next(reader, None) or []
                validate_headers(
                    header[0] if len(header) > 0 else "",
                    header[1] if len(header) > 1 else "",
                )
                for index, row in enumerate(reader, start=2):
                    name = row[0] if len(row) > 0 else ""
                    phone = row[1] if len(row) > 1 else ""
                    add_contact_row(contacts, index, name, phone)
            return contacts
        except UnicodeDecodeError as error:
            last_error = error

    raise ValueError(f"无法读取 CSV 编码：{last_error}")


def escape_vcard_text(text):
    return (
        str(text or "")
        .replace("\\", "\\\\")
        .replace(";", "\\;")
        .replace(",", "\\,")
        .replace("\r\n", "\\n")
        .replace("\n", "\\n")
        .replace("\r", "\\n")
    )


def make_vcard(contacts):
    lines = []
    for item in contacts:
        if item["status"] != VALID_STATUS:
            continue

        name = escape_vcard_text(item["name"])
        phone = item["phone"].strip()
        lines.extend([
            "BEGIN:VCARD",
            "VERSION:3.0",
            f"N:;{name};;;",
            f"FN:{name}",
            f"TEL;TYPE=CELL,VOICE:{phone}",
            "END:VCARD",
        ])
    return "\r\n".join(lines) + "\r\n"


def export_vcf(contacts, path):
    valid = [item for item in contacts if item["status"] == VALID_STATUS]
    if not valid:
        raise ValueError("没有可导出的联系人。请确认第一列是姓名、第二列是电话，并且从第 2 行开始有数据。")

    with open(path, "w", encoding="utf-8", newline="") as file:
        file.write(make_vcard(valid))
    return len(valid)


class App(TkBase):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE} v{APP_VERSION}")
        self.geometry("860x560")
        self.minsize(760, 500)
        self.contacts = []

        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.status_var = tk.StringVar(value="请选择或拖入表头为“姓名、电话”的 Excel 文件。")

        self.build_ui()
        self.enable_file_drop()

    def build_ui(self):
        self.root_frame = ttk.Frame(self, padding=14)
        self.root_frame.pack(fill=tk.BOTH, expand=True)

        root = self.root_frame
        root.columnconfigure(1, weight=1)
        root.rowconfigure(2, weight=1)

        ttk.Label(root, text="Excel 文件").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=(0, 8))
        self.input_entry = ttk.Entry(root, textvariable=self.input_var, state="readonly")
        self.input_entry.grid(row=0, column=1, sticky="ew", pady=(0, 8))
        ttk.Button(root, text="选择...", command=self.choose_input).grid(row=0, column=2, padx=(8, 0), pady=(0, 8))

        ttk.Label(root, text="VCF 输出").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=(0, 12))
        self.output_entry = ttk.Entry(root, textvariable=self.output_var)
        self.output_entry.grid(row=1, column=1, sticky="ew", pady=(0, 12))
        ttk.Button(root, text="另存为...", command=self.choose_output).grid(row=1, column=2, padx=(8, 0), pady=(0, 12))

        columns = ("row", "name", "phone", "status")
        self.tree = ttk.Treeview(root, columns=columns, show="headings", selectmode="browse")
        self.tree.heading("row", text="行号")
        self.tree.heading("name", text="姓名")
        self.tree.heading("phone", text="电话")
        self.tree.heading("status", text="状态")
        self.tree.column("row", width=70, anchor=tk.CENTER, stretch=False)
        self.tree.column("name", width=200)
        self.tree.column("phone", width=220)
        self.tree.column("status", width=170, anchor=tk.CENTER)
        self.tree.grid(row=2, column=0, columnspan=3, sticky="nsew")

        scrollbar = ttk.Scrollbar(root, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.grid(row=2, column=3, sticky="ns")
        self.tree.configure(yscrollcommand=scrollbar.set)

        actions = ttk.Frame(root)
        actions.grid(row=3, column=0, columnspan=4, sticky="ew", pady=(12, 0))
        actions.columnconfigure(2, weight=1)

        ttk.Button(actions, text="重新预览", command=self.preview).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(actions, text="生成 VCF", command=self.export).grid(row=0, column=1, padx=(0, 14))
        ttk.Label(actions, textvariable=self.status_var).grid(row=0, column=2, sticky="w")

    def enable_file_drop(self):
        if DND_FILES is None:
            self.status_var.set("拖入文件不可用；仍可点击“选择...”导入。")
            return

        drop_targets = [
            self,
            self.root_frame,
            self.input_entry,
            self.output_entry,
            self.tree,
        ]
        for widget in drop_targets:
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", self.on_file_drop)

    def on_file_drop(self, event):
        try:
            paths = list(self.tk.splitlist(event.data))
        except tk.TclError:
            paths = [event.data]

        self.handle_dropped_files(paths)
        return "break"

    def choose_input(self):
        path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[
                ("Excel 工作簿", "*.xlsx"),
                ("CSV 文件", "*.csv"),
                ("所有文件", "*.*"),
            ],
        )
        if not path:
            return

        self.set_input_file(path)

    def set_input_file(self, path):
        self.input_var.set(path)
        folder = os.path.dirname(path)
        name = os.path.splitext(os.path.basename(path))[0]
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M")
        self.output_var.set(os.path.join(folder, f"{name}_contacts_{stamp}.vcf"))
        self.preview()

    def handle_dropped_files(self, paths):
        candidates = [
            path for path in paths
            if os.path.isfile(path) and os.path.splitext(path)[1].lower() in {".xlsx", ".csv"}
        ]
        if not candidates:
            messagebox.showinfo(APP_TITLE, "请拖入 .xlsx 或 .csv 文件。")
            return
        self.set_input_file(candidates[0])

    def choose_output(self):
        initial = self.output_var.get()
        path = filedialog.asksaveasfilename(
            title="保存 VCF 文件",
            defaultextension=".vcf",
            filetypes=[("vCard 文件", "*.vcf")],
            initialdir=os.path.dirname(initial) if initial else None,
            initialfile=os.path.basename(initial) if initial else "contacts.vcf",
        )
        if path:
            self.output_var.set(path)

    def preview(self):
        path = self.input_var.get().strip()
        if not path:
            messagebox.showinfo(APP_TITLE, "请先选择 Excel 文件。")
            return False

        try:
            self.status_var.set("正在读取 Excel...")
            self.update_idletasks()
            self.contacts = read_contacts(path)
            self.refresh_table()
            valid_count = sum(1 for item in self.contacts if item["status"] == VALID_STATUS)
            invalid_count = len(self.contacts) - valid_count
            self.status_var.set(
                f"预览完成：{valid_count} 个可导出，{invalid_count} 行需处理。"
            )
            return True
        except Exception as error:
            self.contacts = []
            self.refresh_table()
            self.status_var.set("读取失败。")
            messagebox.showerror("读取失败", str(error))
            return False

    def refresh_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for item in self.contacts:
            self.tree.insert("", tk.END, values=(item["row"], item["name"], item["phone"], item["status"]))

    def export(self):
        # 每次导出都重新读取当前文件，避免使用旧预览数据。
        if not self.preview():
            return

        path = self.output_var.get().strip()
        if not path:
            messagebox.showinfo(APP_TITLE, "请先选择 VCF 输出位置。")
            return

        invalid_count = sum(1 for item in self.contacts if item["status"] != VALID_STATUS)
        if invalid_count and not messagebox.askyesno(
            "存在未导出的行",
            f"有 {invalid_count} 行因姓名或电话问题不会导出。\n\n是否继续生成其余联系人？",
        ):
            return

        try:
            count = export_vcf(self.contacts, path)
            self.status_var.set(f"已生成：{path}")
            messagebox.showinfo("完成", f"已生成 {count} 个联系人。\n\n可以把 .vcf 发送到 iPhone，或通过 iCloud 通讯录导入。")
        except Exception as error:
            self.status_var.set("导出失败。")
            messagebox.showerror("导出失败", str(error))


def run_cli(args):
    contacts = read_contacts(args.input)
    output = args.output
    if not output:
        folder = os.path.dirname(os.path.abspath(args.input))
        name = os.path.splitext(os.path.basename(args.input))[0]
        output = os.path.join(folder, f"{name}_contacts.vcf")
    count = export_vcf(contacts, output)
    print(f"已生成 {count} 个联系人：{output}")


def main():
    parser = argparse.ArgumentParser(description="把 Excel 第一列姓名、第二列电话转换成 iPhone 可导入的 .vcf。")
    parser.add_argument("input", nargs="?", help="Excel .xlsx 或 CSV 文件")
    parser.add_argument("-o", "--output", help="输出 .vcf 路径")
    args = parser.parse_args()

    if args.input:
        run_cli(args)
    else:
        App().mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"错误：{exc}", file=sys.stderr)
        sys.exit(1)
