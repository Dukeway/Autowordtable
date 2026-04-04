import customtkinter as ctk
import tkinter.filedialog as fd
import pandas as pd
import os
import re
import pythoncom
import win32com.client as win32
import threading

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class AutoWordTableApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Autowordtable（版本V3.0） - Github: Dukeway (开源免费，请勿滥用)")
        self.geometry("700x550")

        self.knowledge_path = ctk.StringVar()
        self.word_path = ctk.StringVar()
        self.enable_fuzzy = ctk.BooleanVar(value=False)
        self.enable_highlight = ctk.BooleanVar(value=False)
        self.ignore_spaces = ctk.BooleanVar(value=False)
        self.enable_multirow = ctk.BooleanVar(value=False) 

        self.create_widgets()

    def create_widgets(self):
        ctk.CTkLabel(self, text="知识库 (Excel)：").pack(pady=(20, 5))
        ctk.CTkEntry(self, textvariable=self.knowledge_path, width=500).pack()
        ctk.CTkButton(self, text="选择知识库文件", command=self.select_knowledge,
                      fg_color="#059669", hover_color="#047857").pack(pady=5)

        ctk.CTkLabel(self, text="待填写表格 (Word)：").pack(pady=(20, 5))
        ctk.CTkEntry(self, textvariable=self.word_path, width=500).pack()
        ctk.CTkButton(self, text="选择Word文件", command=self.select_word,
                      fg_color="#059669", hover_color="#047857").pack(pady=5)

        ctk.CTkCheckBox(self, text="启用字段模糊匹配（Field fuzzy）", variable=self.enable_fuzzy).pack(pady=(10, 0))
        ctk.CTkCheckBox(self, text="填写后字体红色标记字段值（Marked in Red）", variable=self.enable_highlight).pack(pady=(5, 0))
        ctk.CTkCheckBox(self, text="字段忽略空格匹配（Matched ignoring spaces）", variable=self.ignore_spaces).pack(pady=(5, 0))
        ctk.CTkCheckBox(self, text="⚠️ Multi-row data batch filling:启用多行数据批量填充模式（注意知识库表格需要多个字段！）", variable=self.enable_multirow,
                        text_color="red").pack(pady=(5, 10))

        ctk.CTkButton(self, text="开始自动填表（Start）", command=self.run_filling_thread,
                      fg_color="#059669", hover_color="#047857").pack(pady=10)

        self.log_box = ctk.CTkTextbox(self, height=200, wrap="word", font=("Segoe UI Emoji", 12))
        self.log_box.pack(padx=10, pady=10, fill="both", expand=True)

    def log(self, message):
        self.log_box.insert("end", message + "\n")
        self.log_box.see("end")

    def select_knowledge(self):
        path = fd.askopenfilename(filetypes=[["Excel Files", "*.xlsx"]])
        if path:
            self.knowledge_path.set(path)

    def select_word(self):
        path = fd.askopenfilename(filetypes=[["Word Files", "*.docx"]])
        if path:
            self.word_path.set(path)

    def run_filling_thread(self):
        threading.Thread(target=self.fill_word_table, daemon=True).start()

    def fill_word_table(self):
        excel_path = self.knowledge_path.get()
        word_path = self.word_path.get()

        if not os.path.exists(excel_path) or not os.path.exists(word_path):
            self.log("❌ 错误：文件路径无效。")
            return

        df = pd.read_excel(excel_path)

        pythoncom.CoInitialize()
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False
        doc = word.Documents.Open(word_path)

        if self.enable_multirow.get():
            self.fill_multirow_table(df, doc)
        else:
            self.fill_single_field_table(df, doc)

        output_path = os.path.join(os.path.dirname(word_path), "已填写表格.docx")
        doc.SaveAs(output_path)
        doc.Close()
        word.Quit()

        self.log("✅ 填表完成，输出文件：" + output_path)
        os.startfile(output_path)

    def fill_single_field_table(self, df, doc):
        if "字段" not in df.columns or "字段值" not in df.columns:
            self.log("❌ 错误：Excel 中必须包含 '字段' 和 '字段值' 两列。")
            return

        def normalize(text):
            return re.sub(r"\s+", "", text) if self.ignore_spaces.get() else text

        fields = {normalize(str(k)): str(v) for k, v in zip(df["字段"], df["字段值"])}

        self.log("📄 Word 文件已打开，开始填表...")

        for table in doc.Tables:
            for row in range(1, table.Rows.Count + 1):
                for col in range(1, table.Columns.Count):
                    try:
                        cell_text = table.Cell(row, col).Range.Text.strip().replace("\r", "").replace("\x07", "")
                        norm_text = normalize(cell_text)
                        match_key = None

                        if norm_text in fields:
                            match_key = norm_text
                        elif self.enable_fuzzy.get():
                            for key in fields:
                                if key in norm_text:
                                    match_key = key
                                    break

                        if match_key:
                            value = fields[match_key]
                            target_col = col + 1 if col + 1 <= table.Columns.Count else col
                            try:
                                target_cell = table.Cell(row, target_col)
                                target_text = (
                                    target_cell.Range.Text.strip()
                                    .replace("\r", "")
                                    .replace("\x07", "")
                                )
                                if target_text != "":
                                    self.log(
                                        f"⏭️ 跳过 ({row},{target_col})：右侧已有内容，不覆盖"
                                    )
                                    continue
                                cell_range = target_cell.Range
                                cell_range.Text = value
                                if self.enable_highlight.get():
                                    cell_range.Font.Color = win32.constants.wdColorRed
                                self.log(f"✅ 填写 '{match_key}' -> 第({row},{target_col}) 单元格: {value}")
                            except Exception as e:
                                self.log(f"⚠️ 无法填写 ({row},{target_col}): {str(e)}")
                    except Exception:
                        continue

    def fill_multirow_table(self, df, doc):
        self.log("📄 批量模式启用，尝试按表格标题匹配字段...")

        for table in doc.Tables:
            if table.Rows.Count < 2:
                continue

            headers = []
            for col in range(1, table.Columns.Count + 1):
                header_text = table.Cell(1, col).Range.Text.strip().replace("\r", "").replace("\x07", "")
                headers.append(header_text)

            if not all(h in df.columns for h in headers):
                continue

            row_index = 2
            while row_index <= table.Rows.Count:
                is_empty = True
                for col in range(1, table.Columns.Count + 1):
                    cell_text = table.Cell(row_index, col).Range.Text.strip().replace("\r", "").replace("\x07", "")
                    if cell_text:
                        is_empty = False
                        break
                if is_empty:
                    table.Rows(row_index).Delete()
                else:
                    row_index += 1

            for idx, row_data in df.iterrows():
                new_row = table.Rows.Add()
                for col_idx, col_name in enumerate(headers):
                    try:
                        value = str(row_data[col_name])
                        cell_range = new_row.Cells(col_idx + 1).Range
                        cell_range.Text = value
                        if self.enable_highlight.get():
                            cell_range.Font.Color = win32.constants.wdColorRed
                        self.log(f"✅ 第{idx+1}行：{col_name} → {value}")
                    except Exception as e:
                        self.log(f"⚠️ 跳过列 {col_name}: {e}")
            break

if __name__ == "__main__":
    app = AutoWordTableApp()
    app.mainloop()
