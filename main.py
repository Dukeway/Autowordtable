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
        self.title("Autowordtable - 个性化自动填表软件by Dukeway@qq.com (开源免费，请勿滥用)")
        self.geometry("700x500")

        self.knowledge_path = ctk.StringVar()
        self.word_path = ctk.StringVar()
        self.enable_fuzzy = ctk.BooleanVar(value=False)
        self.enable_highlight = ctk.BooleanVar(value=False)
        self.ignore_spaces = ctk.BooleanVar(value=False)

        self.create_widgets()

    def create_widgets(self):
        ctk.CTkLabel(self, text="知识库 (Excel)：").pack(pady=(20, 5))
        ctk.CTkEntry(self, textvariable=self.knowledge_path, width=500).pack()
        ctk.CTkButton(self, text="选择知识库文件", command=self.select_knowledge).pack(pady=5)

        ctk.CTkLabel(self, text="待填写表格 (Word)：").pack(pady=(20, 5))
        ctk.CTkEntry(self, textvariable=self.word_path, width=500).pack()
        ctk.CTkButton(self, text="选择Word文件", command=self.select_word).pack(pady=5)

        ctk.CTkCheckBox(self, text="启用字段模糊匹配", variable=self.enable_fuzzy).pack(pady=(10, 0))
        ctk.CTkCheckBox(self, text="填写后字体红色标记字段值", variable=self.enable_highlight).pack(pady=(5, 0))
        ctk.CTkCheckBox(self, text="字段忽略空格匹配", variable=self.ignore_spaces).pack(pady=(5, 10))

        ctk.CTkButton(self, text="开始自动填表", command=self.run_filling_thread).pack(pady=10)

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
        if "字段" not in df.columns or "字段值" not in df.columns:
            self.log("❌ 错误：Excel 中必须包含 '字段' 和 '字段值' 两列。")
            return

        def normalize(text):
            return re.sub(r"\s+", "", text) if self.ignore_spaces.get() else text

        fields = {normalize(str(k)): str(v) for k, v in zip(df["字段"], df["字段值"])}

        pythoncom.CoInitialize()
        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False
        doc = word.Documents.Open(word_path)

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
                                cell_range = table.Cell(row, target_col).Range
                                cell_range.Text = value
                                if self.enable_highlight.get():
                                    cell_range.Font.Color = win32.constants.wdColorRed
                                self.log(f"✅ 填写 '{match_key}' -> 第({row},{target_col}) 单元格: {value}")
                            except Exception as e:
                                self.log(f"⚠️ 无法填写 ({row},{target_col}): {str(e)}")
                    except Exception:
                        continue

        output_path = os.path.join(os.path.dirname(word_path), "已填写表格.docx")
        doc.SaveAs(output_path)
        doc.Close()
        word.Quit()

        self.log("✅ 填表完成，输出文件：" + output_path)

if __name__ == "__main__":
    app = AutoWordTableApp()
    app.mainloop()
