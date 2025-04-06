import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
import pandas as pd
import win32com.client as win32
import traceback

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class AutoWordTableApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Autowordtable - --Office Word自动填表助手 by Dukeway Zhong----开源免费软件")
        self.geometry("720x520")

        self.knowledge_path = ctk.StringVar()
        self.word_path = ctk.StringVar()

        self.create_widgets()

    def create_widgets(self):
        ctk.CTkLabel(self, text="📄 选择知识库 (Excel)").pack(pady=(20, 5))
        frame1 = ctk.CTkFrame(self)
        frame1.pack(pady=5, fill="x", padx=20)
        ctk.CTkEntry(frame1, textvariable=self.knowledge_path, width=500).pack(side="left", padx=5)
        ctk.CTkButton(frame1, text="选择", command=self.browse_knowledge).pack(side="left")

        ctk.CTkLabel(self, text="📄 选择Word模板 (docx)").pack(pady=(20, 5))
        frame2 = ctk.CTkFrame(self)
        frame2.pack(pady=5, fill="x", padx=20)
        ctk.CTkEntry(frame2, textvariable=self.word_path, width=500).pack(side="left", padx=5)
        ctk.CTkButton(frame2, text="选择", command=self.browse_word).pack(side="left")

        ctk.CTkButton(self, text="▶️ 开始自动填表", command=self.run_autofill).pack(pady=20)

        ctk.CTkLabel(self, text="📝 日志输出：").pack()
        self.logbox = ctk.CTkTextbox(self, height=250)
        self.logbox.pack(padx=20, fill="both", expand=True)
        self.log("Autowordtable 启动成功 - 作者：Dukeway Zhong\n")

    def browse_knowledge(self):
        path = filedialog.askopenfilename(filetypes=[("Excel 文件", "*.xlsx")])
        if path:
            self.knowledge_path.set(path)

    def browse_word(self):
        path = filedialog.askopenfilename(filetypes=[("Word 文件", "*.docx")])
        if path:
            self.word_path.set(path)

    def log(self, text):
        self.logbox.insert("end", text + "\n")
        self.logbox.see("end")

    def run_autofill(self):
        excel_path = self.knowledge_path.get()
        word_path = self.word_path.get()

        if not os.path.exists(excel_path) or not os.path.exists(word_path):
            messagebox.showerror("错误", "请确保已选择有效的Excel和Word文件")
            return

        try:
            self.log("🔍 加载知识库...")
            df = pd.read_excel(excel_path)
            knowledge_dict = dict(zip(df['字段'], df['字段值']))

            self.log("📝 打开Word文档...")
            word = win32.gencache.EnsureDispatch('Word.Application')
            doc = word.Documents.Open(word_path)
            word.Visible = False

            for table in doc.Tables:
                for row in range(1, table.Rows.Count + 1):
                    for col in range(1, table.Columns.Count):  # 留一个空格用于填值
                        try:
                            cell_text = table.Cell(row, col).Range.Text.strip().replace('\r', '').replace('\a', '')
                            if cell_text in knowledge_dict:
                                value = knowledge_dict[cell_text]
                                table.Cell(row, col + 1).Range.Text = str(value)
                                self.log(f"✔ 填入字段：{cell_text} → {value}")
                        except Exception as e:
                            self.log(f"⚠️ 跳过 Cell({row},{col}): {e}")

            output_path = os.path.join(os.path.dirname(word_path), "已填写表格.docx")
            doc.SaveAs(output_path)
            doc.Close()
            word.Quit()

            self.log(f"✅ 填表完成，输出文件：{output_path}")
            messagebox.showinfo("完成", f"填表完成，已保存为：\n{output_path}")
            os.startfile(os.path.dirname(output_path))

        except Exception as e:
            self.log("❌ 出错：" + str(e))
            self.log(traceback.format_exc())
            messagebox.showerror("异常", str(e))

if __name__ == '__main__':
    app = AutoWordTableApp()
    app.mainloop()
