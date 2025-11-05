#  PYTHON

## autofit header and footer in excel using python

![alt text](image.png)

step1: install required packages
- positron
pip install xlwings
pip install pyinstaller

step2: prepare code

step3: 
- termial
pyinstaller --noconsole --onefile autofit_excel_head_foot.py

```python
import tkinter as tk
from tkinter import filedialog, messagebox
import xlwings as xw
import os

def run_process(companyName, planNo, planTitle, excelPath, printTitleRows, fontSize, logoPath):
    app = xw.App(visible=False)
    try:
        wbTarget = app.books.open(excelPath, read_only=False)
        for ws in wbTarget.sheets:
            ws.api.PageSetup.PrintTitleRows = f"$1:${int(printTitleRows)}"
            lastRow = ws.range("A" + str(ws.cells.last_cell.row)).end('up').row
            lastCol = ws.range((int(printTitleRows), ws.cells.last_cell.column)).end('left').column
            rng = ws.range((1, 1), (lastRow, lastCol))
            rng.api.Font.Name = "Times New Roman"
            rng.api.Font.Size = int(fontSize)
            ps = ws.api.PageSetup
            ps.LeftHeader = f'&"SimSun-ExtG,Regular"&6{companyName}'
            ps.CenterHeader = f'&"SimSun-ExtG,Regular"&6{planTitle}\n{planNo}'
            ps.CenterFooter = '&"SimSun-ExtG,Regular"&6第 &P 页，共 &N 页'
            if logoPath:
                ps.RightHeaderPicture.Filename = logoPath
                ps.RightHeader = "&G"

        wbTarget.save()
        wbTarget.close()
        app.quit()
        return True, "Finished! Please check your excel file."
    except Exception as e:
        app.quit()
        return False, f"Error: {str(e)}"

def select_file(entry, types):
    path = filedialog.askopenfilename(title="选择文件", filetypes=types)
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)

def on_run(entries):
    companyName = entries['公司名称'].get()
    planNo = entries['项目编号'].get()
    planTitle = entries['项目标题'].get()
    excelPath = entries['Excel路径'].get()
    printTitleRows = entries['表前几行标题打印'].get()
    fontSize = entries['字号'].get()
    logoPath = entries['Logo图片'].get()
    if not (companyName and planNo and planTitle and excelPath and printTitleRows and fontSize and logoPath):
        messagebox.showwarning("提示", "请填写所有参数")
        return
    ok, msg = run_process(companyName, planNo, planTitle, excelPath, printTitleRows, fontSize, logoPath)
    if ok:
        messagebox.showinfo("完成", msg)
    else:
        messagebox.showerror("错误", msg)

if __name__ == '__main__':
    root = tk.Tk()
    root.title("Excel 批量调整页眉+页脚+表头+公司logo工具 by Wenchen 20251105 v0.2")
    params = ["公司名称", "项目编号", "项目标题", "Excel路径", "表前几行标题打印", "字号", "Logo图片"]
    defaults = {
        "公司名称": "XXXX（上海）有限公司",
        "表前几行标题打印": "2",
        "字号": "10"
    }
    entries = {}
    for i, param in enumerate(params):
        tk.Label(root, text=param).grid(row=i, column=0, padx=10, pady=5, sticky='e')
        ent = tk.Entry(root, width=40)
        ent.grid(row=i, column=1, padx=10, pady=5)
        if param in defaults:
            ent.insert(0, defaults[param])
        entries[param] = ent
        # 文件选择按钮
        if param == "Excel路径":
            btn = tk.Button(
                root, text="浏览...",
                command=lambda e=ent: select_file(e, (("Excel 文件", "*.xlsx;*.xlsm"), ("所有文件", "*.*")))
            )
            btn.grid(row=i, column=2, padx=5)
        elif param == "Logo图片":
            btn = tk.Button(
                root, text="浏览...",
                command=lambda e=ent: select_file(e, (("图片文件", "*.png;*.jpg;*.bmp;*.gif"), ("所有文件", "*.*")))
            )
            btn.grid(row=i, column=2, padx=5)

    tk.Button(root, text="Run", width=18, command=lambda: on_run(entries)).grid(row=len(params), column=0, columnspan=3, pady=15)
    root.mainloop()
```
