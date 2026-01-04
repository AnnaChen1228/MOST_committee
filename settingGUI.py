from ruamel.yaml import YAML
from tkinter import Tk, filedialog, messagebox, StringVar, Label, Button, Toplevel, Listbox, MULTIPLE
from tkinter.ttk import Combobox
import os
from openpyxl import load_workbook

# 定義設定檔路徑
DEFAULT_SETTING_YAML = "./setting.yaml"

# 初始化 YAML 物件
yaml = YAML()
yaml.preserve_quotes = True

def select_and_update_project_aim(setting_data):
    """
    使用 GUI 讓使用者選擇目前執行計畫 ('產學合作' 或 '研究計畫') 並更新設定資料。
    :param setting_data: 讀取的設定資料字典
    """
    root = Tk()
    root.title("計畫目標選擇")
    root.geometry("300x150")
    
    valid_project_aim_options = ["產學合作", "研究計畫"]

    selected_aim = StringVar()
    selected_aim.set(valid_project_aim_options[0])  # 預設值

    label = Label(root, text="請選擇目前執行計畫:")
    label.pack(pady=10)

    combobox = Combobox(root, textvariable=selected_aim, values=valid_project_aim_options, state="readonly")
    combobox.pack(pady=10)

    def on_select():
        current_project_aim = selected_aim.get()
        setting_data['SOURCE']['field']['目前執行計畫'] = current_project_aim
        messagebox.showinfo("成功", f"目前執行計劃目標已經設定為: {current_project_aim}")
        root.destroy()  # 正確關閉視窗

    button = Button(root, text="確認", command=on_select)
    button.pack(pady=10)
    
    root.mainloop()

def select_the_file_update_project_name(setting_data):
    """
    使用 GUI 讓使用者選擇一個檔案，並根據目前執行計畫更新對應的檔案名稱和副檔名。
    限制檔案選擇在 ./data 資料夾下的任何子資料夾。
    :param setting_data: 讀取的設定資料字典
    """
    root = Tk()
    root.title("選擇資料名稱")
    root.geometry("300x150")
    
    current_aim = setting_data['SOURCE']['field']['目前執行計畫']

    if current_aim == "研究計畫": 
        base_directory = os.path.abspath("./data/research_proj/")
    elif current_aim == "產學合作": 
        base_directory = os.path.abspath("./data/industry_coop/")

    file_window = Toplevel(root)
    file_window.title("選擇檔案")
    file_window.geometry("400x200")

    def open_file_dialog():
        file_path = filedialog.askopenfilename(initialdir=base_directory, title=f"選擇 {current_aim} 的初始資料檔案")
        if file_path:
            file_path = os.path.abspath(file_path)
            if file_path.startswith(base_directory):
                relative_path = os.path.relpath(file_path, base_directory)
                file_name = os.path.basename(file_path)
                if current_aim == '研究計畫':
                    setting_data['SOURCE']['data']['research_proj']['研究計畫申請名冊'] = relative_path
                elif current_aim == '產學合作':
                    setting_data['SOURCE']['data']['industry_coop']['產學合作申請名冊'] = relative_path
                file_name_only = os.path.splitext(file_name)[0]
                setting_data['OUTPUT']['data']['output']['FINAL_COMMITTEE'] = file_name_only + "_推薦表統合_VBA.xlsx"
                messagebox.showinfo("成功", f"已更新檔案名稱為: {file_name}")
                file_window.destroy()  # 正確關閉視窗
                root.destroy()  # 關閉主視窗
                # 呼叫下一步驟的函數來選擇 Sheet
                select_sheet_from_excel(file_path, setting_data)
            else:
                messagebox.showerror("錯誤", "選擇的檔案不在允許的目錄內。")
        else:
            messagebox.showerror("錯誤", "未選擇任何檔案。")

    label = Label(file_window, text=f"請選擇 {current_aim} 的初始資料檔案")
    label.pack(pady=20)

    select_button = Button(file_window, text="選擇檔案", command=open_file_dialog)
    select_button.pack(pady=20)

    root.withdraw()
    file_window.mainloop()

def select_sheet_from_excel(file_path, setting_data):
    """
    讀取 Excel 檔案中的 Sheet，並讓使用者選擇要使用的 Sheet 名稱。
    更新設定資料的 '計畫SHEET' 欄位為 LIST。
    :param file_path: Excel 檔案的路徑
    :param setting_data: 讀取的設定資料字典
    """
    root = Tk()
    root.title("選擇計畫 SHEET (所有 SHEET 欄位記得統一)")
    root.geometry("300x300")

    # 讀取 Excel 檔案中的 Sheets
    workbook = load_workbook(file_path, read_only=True)
    sheets = workbook.sheetnames

    # 使用 Listbox 進行多選
    listbox = Listbox(root, selectmode=MULTIPLE)
    for sheet in sheets:
        listbox.insert('end', sheet)
    listbox.pack(pady=20)

    def on_select():
        selected_indices = listbox.curselection()
        selected_sheets = [sheets[i] for i in selected_indices]
        if selected_sheets:
            setting_data['SOURCE']['field']['計畫SHEET'] = selected_sheets
            messagebox.showinfo("成功", f"計畫 SHEET 已更新為: {', '.join(selected_sheets)}")
            root.destroy()  # 正確關閉視窗
        else:
            messagebox.showerror("錯誤", "至少選擇一個 SHEET。")

    button = Button(root, text="確認", command=on_select)
    button.pack(pady=10)

    root.mainloop()

def confirm_and_update_project_name_column(file_path, sheet_name, setting_data):
    """
    確認 Excel Sheet 中的欄位，並讓使用者選擇計畫相關的欄位。
    特別注意：此處強制「申請主持人欄位」為必填。
    """
    import openpyxl
    from tkinter import Tk, Label, Button, Listbox, MULTIPLE, messagebox, StringVar, END
    from tkinter.ttk import Combobox, Scrollbar
    from tkinter import Frame, Canvas, VERTICAL, BOTH, LEFT, RIGHT, Y

    root = Tk()
    root.title("確認計畫相關欄位 - 可捲動置中")
    root.geometry("600x600")

    workbook = openpyxl.load_workbook(file_path, read_only=True)
    sheet = workbook[sheet_name]
    columns = [cell for cell in next(sheet.iter_rows(max_row=1, values_only=True))]

    # ----- 建立主框架與 Canvas + Scrollbar -----
    main_frame = Frame(root)
    main_frame.pack(fill=BOTH, expand=1)

    my_canvas = Canvas(main_frame)
    my_canvas.pack(side=LEFT, fill=BOTH, expand=1)

    my_scrollbar = Scrollbar(main_frame, orient=VERTICAL, command=my_canvas.yview)
    my_scrollbar.pack(side=RIGHT, fill=Y)

    my_canvas.configure(yscrollcommand=my_scrollbar.set)

    # 建立一個實際裝表單內容的 Frame，先以 (0,0) anchor='nw' 方式放到 Canvas 裏
    form_frame = Frame(my_canvas)
    form_window = my_canvas.create_window((0, 0), window=form_frame, anchor="nw")

    # 透過事件綁定，讓內容可隨視窗變動而捲動、置中
    def on_configure(event):
        my_canvas.configure(scrollregion=my_canvas.bbox("all"))
        canvas_width = event.width
        form_width = form_frame.winfo_reqwidth()
        if form_width < canvas_width:
            x_offset = (canvas_width - form_width) // 2
        else:
            x_offset = 0
        my_canvas.coords(form_window, x_offset, 0)

    my_canvas.bind("<Configure>", on_configure)

    # ----- 以下放所有的表單元素到 form_frame 裏 -----
    selected_project_name_column = StringVar()
    selected_keyword_column = StringVar()
    selected_abstract_column = StringVar()
    selected_institution_column = StringVar()
    selected_lead_researcher_column = StringVar()
    selected_personal_title_column = StringVar()

    # 讀取 setting 裡的舊值 (若有)
    current_project_name_column = setting_data['SOURCE']['field'].get('計畫名稱', '')
    current_keyword_column = setting_data['SOURCE']['field'].get('中文關鍵字', '')
    current_abstract_column = setting_data['SOURCE']['field'].get('計劃摘要', '')
    current_personal_title_column = setting_data['SOURCE']['field'].get('職稱', '')

    selected_project_name_column.set(current_project_name_column if current_project_name_column in columns else "")
    selected_keyword_column.set(current_keyword_column if current_keyword_column in columns else "")
    selected_abstract_column.set(current_abstract_column if current_abstract_column in columns else "")
    selected_institution_column.set("")
    selected_lead_researcher_column.set("")
    selected_personal_title_column.set(current_personal_title_column if current_personal_title_column in columns else "")

    # 計畫名稱
    Label(form_frame, text="請選擇屬於計畫名稱的欄位:").pack(pady=5)
    project_name_combobox = Combobox(form_frame, textvariable=selected_project_name_column, 
                                     values=[""] + columns, state="readonly")
    project_name_combobox.pack(pady=5)

    # 中文關鍵字
    Label(form_frame, text="請選擇屬於中文關鍵字的欄位:").pack(pady=5)
    keyword_combobox = Combobox(form_frame, textvariable=selected_keyword_column, 
                                values=[""] + columns, state="readonly")
    keyword_combobox.pack(pady=5)
    
    # 計劃摘要
    Label(form_frame, text="請選擇屬於計劃摘要的欄位:").pack(pady=5)
    abstract_combobox = Combobox(form_frame, textvariable=selected_abstract_column, 
                                 values=[""] + columns, state="readonly")
    abstract_combobox.pack(pady=5)

    # 申請機構
    Label(form_frame, text="請選擇屬於申請機構(學校)的欄位:").pack(pady=5)
    institution_combobox = Combobox(form_frame, textvariable=selected_institution_column, 
                                    values=[""] + columns, state="readonly")
    institution_combobox.pack(pady=5)

    # (計畫)主持人 => 必填
    Label(form_frame, text="請選擇屬於(計畫)主持人的欄位(必填):").pack(pady=5)
    lead_researcher_combobox = Combobox(form_frame, textvariable=selected_lead_researcher_column, 
                                        values=[""] + columns, state="readonly")
    lead_researcher_combobox.pack(pady=5)

    # 職稱
    Label(form_frame, text="請選擇屬於職稱的欄位:").pack(pady=5)
    personal_title_combobox = Combobox(form_frame, textvariable=selected_personal_title_column, 
                                       values=[""] + columns, state="readonly")
    personal_title_combobox.pack(pady=5)

    # 計畫相關其他欄位（多選）
    Label(form_frame, text="請選擇計畫相關其他欄位 (可複選):").pack(pady=5)
    other_related_fields_listbox = Listbox(form_frame, selectmode=MULTIPLE, height=5, exportselection=False)
    for column in columns:
        other_related_fields_listbox.insert(END, column)
    other_related_fields_listbox.pack(pady=5)

    # 共同主持人（多選）
    Label(form_frame, text="請選擇共同(計畫)主持人的欄位 (可複選):").pack(pady=5)
    co_lead_researchers_listbox = Listbox(form_frame, selectmode=MULTIPLE, height=5, exportselection=False)
    for column in columns:
        co_lead_researchers_listbox.insert(END, column)
    co_lead_researchers_listbox.pack(pady=5)

    # 共同機構（多選）
    Label(form_frame, text="請選擇共同機構(學校)的欄位 (可複選):").pack(pady=5)
    co_institutions_listbox = Listbox(form_frame, selectmode=MULTIPLE, height=5, exportselection=False)
    for column in columns:
        co_institutions_listbox.insert(END, column)
    co_institutions_listbox.pack(pady=5)

    def on_select():
        # 讀取使用者選擇
        project_name_column = selected_project_name_column.get()
        keyword_column = selected_keyword_column.get()
        abstract_column = selected_abstract_column.get()
        inst_column = selected_institution_column.get()
        lead_researcher_column = selected_lead_researcher_column.get()
        personal_title_col = selected_personal_title_column.get()

        selected_indices_other = other_related_fields_listbox.curselection()
        selected_other_related = [columns[i] for i in selected_indices_other]

        selected_indices_co_lead = co_lead_researchers_listbox.curselection()
        selected_co_lead_researchers = [columns[i] for i in selected_indices_co_lead]

        selected_indices_co_inst = co_institutions_listbox.curselection()
        selected_co_institutions = [columns[i] for i in selected_indices_co_inst]

        # -- 必填檢查: 申請主持人 --
        if not lead_researcher_column:
            messagebox.showerror("錯誤", "【申請主持人欄位】不得為空，請選擇一個欄位。")
            return

        # 其他邏輯可依需求做必填檢查，例如若需要「計畫名稱」「中文關鍵字」也必填：
        # if not project_name_column:
        #     messagebox.showerror("錯誤", "【計畫名稱欄位】不得為空，請選擇一個欄位。")
        #     return

        # if not keyword_column:
        #     messagebox.showerror("錯誤", "【中文關鍵字欄位】不得為空，請選擇一個欄位。")
        #     return

        # 正常更新 setting
        setting_data['SOURCE']['field']['計畫名稱'] = project_name_column
        setting_data['SOURCE']['field']['中文關鍵字'] = keyword_column
        setting_data['SOURCE']['field']['計劃摘要'] = abstract_column
        setting_data['SOURCE']['field']['申請機構欄位名稱'] = inst_column
        setting_data['SOURCE']['field']['申請主持人欄位名稱'] = lead_researcher_column
        setting_data['SOURCE']['field']['職稱'] = personal_title_col
        setting_data['SOURCE']['field']['計畫相關其他欄位'] = selected_other_related
        setting_data['SOURCE']['field']['申請共同主持人'] = selected_co_lead_researchers
        setting_data['SOURCE']['field']['申請共同機構欄位名稱'] = selected_co_institutions

        messagebox.showinfo(
            "成功",
            f"計畫名稱: {project_name_column}\n"
            f"中文關鍵字: {keyword_column}\n"
            f"計劃摘要: {abstract_column}\n"
            f"申請機構: {inst_column}\n"
            f"主持人(必填): {lead_researcher_column}\n"
            f"職稱: {personal_title_col}\n"
            f"其他欄位: {', '.join(selected_other_related) if selected_other_related else '無'}\n"
            f"共同主持人: {', '.join(selected_co_lead_researchers) if selected_co_lead_researchers else '無'}\n"
            f"共同機構: {', '.join(selected_co_institutions) if selected_co_institutions else '無'}"
        )
        root.destroy()

    Button(form_frame, text="確認", command=on_select).pack(pady=10)

    root.mainloop()

try:
    with open(DEFAULT_SETTING_YAML, 'r', encoding='utf-8') as file:
        setting_data = yaml.load(file)

    select_and_update_project_aim(setting_data)
    with open(DEFAULT_SETTING_YAML, 'w', encoding='utf-8') as file:
        yaml.dump(setting_data, file)
        
    select_the_file_update_project_name(setting_data)
    with open(DEFAULT_SETTING_YAML, 'w', encoding='utf-8') as file:
        yaml.dump(setting_data, file)

    # 根據目前執行計畫來決定檔案路徑
    current_aim = setting_data['SOURCE']['field']['目前執行計畫']
    if current_aim == "研究計畫":
        file_path = os.path.join(os.path.abspath("./data/research_proj/"), setting_data['SOURCE']['data']['research_proj']['研究計畫申請名冊'])
    elif current_aim == "產學合作":
        file_path = os.path.join(os.path.abspath("./data/industry_coop/"), setting_data['SOURCE']['data']['industry_coop']['產學合作申請名冊'])

    # 取得選擇的計畫 SHEET 名稱
    sheet_name = setting_data['SOURCE']['field']['計畫SHEET']
    confirm_and_update_project_name_column(file_path, sheet_name[0], setting_data)

    with open(DEFAULT_SETTING_YAML, 'w', encoding='utf-8') as file:
        yaml.dump(setting_data, file)

    messagebox.showinfo("完成", "設定已成功更新並存回檔案。")

except Exception as e:
    messagebox.showerror("錯誤", f"發生錯誤，無法更新設定檔。\n{e}")
finally:
    pass