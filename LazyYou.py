excel_name = str()
hwp_name = str()

def get_hwp():
    from tkinter.filedialog import askopenfilenames
    from tkinter import messagebox

    global hwp_name

    hwpFile = askopenfilenames(title='이미지를 삽입할 파일을 선택하세요', filetypes=[
        ('모든 문서 파일', '*.hwp'),
        ('한글 파일 (.hwp)', '*.hwp')])

    try:
        hwp_name = hwpFile[0]
        messagebox.showinfo(title="선택한 한글 파일", message=hwp_name+"\n")
    except IndexError:
        messagebox.showerror(    "메시지 알림", "  파일을 선택하세요!")
        return

def get_excel():
    from tkinter.filedialog import askopenfilenames
    from tkinter import messagebox

    global excel_name

    excelFile = askopenfilenames(title='이미지를 삽입할 파일을 선택하세요', filetypes=[
        ('모든 문서 파일', '*.xlsx *.xlsm'),
        ('엑셀 파일 (.xlsx .xlsm)', '*.xlsx *.xlsm')])

    try:
        excel_name = excelFile[0]
        messagebox.showinfo(title="선택한 엑셀 파일", message=excel_name+"\n")
    except IndexError:
        messagebox.showerror(    "메시지 알림", "  파일을 선택하세요!")
        return

def get_data():
    import pandas as pd 

    global excel_name

    data = pd.read_excel(excel_name, engine = "openpyxl")
    return data.values

def hwp_find_insert(hwp, msg, newMsg):
    print("HWP msg, newMsg : ", msg, ", ", newMsg) 

    hwp.MovePos(2, 0, 0)
    hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)

    option = hwp.HParameterSet.HFindReplace
    option.FindString = msg
    option.ReplaceString = newMsg
    option.UseWildCards = 1
    option.IgnoreMessage = 1
    option.Direction = hwp.FindDir("Forward")
    option.FindType = False
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
          
def start():
    import pandas as pd 
    import win32com.client as win32
    import os
    from tkinter import messagebox

    global hwp_name

    df = get_data()
    
    
    basename = os.path.basename(hwp_name)
    path = os.path.dirname(hwp_name)

    for dfIndex, datas in enumerate(df):
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.Open(hwp_name, "HWP", "forceopen:true")
        name="({0}) {1}".format(str(dfIndex+1),basename)
        print("index : ", str(dfIndex+1), ", ", datas)
        for dataIndex, data in enumerate(datas):
            if dataIndex == 0:
                name = data + ".hwp"
            else:
                hwp_find_insert(hwp, msg="$"+str(dataIndex)+"$", newMsg=data)
        print("new file path : ", os.path.join(path, name))
        hwp.SaveAs(os.path.join(path, name))
        hwp.Quit()

    root.destroy()
    messagebox.showinfo(title="완료!!!", message="모두 완료 되었습니다!")

if __name__ == '__main__':
    import tkinter as tk
    
    root = tk.Tk()
    root.title(" LazyYou (v.Beta)")
    root.geometry("400x300+50+50")
    root.iconbitmap('./.ico/DSU.ico')

    text = tk.Label(root, text="이 프로그램은 정해진 위치에 텍스트를 \n자동으로 첨부해주는 자동문서화 프로그램입니다.\n")
    text.pack(pady="10")

    btn1 = tk.Button(root, text="한글 양식(.hwp) 선택", command=get_hwp, width=30)
    btn1.pack(side="top", pady="5")

    btn2 = tk.Button(root, text="입력 양식(.xlsx) 선택", command=get_excel, width=30)
    btn2.pack(side="top", pady="5")

    btn3 = tk.Button(root, text="실행하기", command=start, width=30, height=2)
    btn3.pack(side="top", pady="5")

    text2 = tk.Label(root, text="만든 사람 : 동서울대학교 MBS", height=2)
    text2.pack(side="top", pady="10")

    root.mainloop()

