import tkinter as tk

print("test")

class LazyYou(tk.Tk):
    excel_name = str()
    hwp_name = str()
    
    def __init__(self):
        super().__init__()
        self.title(" LazyYou")
        self.geometry("400x260+50+50")
        self.iconbitmap('./.ico/me.ico')

        text = tk.Label(self, text="이 프로그램은 정해진 위치에 텍스트를 \n자동으로 첨부해주는 자동문서화 프로그램입니다.\n")
        btn1 = tk.Button(self, text="한글 양식(.hwp) 선택", command=self.get_hwp, width=30)
        btn2 = tk.Button(self, text="입력 양식(.xlsx) 선택", command=self.get_excel, width=30)
        btn3 = tk.Button(self, text="실행하기", command=self.start, width=30, height=2)
        text2 = tk.Label(self, text="\n만든 사람 : https://github.com/mb5ss95", height=2)

        text.pack(pady="10")
        btn1.pack(side="top", pady="5")
        btn2.pack(side="top", pady="5")
        btn3.pack(side="top", pady="5")
        text2.pack(side="top", pady="10")

    def get_hwp(self):
        from tkinter.filedialog import askopenfilenames
        from tkinter import messagebox

        hwpFile = askopenfilenames(title='이미지를 삽입할 파일을 선택하세요', filetypes=[
            ('모든 문서 파일', '*.hwp'),
            ('한글 파일 (.hwp)', '*.hwp')])

        try:
            self.hwp_name = hwpFile[0]
            messagebox.showinfo(title="선택한 한글 파일", message=self.hwp_name+"\n")
        except IndexError:
            messagebox.showerror(    "메시지 알림", "  파일을 선택하세요!")
            return

    def get_excel(self):
        from tkinter.filedialog import askopenfilenames
        from tkinter import messagebox

        excelFile = askopenfilenames(title='이미지를 삽입할 파일을 선택하세요', filetypes=[
            ('모든 문서 파일', '*.xlsx *.xlsm'),
            ('엑셀 파일 (.xlsx .xlsm)', '*.xlsx *.xlsm')])

        try:
            self.excel_name = excelFile[0]
            messagebox.showinfo(title="선택한 엑셀 파일", message=self.excel_name+"\n")
        except IndexError:
            messagebox.showerror(    "메시지 알림", "  파일을 선택하세요!")
            return

    def get_data(self):
        import pandas as pd 

        data = pd.read_excel(self.excel_name, engine = "openpyxl")
        return data.values

    def hwp_find_insert(self, hwp, msg, newMsg):
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
            
    def start(self):
        import win32com.client as win32
        import os
        from tkinter import messagebox

        df = self.get_data()
        
        
        basename = os.path.basename(self.hwp_name)
        path = os.path.dirname(self.hwp_name)

        for dfIndex, datas in enumerate(df):
            hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
            hwp.Open(self.hwp_name, "HWP", "forceopen:true")
            name="({0}) {1}".format(str(dfIndex+1),basename)
            print("index : ", str(dfIndex+1), ", ", datas)
            for dataIndex, data in enumerate(datas):
                if dataIndex == 0:
                    name = data + ".hwp"
                else:
                    self.hwp_find_insert(hwp, msg="$"+str(dataIndex)+"$", newMsg=data)
            print("new file path : ", os.path.join(path, name))
            hwp.SaveAs(os.path.join(path, name))
            hwp.Quit()

        messagebox.showinfo(title="완료!!!", message="모두 완료 되었습니다!")
        self.destroy()


if __name__ == '__main__':
    lazyYou = LazyYou()
    lazyYou.mainloop()