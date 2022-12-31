"""
Date: 03/11/2022
Author: Nguyễn Thành Công
Status: HOÀN THÀNH
--
Chương trình automation cho sản phẩm UL2020_RP
- Sử dụng python + một số thư viện, viết lớp cơ bản, phương thực cơ bản
- Sử dụng tiện ích của SAP GUI "Script Recording and Playback ..." và công cụ khác mà hãng SAP cung cấp để xác định GUI ID/Control ID/property
- Bật chế độ "Script Recording lên
- Tạo một hợp đồng mẫu = tay để record
- Copy script (VBS) sinh ra ở thư mục ""
- Paste vào thân lớp cơ bản
- Chính sủa code VBS > sang code PYTHON
- Viết code để sinh một số chuỗi để tự động sinh > sau các lần chạy khác nhau: số CCCD,...
- Chạy chương trình: sau khi chạy chương trình xong > một hđ đc tạo (Request CBC Amount) thành công > người dùng tiếp tục các bước khác của
quy trình (câu phí/đóng tiền, release hđ, chạy batch job, clearing downpayment...)

- Lưu ý: hđ này là hđ cơ bản với PH = LA, vậy nếu một hđ phức tạp mà cần lặp lại nhiều lần chạy => ta cũng áp dụng phương pháp trên là OK
"""
from tkinter import *
import tkinter.ttk as ttk
from tktooltip import ToolTip

from tkinter import messagebox

import sqatest.SAP_1.ul2020_RP as ul2020_RP
import sqatest.SAP_1.ul2020_SP as ul2020_SP
import sqatest.SAP_1.ul2017 as ul2017
import sqatest.SAP_1.ea as ea
import sqatest.SAP_1.edu as edu
import sqatest.SAP_1.ulk_rp as ulk_RP
import sqatest.SAP_1.ulk_sp as ulk_SP


## Khai báo lớp main để thực thi đối tượng __main__, sử dụng window để điều khiển + gọi file SAP.exe để thực thi chương trình
if __name__ == '__main__':
    ## create a list = []: ordered and changeable. Duplicates OK
    selectProducts = ['QDW0000A0000', '9R80000A0000', 'XJ40000A0000', '3N50000A0000', '3FZ0000A0000', 'QJZ0000A0000']

    ## create a set = {}: unordered and immutable, but Add/Remove OK. No duplicates
    #productDetails = {"QDW0000A0000", "9R80000A0000", "XJ40000A0000", "3N50000A0000", "3FZ0000A0000", "QJZ0000A0000"}
    productDetails = ["QDW0000A0000", "9R80000A0000", "XJ40000A0000", "3N50000A0000", "3FZ0000A0000", "QJZ0000A0000"]
    #print(type(productDetails))  # just for test

    ## create a set = {}: unordered and immutable, but Add/Remove OK. No duplicates
    #productDescriptions = {"MBAL Universal Life 2020", "MBAL Universal Life", "Endowment Assurance",
                           #"Child Educational Endowment","Whole Of Life ULK - WULKRP", "Whole Of Life ULK - WULKSP"}
    productDescriptions = ["- MBAL Universal Life 2020", "- MBAL Universal Life", "- Endowment Assurance",
                          "- Child Educational Endowment", "- Whole Of Life ULK - WULKRP", "- Whole Of Life ULK - WULKSP"]

    ## Tuple = (): ordered and unchangeable. Duplicates Ok. Faster
    #############################################################
    ## **** Create window object ****
    window = Tk()
    window.geometry('675x390')  # Window with 2 dimensions

    ## Validation:
    def delete():
        miniWindow.destroy()

    def delete1():
        miniWindow2.destroy()

    def error():
        global miniWindow  # variable function level
        miniWindow = Toplevel(window)
        miniWindow.geometry("150x90")
        miniWindow.title("Warning!")
        Label(miniWindow, text="All fields are required!", fg="red").pack()
        Button(miniWindow, text="OK", command=delete).pack()

    def success():
        global miniWindow2  # variable function level
        miniWindow2 = Toplevel(window)
        miniWindow2.geometry("150x90")
        miniWindow2.title("Warning!")
        Label(miniWindow2, text="Input info success!", fg="green").pack()
        Button(miniWindow2, text="OK", command=delete1).pack()

    ## Submit button ~ after left click to select Product from Listbox
    def submit():
        try:
            # get value from listbox 'Life insurance product' and assign to variable 'choseProduct'
            choseProduct = displayProductCode_Entry.get(displayProductCode_Entry.curselection())
            print(f"Wow, you have selected a {choseProduct} product Life Insurance!")
            #print(type(choseProduct))  # just for test
        except:
            messagebox.showwarning("Warning", "Please, select product by Code from listbox!")

        # check product that selected from listbox:
        if choseProduct == "QDW0000A0000":  # UL2020_RP ===========> DONE 03/11/2022
            ul2020_RP.SapGui().sapLogin()
        elif choseProduct == "QDW0000A0000":  # UL2020_SP
            ul2020_SP.SapGui().sapLogin()
        elif choseProduct == "9R80000A0000":  # UL2017
            ul2017.SapGui().sapLogin()
        elif choseProduct == "XJ40000A0000":  # EA ===========> DONE 12/11/2022
            ea.SapGui().sapLogin()
        elif choseProduct == "3N50000A0000":  # EDU
            edu.SapGui().sapLogin()
        elif choseProduct == "3FZ0000A0000":  # ULK_RP
            ulk_RP.SapGui().sapLogin()
        elif choseProduct == "QJZ0000A0000":  # ULK_SP
            ulk_SP.SapGui().sapLogin()
        else:
            Exception(error())

    ## Select a product from Listbox ~ by one left click
    def selectElementProduct(event):
        selection = event.widget.curselection()
        index = selection[0]
        value = event.widget.get(index)
        result.set(value)
        print(f"Product's index: " + str(index), ' -> ', "name: " + value)

    ## Labels & heading
    heading = Label(text="SAP Easy Access", bg="gray", fg="blue", width="500", height="2", font="arial 14")
    heading.pack()
    # Labels
    applReceiptDateText = Label(text="Appl.Receipt Date(*)",)
    applicationDateText = Label(text="Application Date(*)",)
    premiumDateText = Label(text="Premium Date(*)",)
    disChannelText = Label(text="Distribution Channel(*)",)
    saleProGroupText = Label(text="Sales Product Group(*)",)
    saleProTempIDText = Label(text="Sales Product Temp.ID(*)")
    # Label display Result of selected Product
    result = StringVar()
    result1 = StringVar()

    resultLabelText = Label(text="", textvariable=result, font="arial 12", fg="green")  # bg='light gray' ## or text="result"
    # 13/11/2022 do not use
    resultLabelDetailText = Label(font="arial")  #, bg='yellow', textvariable=result1

    # place labels on 'window' Tkinter widget
    applReceiptDateText.place(x=15, y=60)
    applicationDateText.place(x=15, y=90)
    premiumDateText.place(x=15, y=120)
    disChannelText.place(x=15, y=150)
    saleProGroupText.place(x=15, y=180)
    saleProTempIDText.place(x=15, y=210)
    # Display selected Product that chose from Listbox
    resultLabelText.place(x=15, y=290)

    resultLabelDetailText.place(x=150, y=290)

    #--Sử dụng map function-- # 13/11/2022 do not use
    def concateIdAndDetailOfProduct(a, b):
        return a + b

    # 13/11/2022 do not use
    wholeString = map(concateIdAndDetailOfProduct, productDetails, productDescriptions)
    lstWholeString = list(wholeString)
    print(lstWholeString)
    #result1.set(str(productDescriptions))
    result1.set(lstWholeString) # 13/11/2022 do not use

    ## Variables
    appReceiptDate = StringVar()  # date
    applicationDate = StringVar()  # date
    premiumDate = StringVar()  # date
    disChannel = StringVar()  # channels
    saleProGroup = StringVar()  # L
    saleProTempID = StringVar()  # select products by ID/code
    disProDetail = StringVar()  # display product detail

    def displaySelectedChannels(choice):  # không dùng function này
        choice = saleProTempID.get()
        print(choice)
        # Label(window, text=choice).pack()

    def displaySelectedProducts(choice1):  # không dùng function này
        choice1 = disChannel.get()
        print(choice1)
        # Label(window, text=choice1).pack()

    ## select Channel
    selectChanels = ["Agency", "Bancassurance"]

    ## Set default value:
    disChannel.set(selectChanels[0])
    saleProTempID.set(selectProducts[0])

    ## Tex fields:
    appReceiptDate_Entry = Entry(textvariable=appReceiptDate, width=16, state=DISABLED)
    applicationDate_Entry = Entry(textvariable=applicationDate, width=16, state=DISABLED)
    premiumDate_Entry = Entry(textvariable=premiumDate, width=16, state=DISABLED)

    ## dropdown list/combobox channels
    #disChannel_Entry = Entry(textvariable=disChannel, width=18)
    dropDisChannel_Entry = OptionMenu(
        window,
        disChannel,
        *selectChanels
        #command=displaySelectedChannels
    )

    saleProGroup_Entry = Entry(textvariable=saleProGroup, width=16, state=DISABLED)  # L is default value!

    ## dropdown list/combo box products
    #saleProTempID_Entry = Entry(textvariable=saleProTempID, width=38)
    dropSaleProTempID_Entry = OptionMenu(
        window,
        saleProTempID,
        *selectProducts
        #command=displaySelectedProducts
    )
    #dropSaleProTempID_Entry.config(width=12)

    # Display product with Code
    displayProductCode_Entry = Listbox(window, height=6, width=15, font=('Times', 12))
    # Display product with Detail
    displayProductDetail_Entry = Listbox(window, height=6, width=29, font=('Times', 12))

    #### place Controls on GUI
    ## Date:
    appReceiptDate_Entry.place(x=130, y=60)  # x=15, y=100
    applicationDate_Entry.place(x=130, y=90)
    premiumDate_Entry.place(x=130, y=120)

    ## Select channels:
    #disChannel_Entry.place(x=15, y=320)
    dropDisChannel_Entry.pack(expand=True)
    dropDisChannel_Entry.place(x=145, y=145, height=28, width=100)

    ## Life insurance: L
    saleProGroup_Entry.place(x=145, y=180)

    ## Select Product IDs/Codes:
    dropSaleProTempID_Entry.pack(expand=True)
    dropSaleProTempID_Entry.place(x=155, y=210, height=28)

    ## Product with code & detail:
    # Title - label in general:
    titleLabelListbox = Label(text="Select product by Code")
    titleLabelListbox.place(x=285, y=110)
    # Listbox entry with code:
    for pro in productDetails:
        # insert element from set{productDetails} into Listbox "displayProductDetail_Entry"
        displayProductCode_Entry.insert(END, pro)
        displayProductCode_Entry.place(x=290, y=136)
    ToolTip(displayProductCode_Entry, msg="Click here to select product by Code, first!")
    # Tự động co dãn chiều cao của Listbox khi thêm bớt phần từ vào listbox
    displayProductCode_Entry.config(
        height=displayProductCode_Entry.size()
    )
    # bind the "selectElementProduct" function with the '<<ListboxSelect>>' event:
    displayProductCode_Entry.bind(
        '<<ListboxSelect>>', selectElementProduct
    )

    # Listbox entry with detail:
    for proDetail in productDescriptions:
        displayProductDetail_Entry.insert(END, proDetail)
        displayProductDetail_Entry.place(x=418, y=136)
    displayProductDetail_Entry.config(height=displayProductDetail_Entry.size())


    # **** Functions
    def backspace():
        entryBack = Entry()
        entryBack.delete(len(entryBack.get())-1, END)  # delete the last character ~ press "backspace" keyboard

    def inputInfo():
        ## date
        receiptDateInfo = appReceiptDate.get()
        receiptDateInfo = str(receiptDateInfo)

        applicationDateInfo = applicationDate.get()
        applicationDateInfo = str(applicationDateInfo)

        premiumDateInfo = premiumDate.get()
        premiumDateInfo = str(premiumDateInfo)

        ## channel
        disChannelInfo = disChannel.get()

        ## Life Insurance
        saleProGroupInfo = saleProGroup.get()
        saleProGroupInfo = str(saleProGroupInfo)

        ## Products
        saleProTempIDInfo = saleProTempID.get() #saleProTempIDInfo = saleProTempID.get()
        ## Display product infos
        displayProductDetailInfo = disProDetail.set(saleProTempIDInfo)

        # Basic validate fields:
        if receiptDateInfo == "":  # date
            error()
        elif applicationDateInfo == "":  # date
            error()
        elif premiumDateInfo == "":  # date
            error()
        elif disChannelInfo == "":  # Agency, Bancassurance
            error()
        elif saleProGroupInfo == "":  # L
            error()
        elif saleProTempIDInfo == "":  # all products_id
            error()
        elif displayProductDetailInfo == "":  # all products info
            error()
        else:
            success()
            #Label(window, text="Input infomation successfully").place(x=15, y=230)

        #### Lưu file vào ổ cứng
        ## Mở một file (chuẩn bị lưu vào ổ cứng)
        file = open("inputInfo.txt", "a", encoding="utf-8")

        ## cách 1:
        #file.write("Receipt Date: " + receiptDateInfo)
        #file.write(" Application Date: " + applicationDateInfo)
        #file.write(" Premium Date: " + premiumDateInfo)

        ## cách 2:
        file.write("\t".join([receiptDateInfo, applicationDateInfo, premiumDateInfo,
                              disChannelInfo, saleProGroupInfo, saleProTempIDInfo]) + "\n")  # xuống dòng
        # close file
        file.close()

        # display message after write info successful
        print(" You've just input date successfully!")

        # Reset value to default value after write each transaction information
        ## Date
        appReceiptDate_Entry.delete(0, END)
        applicationDate_Entry.delete(0, END)
        premiumDate_Entry.delete(0, END)

        ## Channel
        #dropDisChannel_Entry.????

        ## Life Insurance
        saleProGroup_Entry.delete(0, END)

        ## Products
        #dropSaleProTempID_Entry.set("--select a product--")????
        displayProductDetail_Entry.delete(0, END)

    # **** Add window content ****
    window.title("mBOT - Automate SAP Core using Python!")

    ## fields:
    #entry = Entry()
    #entry.config(font=('Ink Free', 100))
    #entry.config(bg='#111111')
    #entry.config(fg='#00FF00')
    #entry.config(width=10)
    #entry.config(show='*') # with password
    #entry.config(show='$') # only with $ character input
    #entry.pack(side=LEFT)

    ## button:
    #btnBackspace = Button(window,text="Backspace", command=backspace, width="25")
    #btnBackspace.pack(pady=5, side=BOTTOM)

    ## Exit
    imageNoneActive = 'C:\\Users\\Admin\Desktop\\venv-pysel-prj\\sqatest\\sqatest\\SAP_1\\images\\buton_non-active.png'
    imageActive = 'C:\\Users\\Admin\Desktop\\venv-pysel-prj\\sqatest\\sqatest\\SAP_1\\images\\buton_active.png'
    def on_leaveExit(e):
        button_image.config(file=imageNoneActive)

    def on_enterExit(e):
        btnExit.config(cursor='hand2')
        button_image.config(file=imageActive)

    ## Create policy
    def on_enter(event):
        btnCreatePolicy['activebackground'] = 'green'

    def on_leave(event):
        btnCreatePolicy['bg'] = 'light gray'

    ## Button controllers:
    #### Exit
    button_image = PhotoImage(file=imageNoneActive)
    btnExit = Button(
        window, text="Exit",
        command=lambda: window.destroy(),
        width="14",
        underline=0
    )  # bg="yellow", fg="purple",
    btnExit.pack(pady=5, side=BOTTOM)  # padding surround CANCEL button
    #btnExit.bind("<Enter>", on_enterExit)
    #btnExit.bind("<Leave>", on_leaveExit)
    ToolTip(btnExit, msg="Click to Exit tkinter application form!", follow=True)
    btnExit.place(x=545, y=345)

    """
    btnCreatePol = Button(
        window, text="Create Policy",
        command=lambda: ul2020_RP.SapGui().sapLogin(),
        width="15", underline=0
    )
    btnCreatePol.pack(pady=5, side=BOTTOM)
    btnCreatePol.bind('<Leave>', on_leave)
    btnCreatePol.bind('<Enter>', on_enter)
    btnCreatePol.place(x=540, y=310)
    ## or: btn = tk.Button(text='Create Policy', activebackground='red').pack()
    """

    #### Input information
    btnInputs = Button(
        window, text="Input info", command=inputInfo,
        width="14", underline=0, state=DISABLED
    )
    btnInputs.pack(pady=5, side=BOTTOM)
    #ToolTip(btnInputs, msg="Click to input base information!", follow=True)
    btnInputs.place(x=545, y=275)

    #### Click "Create Policy" to submit and create new policy
    btnCreatePolicy = Button(
        window, text="Create Policy", command=submit,
        underline=0, width=14, border=2
    )
    btnCreatePolicy.pack()
    ToolTip(btnCreatePolicy, msg="Click here to create new policy!", follow=True)
    btnCreatePolicy.bind("<Enter>", on_enter)
    btnCreatePolicy.bind("<Leave>", on_leave)
    btnCreatePolicy.place(x=545, y=310)

    # **** Run window loop ****
    window.mainloop()

    ################# refer tooltip
    # https://stackoverflow.com/questions/68491691/hovertip-tooltip-for-each-item-in-python-ttk-combobox