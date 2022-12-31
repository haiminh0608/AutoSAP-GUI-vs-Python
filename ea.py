# TODO

import win32com.client
import subprocess
import sys

from tkinter import *
from tkinter import messagebox

import time
import random

class SapGui:

    # Contractor
    def __init__(self):
        self.path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        subprocess.Popen(self.path)
        time.sleep(10)

        self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
        if not type(self.SapGuiAuto) == win32com.client.CDispatch:
            return

        application = self.SapGuiAuto.GetScriptingEngine
        self.connection = application.OpenConnection("QAS_FSPM", True)  #("QAS_ERP", True)
        time.sleep(1)

        self.session = self.connection.Children(0)
        self.session.findById("wnd[0]").maximize

    def sapLogin(self):
        try:
            self.session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "200"
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "FSOFT08"  # enter username
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "Abc123456"  # enter password
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").setFocus()  # thêm dấu () cho đúng với code PYTHON; code VBScript kg có dấu ()
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 9
            self.session.findById("wnd[0]").sendVKey(0)  # thêm dấu (0) cho đúng với code PYTHON; code VBScript kg có dấu (0)
            time.sleep(1)  # đợi 1 giây

            ## Câu hỏi bảo mật
            if self.session.findById("/app/con[0]/ses[0]/wnd[1]").text == "License Information for Multiple Logons":
                self.session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select() # thêm dấu () cho đúng với code PYTHON; code VBScript kg có dấu ()
                self.session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").setFocus() # thêm dấu () cho đúng với code PYTHON; code VBScript kg có dấu ()
                self.session.findById("wnd[1]/tbar[0]/btn[0]").press() # thêm dấu () cho đúng với code PYTHON; code VBScript kg có dấu ()
                print("You have select a second choose selection!")
            else:
                print("System is continuing now!")

            ## Tạo hợp đồng EA
            self.session.findById(
                "wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = "F00004"  # select menu to create new policy
            self.session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").doubleClickNode(
                "F00004")  # double-click on menu to open form
            time.sleep(1)

            #########################CODE NÀY ĐANG HARD CODE ĐỂ AUTOMAT#############
            #########################UL2020_RP######################################
            #### EA:
            ##### RCD: 01.01.2018
            ##### Frequency: Annually
            ##### SA: 300M
            ##### PH-LA: 01.01.1981/Z001/M/0/0
            ##### Terms: 10/10

            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPL3FCM_MAIN_DYNPROS:0330/subSECTION1:SAPLZABP_CDM_APOLICY:1004/ctxt/PM0/ABCAPOLICY-APPLIN_DT").text = "01.01.2018"
            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPL3FCM_MAIN_DYNPROS:0330/subSECTION1:SAPLZABP_CDM_APOLICY:1004/ctxt/PM0/ABCAPOLICY-APPL_DT").text = "01.01.2018"
            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPL3FCM_MAIN_DYNPROS:0330/subSECTION1:SAPLZABP_CDM_APOLICY:1004/ctxt/PM0/ABCAPOLICY-ZZ_PREM_DT").text = "01.01.2018"

            #chanelSelection = [90, 91] ## 90: Banca; 91: Agency
            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPL3FCM_MAIN_DYNPROS:0330/subSECTION1:SAPLZABP_CDM_APOLICY:1004/cmb/PM0/ABCAPOLICY-SALECH_CD").key = "91"

            # Life Insurance
            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPL3FCM_MAIN_DYNPROS:0330/subSECTION1:SAPLZABP_CDM_APOLICY:1004/ctxt/PM0/ABCAPOLICY-TEMPLATEGRP_CD").text = "L"

            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPL3FCM_MAIN_DYNPROS:0330/subSECTION1:SAPLZABP_CDM_APOLICY:1004/ctxt/PM0/ABCAPOLICY-PM_ID").setFocus()
            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPL3FCM_MAIN_DYNPROS:0330/subSECTION1:SAPLZABP_CDM_APOLICY:1004/ctxt/PM0/ABCAPOLICY-TEMPLATEGRP_CD").caretPosition = 1
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPL3FCM_MAIN_DYNPROS:0330/subSECTION1:SAPLZABP_CDM_APOLICY:1004/ctxt/PM0/ABCAPOLICY-PM_ID").setFocus()
            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPL3FCM_MAIN_DYNPROS:0330/subSECTION1:SAPLZABP_CDM_APOLICY:1004/ctxt/PM0/ABCAPOLICY-PM_ID").caretPosition = 0
            self.session.findById("wnd[0]").sendVKey(4)
            self.session.findById("wnd[1]/usr/lbl[14,14]").setFocus()
            self.session.findById("wnd[1]/usr/lbl[14,14]").caretPosition = 14
            self.session.findById("wnd[1]").sendVKey(2)
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            self.session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").topNode = "Key1"
            ##################################################################################################################################
            ###### Ext. Application No.: Giá trị trường này được copy từ field "Policy Number", do vậy khi chạy lần 2 sẽ có vấn đề
            ###### hệ thống sẽ báo lỗi "Entered external application number 190000028615 already exists" => Các cách xử lý
            ################################
            ## Cách 1: Tạo dãy số random
            #extPolicyNo = random.randrange(10000, 99999, 1)  # copyPolicy = random.randint(1,20)
            #self.session.findById(
            #    "wnd[0]/usr/subMAINSCREEN:/PM0/SAPL3FCM_MAIN_DYNPROS:0200/tabsTABSTRIP/tabpCMD_TAB1/ssubTABSUB:/PM0/SAPLABX_CDM_MAIN_DYNPROS:1003/subSECTION1:SAPLZPM_ABP_CDM_APOLICY:9000/txt/PM0/ABCAPOLICY-PAGNO_ID").text = "1900000" + str(
            #    extPolicyNo)  # "190000028615"
            ################################
            ## Cách 2: Lấy giá trị của đối tượng "Policy Number"
            #### Dùng "Scripting Trackers" tool để xác định đối tượng "Policy Number" và lấy giá trị đối tượng "Policy Number" này, dán/gán vào giá trị của đối tượng "Ext.  Application No."
            #### Lấy giá trị đối tượng "Policy Number" mà SAP vừa sinh tự động (trước đó)
            tempGetPolicyNumber = self.session.findById(
                "wnd[0]/usr/subHEADER:/PM0/SAPLABP_CDM_HEADER:1000/ctxt/PM0/ABCHEADER-POLICYNR_TT")\
                .text
            #### Gán/dán giá trị này cho giá trị đối tượng "Ext Policy No."
            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPL3FCM_MAIN_DYNPROS:0200/tabsTABSTRIP/tabpCMD_TAB1/ssubTABSUB:/PM0/SAPLABX_CDM_MAIN_DYNPROS:1003/subSECTION1:SAPLZPM_ABP_CDM_APOLICY:9000/txt/PM0/ABCAPOLICY-PAGNO_ID")\
                .text = tempGetPolicyNumber
            ####
            ##################################################################################################################################
            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPL3FCM_MAIN_DYNPROS:0200/tabsTABSTRIP/tabpCMD_TAB1/ssubTABSUB:/PM0/SAPLABX_CDM_MAIN_DYNPROS:1003/subSECTION1:SAPLZPM_ABP_CDM_APOLICY:9000/cmb/PM0/ABCAPOLICY-ZZ_PAPERPOLICY_PACK").key = "0"
            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPL3FCM_MAIN_DYNPROS:0200/tabsTABSTRIP/tabpCMD_TAB1/ssubTABSUB:/PM0/SAPLABX_CDM_MAIN_DYNPROS:1003/subSECTION1:SAPLZPM_ABP_CDM_APOLICY:9000/cmb/PM0/ABCAPOLICY-ZZ_PAYMENT_MODE").key = "1"
            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPL3FCM_MAIN_DYNPROS:0200/tabsTABSTRIP/tabpCMD_TAB1/ssubTABSUB:/PM0/SAPLABX_CDM_MAIN_DYNPROS:1003/subSECTION1:SAPLZPM_ABP_CDM_APOLICY:9000/cmb/PM0/ABCAPOLICY-ZZ_PAYMENT_MODE").setFocus()
            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPL3FCM_MAIN_DYNPROS:0200/tabsTABSTRIP/tabpCMD_TAB2").select()
            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPL3FCM_MAIN_DYNPROS:0200/tabsTABSTRIP/tabpCMD_TAB2/ssubTABSUB:/PM0/SAPLABP_CDM_APOLICY:0300/subSECTION2:/PM0/SAPLABP_CDM_ACOMMIS:1001/cntlCCR_COMMIS/shellcont/shell").pressToolbarButton("CMD_F_COMMIS_ADD_NEW")
            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPLABP_CDM_ACOMMIS:0320/subSECTION1:/PM0/SAPLABP_CDM_ACOMMIS:1010/ctxt/PM0/ABCACOMMIS-COMMCONTRNR_ID").text = "2210001689"
            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPLABP_CDM_ACOMMIS:0320/subSECTION1:/PM0/SAPLABP_CDM_ACOMMIS:1010/ctxt/PM0/ABCACOMMIS-COMMCONTRNR_ID").caretPosition = 10
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()

            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPL3FCM_MAIN_DYNPROS:0200/tabsTABSTRIP/tabpCMD_TAB1").select()
            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPL3FCM_MAIN_DYNPROS:0200/tabsTABSTRIP/tabpCMD_TAB1/ssubTABSUB:/PM0/SAPLABX_CDM_MAIN_DYNPROS:1003/subSECTION2:/PM0/SAPLABP_CDM_APOLHLDR:1400/cntlCCR_POLHLDR/shellcont/shell").pressToolbarButton(
                "CMD_F_POLHLDR_ADD")
            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPLABX_CDM_MAIN_DYNPROS:2096/subSECTION1:/PM0/SAPLABP_CDM_APOLHLDR:1001/ctxt/PM0/ABCAPOLHLDR-PARTNER_ID").text = "100242937"  # chọn tạm 01 BP trước khi tạo BP cho PH này
            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPLABX_CDM_MAIN_DYNPROS:2096/subSECTION1:/PM0/SAPLABP_CDM_APOLHLDR:1001/ctxt/PM0/ABCAPOLHLDR-PARTNER_ID").caretPosition = 9
            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPLABX_CDM_MAIN_DYNPROS:2096/subSECTION1:/PM0/SAPLABP_CDM_APOLHLDR:1001/btnCMD_F_PTR_PHD_ADD").press()
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()

            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7015/subA02P01:SAPLBUD0:1130/cmbBUS000FLDS-TITLE_MEDI").key = "0002"

            # Start Congnt modified:
            phNameExt = random.randrange(1, 100, 1)
            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7015/subA02P03:SAPLBUD0:1301/txtBUT000-NAME_FIRST").text = "PH_LA_EA_TEST" + str(
                phNameExt)  # random string để sinh chuỗi phần cuối của PH này
            # End Congnt modified

            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7015/subA02P03:SAPLBUD0:1301/txtBUT000-NAME_FIRST").setFocus()
            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_01/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7015/subA02P03:SAPLBUD0:1301/txtBUT000-NAME_FIRST").caretPosition = 17
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_02").select()
            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_02/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7097/subA02P01:SAPLBUD0:1310/btnPUSH_BUPL").press()
            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_02/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7097/subA02P01:SAPLBUD0:1310/ctxtBUT000-MARST").text = "2"
            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_02/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7097/subA02P01:SAPLBUD0:1310/ctxtBUT000-MARST").setFocus()
            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_02/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7097/subA02P01:SAPLBUD0:1310/ctxtBUT000-MARST").caretPosition = 1
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_02/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7097/subA02P01:SAPLBUD0:1310/ctxtBUT000-JOBGR").text = "Z001"
            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_02/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7097/subA02P01:SAPLBUD0:1310/ctxtBUS000FLDS-BIRTHDT").text = "01.01.1981"
            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_02/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7097/subA02P06:SAPLZ_BP_AGEAS_VN_IDNF_TAB_ENH:9001/txtGS_BUT000-ZBU_INCOME").text = "200000000"
            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_02/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7097/subA06P01:SAPLBUD0:1520/tblSAPLBUD0TCTRL_BUT0ID/ctxtGT_BUT0ID-TYPE[0,0]").setFocus()
            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_02/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7097/subA06P01:SAPLBUD0:1520/tblSAPLBUD0TCTRL_BUT0ID/ctxtGT_BUT0ID-TYPE[0,0]").caretPosition = 0
            self.session.findById("wnd[0]").sendVKey(4)
            self.session.findById("wnd[1]/usr/lbl[8,5]").setFocus()
            self.session.findById("wnd[1]/usr/lbl[8,5]").caretPosition = 15
            self.session.findById("wnd[1]").sendVKey(2)

            # Start Congnt modified: điều chỉnh để số CMND/CCCD sinh tự động các số cuối, so_cccd_Random = random.randint(1,6)
            so_cccd_Random = random.randrange(100, 10000, 1)  # start: 100; stop: 10000; 1: ở đây là bước nhảy
            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_02/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7097/subA06P01:SAPLBUD0:1520/tblSAPLBUD0TCTRL_BUT0ID/txtGT_BUT0ID-IDNUMBER[2,0]").text = "00108111" + str(
                so_cccd_Random)  # "001081928237" random sinh 05 integer cuối của dãy số cccd 12 số
            # End Congnt modified

            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_02/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7097/subA06P01:SAPLBUD0:1520/tblSAPLBUD0TCTRL_BUT0ID/txtGT_BUT0ID-IDNUMBER[2,0]").setFocus()
            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_02/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7097/subA06P01:SAPLBUD0:1520/tblSAPLBUD0TCTRL_BUT0ID/txtGT_BUT0ID-IDNUMBER[2,0]").caretPosition = 12
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_03").select()
            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_03/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7017/subA02P01:SAPLBUD0:1500/tblSAPLBUD0TCTRL_BUT0BK/ctxtGT_BUT0BK-BANKS[1,0]").text = "VN"
            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_03/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7017/subA02P01:SAPLBUD0:1500/tblSAPLBUD0TCTRL_BUT0BK/ctxtGT_BUT0BK-BANKS[1,1]").text = "VN"
            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_03/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7017/subA02P01:SAPLBUD0:1500/tblSAPLBUD0TCTRL_BUT0BK/ctxtGT_BUT0BK-BANKL[2,0]").text = "01201001"
            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_03/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7017/subA02P01:SAPLBUD0:1500/tblSAPLBUD0TCTRL_BUT0BK/ctxtGT_BUT0BK-BANKL[2,1]").text = "01201001"
            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_03/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7017/subA02P01:SAPLBUD0:1500/tblSAPLBUD0TCTRL_BUT0BK/txtGT_BUT0BK-BANKN[3,0]").text = "010101010101"
            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_03/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7017/subA02P01:SAPLBUD0:1500/tblSAPLBUD0TCTRL_BUT0BK/txtGT_BUT0BK-BANKN[3,1]").text = "000000000001"
            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_03/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7017/subA02P01:SAPLBUD0:1500/tblSAPLBUD0TCTRL_BUT0BK/txtGT_BUT0BK-BKREF[7,1]").text = "MB BANK CAT LINH"
            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_03/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7017/subA02P01:SAPLBUD0:1500/tblSAPLBUD0TCTRL_BUT0BK/txtGT_BUT0BK-BKREF[7,1]").setFocus()
            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_03/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7017/subA02P01:SAPLBUD0:1500/tblSAPLBUD0TCTRL_BUT0BK/txtGT_BUT0BK-BKREF[7,1]").caretPosition = 16
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_03/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7017/subA02P01:SAPLBUD0:1500/tblSAPLBUD0TCTRL_BUT0BK").getAbsoluteRow(
                0).selected = True
            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_03/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7017/subA02P01:SAPLBUD0:1500/tblSAPLBUD0TCTRL_BUT0BK/txtGT_BUT0BK-BKVID[0,0]").setFocus()
            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_03/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7017/subA02P01:SAPLBUD0:1500/tblSAPLBUD0TCTRL_BUT0BK/txtGT_BUT0BK-BKVID[0,0]").caretPosition = 0
            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_03/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7017/subA02P01:SAPLBUD0:1500/btnPUSH_BUP_BK_VALIDITY").press()
            self.session.findById("wnd[1]/usr/ctxtBUS000FLDS-BK_VALID_FROM").text = "01.01.2016"
            self.session.findById("wnd[1]/usr/ctxtBUS000FLDS-BK_VALID_FROM").caretPosition = 10
            self.session.findById("wnd[1]").sendVKey(0)
            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_03/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7017/subA02P01:SAPLBUD0:1500/tblSAPLBUD0TCTRL_BUT0BK").getAbsoluteRow(
                1).selected = True
            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_03/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7017/subA02P01:SAPLBUD0:1500/tblSAPLBUD0TCTRL_BUT0BK/txtGT_BUT0BK-BKVID[0,1]").setFocus()
            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_03/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7017/subA02P01:SAPLBUD0:1500/tblSAPLBUD0TCTRL_BUT0BK/txtGT_BUT0BK-BKVID[0,1]").caretPosition = 0
            self.session.findById(
                "wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2000/subSCREEN_1010_RIGHT_AREA:SAPLBUPA_DIALOG_JOEL:1000/ssubSCREEN_1000_WORKAREA_AREA:SAPLBUPA_DIALOG_JOEL:1100/ssubSCREEN_1100_MAIN_AREA:SAPLBUPA_DIALOG_JOEL:1101/tabsGS_SCREEN_1100_TABSTRIP/tabpSCREEN_1100_TAB_03/ssubSCREEN_1100_TABSTRIP_AREA:SAPLBUSS:0028/ssubGENSUB:SAPLBUSS:7017/subA02P01:SAPLBUD0:1500/btnPUSH_BUP_BK_VALIDITY").press()
            self.session.findById("wnd[1]/usr/ctxtBUS000FLDS-BK_VALID_FROM").text = "01.01.2016"
            self.session.findById("wnd[1]/usr/ctxtBUS000FLDS-BK_VALID_FROM").caretPosition = 10
            self.session.findById("wnd[1]").sendVKey(0)
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            self.session.findById("wnd[0]/tbar[0]/btn[3]").press()
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            self.session.findById("wnd[0]/shellcont[0]/shell/shellcont[1]/shell").selectItem("Key19", "Column1")
            self.session.findById("wnd[0]/shellcont[0]/shell/shellcont[1]/shell").ensureVisibleHorizontalItem("Key19", "Column1")
            self.session.findById("wnd[0]/shellcont[0]/shell/shellcont[1]/shell").doubleClickItem("Key19", "Column1")
            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPL3FCM_MAIN_DYNPROS:0200/tabsTABSTRIP/tabpCMD_TAB1/ssubTABSUB:/PM0/SAPL3FCM_MAIN_DYNPROS:0330/subSECTION1:/PM0/SAPLABX_CDM_MAIN_DYNPROS:2056/subSECTION2:SAPLZ_PM_AGEAS_ALP_CDM_APOLPR:9001/ctxt/PM0/ALEAPOLPR-INSDURINYEARS_AM").text = "10"
            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPL3FCM_MAIN_DYNPROS:0200/tabsTABSTRIP/tabpCMD_TAB1/ssubTABSUB:/PM0/SAPL3FCM_MAIN_DYNPROS:0330/subSECTION1:/PM0/SAPLABX_CDM_MAIN_DYNPROS:2056/subSECTION2:SAPLZ_PM_AGEAS_ALP_CDM_APOLPR:9001/ctxt/PM0/ALEAPOLPR-INSDURINYEARS_AM").setFocus()
            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPL3FCM_MAIN_DYNPROS:0200/tabsTABSTRIP/tabpCMD_TAB1/ssubTABSUB:/PM0/SAPL3FCM_MAIN_DYNPROS:0330/subSECTION1:/PM0/SAPLABX_CDM_MAIN_DYNPROS:2056/subSECTION2:SAPLZ_PM_AGEAS_ALP_CDM_APOLPR:9001/ctxt/PM0/ALEAPOLPR-INSDURINYEARS_AM").caretPosition = 2
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPL3FCM_MAIN_DYNPROS:0200/tabsTABSTRIP/tabpCMD_TAB2").select()
            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPL3FCM_MAIN_DYNPROS:0200/tabsTABSTRIP/tabpCMD_TAB2/ssubTABSUB:/PM0/SAPLABP_CDM_APREM:4502/subSECTION1:/PM0/SAPLABP_CDM_APREM:4516/subVIEW2:/PM0/SAPLALP_CDM_APREM:4522/cmb/PM0/ALEAPREM-PAYFRQ_CD").setFocus()
            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPL3FCM_MAIN_DYNPROS:0200/tabsTABSTRIP/tabpCMD_TAB2/ssubTABSUB:/PM0/SAPLABP_CDM_APREM:4502/subSECTION1:/PM0/SAPLABP_CDM_APREM:4516/subVIEW2:/PM0/SAPLALP_CDM_APREM:4522/cmb/PM0/ALEAPREM-PAYFRQ_CD").key = "01"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            self.session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectItem("Key21", "Column1")
            self.session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").ensureVisibleHorizontalItem("Key21", "Column1")
            self.session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").doubleClickItem("Key21", "Column1")
            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPL3FCM_MAIN_DYNPROS:0200/tabsTABSTRIP/tabpCMD_TAB2").select()
            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPL3FCM_MAIN_DYNPROS:0200/tabsTABSTRIP/tabpCMD_TAB2/ssubTABSUB:/PM0/SAPLABP_CDM_ACOV:4503/subSECTION2:SAPLZMBAL_CDM_BNF:9001/txt/PM0/ALEABNF-ENDOWMATURITY_AM").text = "300000000"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            self.session.findById("wnd[0]/tbar[1]/btn[18]").press()
            self.session.findById("wnd[0]/tbar[1]/btn[19]").press()
            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPLABP_CDM_INFO_POLJBP:2000/subSECTION2:/PM0/SAPL3FCM_MAIN_DYNPROS:2200/tabsTABSTRIP/tabpCMD_TAB1/ssubTABSUB:/PM0/SAPLABP_CDM_INFO_POLJBP:2050/subSECTION1:/PM0/SAPLABP_CDM_INFO_POLJBP:2200/txt/PM0/ABCAPOLSUMPOL-LV_POL_3").setFocus()
            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPLABP_CDM_INFO_POLJBP:2000/subSECTION2:/PM0/SAPL3FCM_MAIN_DYNPROS:2200/tabsTABSTRIP/tabpCMD_TAB1/ssubTABSUB:/PM0/SAPLABP_CDM_INFO_POLJBP:2050/subSECTION1:/PM0/SAPLABP_CDM_INFO_POLJBP:2200/txt/PM0/ABCAPOLSUMPOL-LV_POL_3").caretPosition = 8
            self.session.findById("wnd[0]/tbar[0]/btn[11]").press()
            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPLABP_CDM_INFO_POLJBP:2000/subSECTION2:/PM0/SAPL3FCM_MAIN_DYNPROS:2200/tabsTABSTRIP/tabpCMD_TAB1/ssubTABSUB:/PM0/SAPLABP_CDM_INFO_POLJBP:2050/subSECTION1:/PM0/SAPLABP_CDM_INFO_POLJBP:2200/txt/PM0/ABCAPOLSUMPOL-LV_POL_4").setFocus()
            self.session.findById(
                "wnd[0]/usr/subMAINSCREEN:/PM0/SAPLABP_CDM_INFO_POLJBP:2000/subSECTION2:/PM0/SAPL3FCM_MAIN_DYNPROS:2200/tabsTABSTRIP/tabpCMD_TAB1/ssubTABSUB:/PM0/SAPLABP_CDM_INFO_POLJBP:2050/subSECTION1:/PM0/SAPLABP_CDM_INFO_POLJBP:2200/txt/PM0/ABCAPOLSUMPOL-LV_POL_4").caretPosition = 7

        except:
            print(sys.exc_info()[0])
        messagebox.showinfo('Message', 'You have just created the Policy successfully!')