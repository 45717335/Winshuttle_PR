Attribute VB_Name = "MOD_PR_Uploading"
Option Explicit
Private pic1 As String
Private pic2 As String
Private pic3 As String
Private pic4 As String
Private pic5 As String
Dim majjl As New C_AJJL
Dim mfso As New CFSO
Dim mokc_email As New OneKeyCls



Function init_pic() As Boolean
    init_pic = True
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Dim ws As Worksheet
    Dim str1 As String

    Set ws = get_ws(wb, "PARA")
    Dim usname As String
    usname = Environ("Computername")


    str1 = get_para_rg(ws.Range("A1:Z1"), "WINSHUTTLE_1" & usname, "N")
    If mfso.FileExists(str1) = False Then
        str1 = get_para_rg(ws.Range("A1:Z1"), "WINSHUTTLE_1" & usname, "Y")
        If mfso.FileExists(str1) = False Then
            init_pic = False
            Exit Function
        End If
    End If
    majjl.Add_Pic str1
    pic1 = str1


    str1 = get_para_rg(ws.Range("A1:Z1"), "LOGIN_1" & usname, "N")
    If mfso.FileExists(str1) = False Then
        str1 = get_para_rg(ws.Range("A1:Z1"), "LOGIN_1" & usname, "Y")
        If mfso.FileExists(str1) = False Then
            init_pic = False
            Exit Function
        End If
    End If
    majjl.Add_Pic str1
    pic2 = str1

    str1 = get_para_rg(ws.Range("A1:Z1"), "RUN_1" & usname, "N")
    If mfso.FileExists(str1) = False Then
        str1 = get_para_rg(ws.Range("A1:Z1"), "RUN_1" & usname, "Y")
        If mfso.FileExists(str1) = False Then
            init_pic = False
            Exit Function
        End If
    End If
    majjl.Add_Pic str1
    pic3 = str1

    '
    str1 = get_para_rg(ws.Range("A1:Z1"), "SAP_AUTOLOG_1" & usname, "N")
    If mfso.FileExists(str1) = False Then
        str1 = get_para_rg(ws.Range("A1:Z1"), "SAP_AUTOLOG_1" & usname, "Y")
        If mfso.FileExists(str1) = False Then
            init_pic = False
            Exit Function
        End If
    End If
    majjl.Add_Pic str1
    pic4 = str1



    str1 = get_para_rg(ws.Range("A1:Z1"), "SAP_OK_1" & usname, "N")
    If mfso.FileExists(str1) = False Then
        str1 = get_para_rg(ws.Range("A1:Z1"), "SAP_OK_1" & usname, "Y")
        If mfso.FileExists(str1) = False Then
            init_pic = False
            Exit Function
        End If
    End If
    majjl.Add_Pic str1
    pic5 = str1



End Function

Function Pr_U(flfp_pr As String) As String
    Single_V1_to_V0 flfp_pr
    majjl.delay 1000
    Pr_U = upl_spr(flfp_pr)

    majjl.delay 1000
    Single_V0_to_V1_M flfp_pr
End Function

Private Function read_pr(flfp_pr As String, mokc_pr As OneKeyCls) As Boolean
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim str1 As String, str2 As String, str3 As String, str4 As String, str5 As String, str6 As String
    Dim para1 As String, para2 As String, para3 As String, para4 As String
    Dim wsn As String
    Dim s_type As String
    Dim h_code As String
    Dim i_start As Integer
    Dim i As Integer, i_last As Integer
    Dim j As Integer



    mokc_pr.ClearAll
    read_pr = False
    If open_wb(wb, flfp_pr, True, True, "TKSY") Then

        '定义格式
        If ws_exist(wb, "PA") Then
            If wb.Worksheets("PA").Range("B10") = "Protocol:" And wb.Worksheets("PA").Range("A20") = "Validation" Then

                wsn = "PA": s_type = "V0": i_start = 21
                str1 = "WSN": str2 = wsn: mokc_pr.Add str1, str1: mokc_pr.Item(str1).Add str2, str2
                str1 = "TYPE": str2 = s_type: mokc_pr.Add str1, str1: mokc_pr.Item(str1).Add str2, str2
                str1 = "H_USER": str2 = "C3": str3 = wb.Worksheets(wsn).Range(str2): mokc_pr.Add str1, str1: mokc_pr.Item(str1).Add str2, "ADDRESS": mokc_pr.Item(str1).Add str3, "VAL"
                str1 = "H_DATE": str2 = "M3": str3 = wb.Worksheets(wsn).Range(str2): mokc_pr.Add str1, str1: mokc_pr.Item(str1).Add str2, "ADDRESS": mokc_pr.Item(str1).Add str3, "VAL"
                str1 = "H_PKA": str2 = "B7": str3 = wb.Worksheets(wsn).Range(str2): mokc_pr.Add str1, str1: mokc_pr.Item(str1).Add str2, "ADDRESS": mokc_pr.Item(str1).Add str3, "VAL"
                str1 = "H_PJNU": str2 = "G7": str3 = wb.Worksheets(wsn).Range(str2): mokc_pr.Add str1, str1: mokc_pr.Item(str1).Add str2, "ADDRESS": mokc_pr.Item(str1).Add str3, "VAL"
                str1 = "H_PJNA": str2 = "M7": str3 = wb.Worksheets(wsn).Range(str2): mokc_pr.Add str1, str1: mokc_pr.Item(str1).Add str2, "ADDRESS": mokc_pr.Item(str1).Add str3, "VAL"
                str1 = "H_PRNU": str2 = "O7": str3 = wb.Worksheets(wsn).Range(str2): mokc_pr.Add str1, str1: mokc_pr.Item(str1).Add str2, "ADDRESS": mokc_pr.Item(str1).Add str3, "VAL"
                str1 = "H_CODE": str2 = "C10": str3 = wb.Worksheets(wsn).Range(str2): mokc_pr.Add str1, str1: mokc_pr.Item(str1).Add str2, "ADDRESS": mokc_pr.Item(str1).Add str3, "VAL"
                str1 = "H_CC": str2 = "G10": str3 = wb.Worksheets(wsn).Range(str2): mokc_pr.Add str1, str1: mokc_pr.Item(str1).Add str2, "ADDRESS": mokc_pr.Item(str1).Add str3, "VAL"
                str1 = "H_AN": str2 = "G13": str3 = wb.Worksheets(wsn).Range(str2): mokc_pr.Add str1, str1: mokc_pr.Item(str1).Add str2, "ADDRESS": mokc_pr.Item(str1).Add str3, "VAL"
                mokc_pr.Add "BODY_H", "BODY_H"
                str2 = "A": str1 = "Validation": mokc_pr.Item("BODY_H").Add str1, str1: mokc_pr.Item("BODY_H").Item(str1).Add str2, str2
                str2 = "B": str1 = "SAP Item No.": mokc_pr.Item("BODY_H").Add str1, str1: mokc_pr.Item("BODY_H").Item(str1).Add str2, str2
                str2 = "C": str1 = "Item": mokc_pr.Item("BODY_H").Add str1, str1: mokc_pr.Item("BODY_H").Item(str1).Add str2, str2
                str2 = "D": str1 = "Short_text": mokc_pr.Item("BODY_H").Add str1, str1: mokc_pr.Item("BODY_H").Item(str1).Add str2, str2
                str2 = "E": str1 = "Sub_Ass": mokc_pr.Item("BODY_H").Add str1, str1: mokc_pr.Item("BODY_H").Item(str1).Add str2, str2
                str2 = "F": str1 = "Manufacturer": mokc_pr.Item("BODY_H").Add str1, str1: mokc_pr.Item("BODY_H").Item(str1).Add str2, str2
                str2 = "G": str1 = "Part_Name": mokc_pr.Item("BODY_H").Add str1, str1: mokc_pr.Item("BODY_H").Item(str1).Add str2, str2
                str2 = "H": str1 = "WBS": mokc_pr.Item("BODY_H").Add str1, str1: mokc_pr.Item("BODY_H").Item(str1).Add str2, str2
                str2 = "I": str1 = "QTY": mokc_pr.Item("BODY_H").Add str1, str1: mokc_pr.Item("BODY_H").Item(str1).Add str2, str2
                str2 = "J": str1 = "UNIT": mokc_pr.Item("BODY_H").Add str1, str1: mokc_pr.Item("BODY_H").Item(str1).Add str2, str2
                str2 = "K": str1 = "Budget": mokc_pr.Item("BODY_H").Add str1, str1: mokc_pr.Item("BODY_H").Item(str1).Add str2, str2
                str2 = "L": str1 = "Cost Element / Asset": mokc_pr.Item("BODY_H").Add str1, str1: mokc_pr.Item("BODY_H").Item(str1).Add str2, str2
                str2 = "M": str1 = "GLA": mokc_pr.Item("BODY_H").Add str1, str1: mokc_pr.Item("BODY_H").Item(str1).Add str2, str2
                str2 = "N": str1 = "Delivery Date": mokc_pr.Item("BODY_H").Add str1, str1: mokc_pr.Item("BODY_H").Item(str1).Add str2, str2
                str2 = "O": str1 = "Remark ": mokc_pr.Item("BODY_H").Add str1, str1: mokc_pr.Item("BODY_H").Item(str1).Add str2, str2
                i_last = wb.Worksheets(wsn).UsedRange.Rows(wb.Worksheets(wsn).UsedRange.Rows.Count).Row
                mokc_pr.Add "BODY", "BODY"
                For i = i_start To i_last
                    str1 = Trim(wb.Worksheets(wsn).Range(mokc_pr.Item("BODY_H").Item("QTY").Item(1).key & i))
                    str2 = Trim(wb.Worksheets(wsn).Range(mokc_pr.Item("BODY_H").Item("Item").Item(1).key & i))
                    str3 = Trim(wb.Worksheets(wsn).Range(mokc_pr.Item("BODY_H").Item("Short_text").Item(1).key & i))
                    str4 = Trim(wb.Worksheets(wsn).Range(mokc_pr.Item("BODY_H").Item("QTY").Item(1).key & i + 1))
                    str5 = Trim(wb.Worksheets(wsn).Range(mokc_pr.Item("BODY_H").Item("Item").Item(1).key & i + 1))
                    str6 = Trim(wb.Worksheets(wsn).Range(mokc_pr.Item("BODY_H").Item("Short_text").Item(1).key & i + 1))
                    If Len(str1) = 0 And Len(str2) = 0 And Len(str3) = 0 And Len(str4) = 0 And Len(str5) = 0 And Len(str6) = 0 Then
                        Exit For
                    End If
                    mokc_pr.Item("BODY").Add CStr(i), CStr(i)
                    For j = 1 To mokc_pr.Item("BODY_H").Count
                        str1 = mokc_pr.Item("BODY_H").Item(j).key
                        str2 = mokc_pr.Item("BODY_H").Item(j).Item(1).key





                        str3 = wb.Worksheets(wsn).Range(str2 & i)

                        mokc_pr.Item("BODY").Item(CStr(i)).Add str3, str1
                    Next
                Next
                read_pr = True
            End If
        End If
        '定义格式

        '赋值

        '赋值



        wb.Close 0
    Else
        mokc_pr.Item("TYPE").Add "NA", "NA"
    End If
End Function

Sub selectandupload()
    If init_pic = False Then Exit Sub
    Dim rg As Range
    Dim str1 As String
    For Each rg In Selection
        str1 = rg
        Pr_U str1
    Next
End Sub
Function save_pr_newtemplate(mokc As OneKeyCls, flfp As String) As Boolean

End Function
Function save_pr_newtemplate_VIEW(mokc As OneKeyCls, flfp As String) As Boolean

End Function
Function upl_spr(flfp_pr As String) As String

    '返回上传码
    Dim wdname As String
    Dim wb As Workbook
    'open_wb wb, flfp_pr, False, True, "TKSY"
    open_wb2 wb, flfp_pr, "TKSY"


    wdname = "Microsoft Excel - " & P_SPLIT(flfp_pr, "\", -1)



    If majjl.my_findwindow(wdname) > 0 Then


    Else



        wdname = P_SPLIT(flfp_pr, "\", -1) & " - Microsoft Excel"
        If majjl.my_findwindow(wdname) > 0 Then

        Else


        End If



    End If



    winshuttle_studio_pr wb
    Dim para1 As String, para2 As String, para3 As String, para4 As String

    majjl.delay 2000

    If majjl.L_CLICK_PIC(wdname, pic1, 10, 10) = False Then
        wb.Application.StatusBar = "Can not find pic winshuttle . PC1"
        If majjl.L_CLICK_PIC(wdname, pic1, 10, 10) = False Then
            wb.Application.StatusBar = "Can not find pic winshuttle . PC1"



        Else
            wb.Application.StatusBar = "Winshuttle Clicked!"
        End If


    Else
        wb.Application.StatusBar = "Winshuttle Clicked!"
    End If




    majjl.delay 1000


    If majjl.L_CLICK_PIC(wdname, pic2, 10, 10) = False Then
        wb.Application.StatusBar = "Can not find pic logon . PC2"
        MsgBox "please find pic1 (winshuttle) "


    Else
        wb.Application.StatusBar = "logon Clicked!"
        majjl.delay 3000
    End If




    'run
    If majjl.L_CLICK_PIC(wdname, pic3, 10, 10) = False Then
        wb.Application.StatusBar = "Run fall after 10s try again!"
        majjl.delay 10000

        '判断是否有 登陆窗口
        If majjl.my_findwindow("Log on to Connect") > 0 Then
            majjl.my_actwindow "Log on to Connect"
            majjl.L_CLICK_WIN "Log on to Connect", 256, 375
            majjl.delay 10000
            '再次点ＲＵＮ
            If majjl.L_CLICK_PIC(wdname, pic3, 10, 10) = False Then
                MsgBox "CAN NOT find Run"

            Else

            End If
            '再次点ＲＵＮ
        End If


        '判断是否有 登陆窗口


        If majjl.L_CLICK_PIC(wdname, pic3, 10, 10) = False Then
            MsgBox "Can NOt find Run,Please find manually"
        Else
            wb.Application.StatusBar = "Run Clicked!"
            majjl.delay 1000
        End If
    Else
        wb.Application.StatusBar = "Run Clicked!"
        majjl.delay 1000
    End If


    majjl.delay 1000


    'sap auto logon
    If majjl.my_findwindow("SAP Shuttle Logon") = 0 Then
        majjl.delay 3000
        wb.Application.StatusBar = "Can not find SAP shuttle Logon window, wait 3s"
        If majjl.my_findwindow("SAP Shuttle Logon") = 0 Then
            majjl.delay 3000
            wb.Application.StatusBar = "Can not find SAP shuttle Logon window, wait 3s"

        End If
    End If

    majjl.L_CLICK_PIC "SAP Shuttle Logon", pic4, 10, 10






    majjl.delay 1000
    ' sap on
    If majjl.L_CLICK_PIC("SAP Shuttle Logon", pic5, 10, 10) = False Then
        '点击 logon 之后没有 点到run 可能是因为出现 了登陆框，检查登陆框 点击之后，再次 点击 run

        If majjl.my_findwindow("Log on to Connect") > 0 Then
            majjl.L_CLICK_WIN "Log on to Connect", 252, 375
            majjl.delay 10000
            If majjl.L_CLICK_PIC("SAP Shuttle Logon", pic5, 10, 10) = False Then
                MsgBox "Can Not Run Winshuttle,Please Run Manually"
            End If
        End If

    End If






    Do While wb.Worksheets(1).Range("C10") = ""
        majjl.delay 3000

    Loop
    upl_spr = wb.Worksheets(1).Range("C10")


    majjl.delay 20000


    wb.Save
    wb.Application.DisplayAlerts = False
    wb.SaveAs Filename:=wb.Fullname, WriteResPassword:="TKSY"

'sendmail
para4 = wb.Worksheets(1).Range("C3")
para3 = wb.Worksheets(1).Range("C10")
If mokc_email.Item(para4) Is Nothing Then
mokc_email.Add para4, para4
mokc_email.Item(para4).Add para3 & wb.Fullname, para3 & wb.Fullname
Else
mokc_email.Item(para4).Add para3 & wb.Fullname, para3 & wb.Fullname
End If
'sendmail


    Close_wb2 wb
    If wb Is Nothing Then
    Else
        'MsgBox "have not close"
        '按键精灵点ok关闭
        '按键精灵点ok关闭
    End If

    'wb.Close 0
    '返回上传码

End Function

Function save_pr(mokc_pr As OneKeyCls, flfp_pr As String, Optional s_template As String = "VIEW")
    's_template 有多种模板 "VIEW","UPLOAD",...
    Dim para1 As String, para2 As String, para3 As String, para4 As String, para5 As String
    Dim str1 As String, str2 As String, str3 As String, str4 As String, str5 As String
    Dim wb As Workbook
    If s_template = "UPLOAD" Then
        para1 = ""

    End If



End Function


Private Function Single_V1_to_V0(flfp_pr As String) As String
    '检查目标文件是否为新版本


    Dim c10 As String

    Dim b_continue As Boolean

    Dim b_upload As Boolean
    Dim address_email As String

    Dim usern As String
    Dim pjn As String

    b_upload = False

    Dim j As Integer
    Dim j_last As Integer
    'Dim mokc As New OneKeyCls
    Dim wb As Workbook
    Dim windowname As String
    Dim mfso As New CFSO
    Dim fln As String
    Dim fdn As String
    Dim flfp As String
    Dim i_last As Integer
    Dim i As Integer
    Dim wb_list As Workbook
    fln = Right(flfp_pr, Len(flfp_pr) - InStrRev(flfp_pr, "\"))
    fdn = Left(flfp_pr, InStrRev(flfp_pr, "\"))

    If (fln Like "P?????_CN*.xlsm" Or fln Like "M?????_CN*.xlsm" Or fln Like "D?????_CN*.xlsm") And mfso.FileExists(fdn & fln) Then
        open_wb2 wb, flfp_pr, "TKSY"


        c10 = Get_RangeVal(wb.Worksheets(1), "C10")

        '====================如果 已经是旧的格式，不进行转换
        '====================如果 不是新格式，推出也不进行转换，ERROR
        If ws_exist(wb, "PA.") = False Then
            'wb.Close
            '20190108
            ' wb.Close 0
            Close_wb2 wb

            Exit Function
        End If
        If wb.Worksheets("PA.").Range("C18") <> "Total Item Numbers" Then
            'wb.Close
            '20190108
            Close_wb2 wb
            'wb.Close 0
            Exit Function
        End If
        '================================================================================================
        Dim s_date As String
        Dim fdn_bak_bef As String
        Dim fdn_bak_aft As String
        Dim flfp2 As String

        fdn_bak_bef = "Z:\24_Temp\PA_Logs\V1.2\V1.2_TO_V1.1\BEF\"
        fdn_bak_aft = "Z:\24_Temp\PA_Logs\V1.2\V1.2_TO_V1.1\AFT\"



        s_date = Format(now(), "YYYYMMDD")
        If mfso.folderexists(fdn_bak_bef & s_date & "\") = False Then
            mfso.CreateFolder fdn_bak_bef & s_date & "\"
        End If
        If mfso.folderexists(fdn_bak_aft & s_date & "\") = False Then
            mfso.CreateFolder fdn_bak_aft & s_date & "\"
        End If

        If mfso.FileExists(fdn_bak_bef & s_date & "\" & fln) Then
            'Kill fdn_bak_bef & s_date & "\" & fln
            mfso.deletefile fdn_bak_bef & s_date & "\" & fln
        End If
        mfso.copy_file flfp_pr, fdn_bak_bef & s_date & "\" & fln



        '=================================模板，Z:\24_Temp\PA_Logs\V1.2\TEMPLATE\010c1612_Purchase Requisition(20170503).xlsm
        '=================================本地模板，D:\VBA\EXCEL_MODULE\PR\V1.2\TEMPLATE\010c1612_Purchase Requisition(20170503).xlsm
        '取最新的拷贝到本地然后打开，
        Dim fdn_NewT_Net As String
        Dim fdn_NewT_LOC As String
        Dim fln_NewT As String

        fdn_NewT_Net = "Z:\24_Temp\PA_Logs\V1.1\TEMPLATE\"
        fdn_NewT_LOC = "D:\VBA\EXCEL_MODULE\PR\V1.1\TEMPLATE\"
        fln_NewT = "010c1612_Purchase Requisition.xlsm"
        If mfso.folderexists(fdn_NewT_LOC) = False Then mfso.CreateFolder fdn_NewT_LOC
        If mfso.folderexists(fdn_NewT_Net) = False Then mfso.CreateFolder fdn_NewT_Net
        If mfso.FileExists(fdn_NewT_LOC & fln_NewT) = False Then
            If mfso.FileExists(fdn_NewT_Net & fln_NewT) = False Then
                Single_V1_to_V0 = "No Template Exists! " & Chr(10) & fdn_NewT_Net & fln_NewT
                Close_wb2 wb
                Exit Function
            Else
                mfso.copy_file fdn_NewT_Net & fln_NewT, fdn_NewT_LOC & fln_NewT
            End If
        Else
            If mfso.FileExists(fdn_NewT_Net & fln_NewT) Then
                If mfso.Datelastmodify(fdn_NewT_Net & fln_NewT) > mfso.Datelastmodify(fdn_NewT_LOC & fln_NewT) Then
                    Kill fdn_NewT_LOC & fln_NewT
                    mfso.copy_file fdn_NewT_Net & fln_NewT, fdn_NewT_LOC & fln_NewT
                End If
            End If
        End If

        '===========================================================模板预先删除全部宏，并另存
        Dim wb_new As Workbook
        Application.DisplayAlerts = False

        Application.AskToUpdateLinks = False

        If open_wb(wb_new, fdn_NewT_LOC & fln_NewT) Then
        Else
            Single_V1_to_V0 = "Can not open template!"
            MsgBox Single_V1_to_V0
            Exit Function
        End If
        wb_new.SaveAs fdn_bak_aft & s_date & "\" & fln

        '=======================开始转换格式






        Dim wsf As Worksheet
        Dim wst As Worksheet
        Set wsf = wb.Worksheets("PA.")
        Set wst = wb_new.Worksheets("PA")

        wst.Range("C3") = wsf.Range("E3")
        wst.Range("M3") = wsf.Range("N3")
        wst.Range("G13") = wsf.Range("J7")
        wst.Range("G7") = wsf.Range("E7")
        wst.Range("O7") = wsf.Range("P7")
        wst.Range("C10") = wsf.Range("D10")
        wst.Range("G10") = wsf.Range("G7")
        wst.Range("B7") = wsf.Range("C7")
        wst.Range("M7") = wsf.Range("M7")



        '取最新的拷贝到本地然后打开，
        Dim i_count As Integer
        Dim str_total As String

        'i_last = wsf.UsedRange.Rows(wsf.UsedRange.Rows.Count).row
        i_last = get_rowscount(wsf)


        Dim stra As String
        Dim strb As String
        Dim strc As String
        i_count = 0

        Dim dbl1 As Double
        Dim dbl2 As Double
        Dim dbl3 As Double


        For i = 20 To i_last
            stra = Trim(wsf.Range("B" & i))
            strb = Trim(wsf.Range("J" & i))
            strc = Trim(wsf.Range("L" & i))

            If Len(stra & strb & strc) > 0 Then
                wst.Range("B" & i + 1) = wsf.Range("B" & i)
                wst.Range("C" & i + 1) = wsf.Range("C" & i)
                wst.Range("D" & i + 1) = wsf.Range("E" & i)
                wst.Range("E" & i + 1) = wsf.Range("F" & i)
                wst.Range("F" & i + 1) = wsf.Range("G" & i)
                wst.Range("G" & i + 1) = wsf.Range("H" & i)


                wst.Range("H" & i + 1) = wsf.Range("I" & i)
                wst.Range("I" & i + 1) = wsf.Range("J" & i)
                wst.Range("J" & i + 1) = wsf.Range("K" & i)
                wst.Range("K" & i + 1) = wsf.Range("L" & i)
                wst.Range("L" & i + 1) = wsf.Range("M" & i)
                wst.Range("M" & i + 1) = wsf.Range("Q" & i)

                wst.Range("N" & i + 1) = wsf.Range("O" & i)
                wst.Range("O" & i + 1) = wsf.Range("P" & i)

                'O,颜色，备注一并复制
                If Not wsf.Range("P" & i).comment Is Nothing Then
                    wst.Range("O" & i - 1).AddComment wsf.Range("P" & i).comment.Text
                End If
                wst.Range("O" & i - 1).Interior.Color = wsf.Range("P" & i).Interior.Color
                'O,颜色，备注一并复制
                i_count = i_count + 1

                my_CDBL strb, dbl1
                my_CDBL strc, dbl2
                my_CDBL str_total, dbl3

                str_total = CStr(dbl3 + dbl1 * dbl2)

                'wst.Range("L" & i + 1) = dbl2


            End If





        Next


        'wst.Range("E18") = i_count
        'wst.Range("L18") = "CNY"
        wst.Range("N19") = str_total

        '设置打印区域
        wst.PageSetup.PrintArea = "$C$1:$P$" & i_count + 21


        '设置打印区域


    End If


    flfp = wb.Fullname

    flfp2 = wb_new.Fullname


    wb_new.SaveAs Filename:=wb_new.Fullname, WriteResPassword:="TKSY"

    '插入代码

    '插入代码


    wb_new.Close

    Close_wb2 wb


    '=============================================
    '                     'Insert Macro PR_CHECK
    '                     If open_wb(wb_new, "Z:\51_Engineering\06_Service\EXCEL_MODULE\PR\MACRO_INSERT.xlsm") Then
    '                     wb_new.Worksheets(1).Range("A1") = flfp2
    '                     wb_new.Worksheets(1).Range("A2") = "TKSY"
    '                     wb_new.Worksheets(1).Range("A3") = "ThisWorkbook"
    '                     wb_new.Worksheets(1).Range("A4") = "PR_MANUAL_CHECK"
    '                     Application.Run "MACRO_INSERT.xlsm!MACROINSERT"
    '                     wb_new.Saved = True
    '                     wb_new.Close
    '                     End If
    '                     'Insert Macro PR_CHECK
    '=============================================





    '删掉原始文件，并复制新作的文件
    'Kill flfp
    mfso.deletefile flfp
    mfso.copy_file flfp2, flfp

End Function

Private Function Single_V0_to_V1_M(flfp_pr As String) As String
    '直接是带宏 带 密码的
    Dim str_temp As String
    str_temp = ""


    Dim c10 As String

    Dim b_continue As Boolean

    Dim b_upload As Boolean
    Dim address_email As String

    Dim usern As String
    Dim pjn As String

    b_upload = False

    Dim j As Integer
    Dim j_last As Integer
    'Dim mokc As New OneKeyCls
    Dim wb As Workbook
    Dim windowname As String
    Dim mfso As New CFSO
    Dim fln As String
    Dim fdn As String
    Dim flfp As String
    Dim i_last As Integer
    Dim i As Integer
    Dim wb_list As Workbook
    fln = Right(flfp_pr, Len(flfp_pr) - InStrRev(flfp_pr, "\"))
    fdn = Left(flfp_pr, InStrRev(flfp_pr, "\"))

    If (fln Like "P?????_CN*.xlsm" And mfso.FileExists(fdn & fln)) Or (fln Like "M?????_CN*.xlsm" And mfso.FileExists(fdn & fln)) Or (fln Like "D?????_CN*.xlsm" And mfso.FileExists(fdn & fln)) Then
        open_wb2 wb, flfp_pr, "TKSY"
        c10 = Get_RangeVal(wb.Worksheets("PA"), "C10")
        '====================如果没有设置写入密码保护TKSY则直接推出，不进行转换
        '====================如果是只读，直接退出，不进行转换。
        '====================如果C10单元格 格式不是“Purchase requisition number * created”直接退出,不进行转换
        If wb.ReadOnly Then
            Single_V0_to_V1_M = "Can not be readonly!"
            'MsgBox Single_V0_to_V1_M
            Close_wb2 wb
            Exit Function
        End If

        If wb.WriteReserved = False Then
            Single_V0_to_V1_M = "Must have the writepassword of 'TKSY'."
            'MsgBox Single_V0_to_V1_M
            Close_wb2 wb
            Exit Function

        End If


        If Not (c10 Like "Purchase requisition number * created") Then
            Single_V0_to_V1_M = "Must have PR number!"
            'MsgBox Single_V0_to_V1_M
            Close_wb2 wb
            Exit Function
        End If
        '================================================================================================
        Dim s_date As String
        Dim fdn_bak_bef As String
        Dim fdn_bak_aft As String
        Dim flfp2 As String

        fdn_bak_bef = "Z:\24_Temp\PA_Logs\V1.2\PR_UPLOADED\BEF\"
        fdn_bak_aft = "Z:\24_Temp\PA_Logs\V1.2\PR_UPLOADED\AFT\"

        s_date = Format(now(), "YYYYMMDD")
        If mfso.folderexists(fdn_bak_bef & s_date & "\") = False Then
            mfso.CreateFolder fdn_bak_bef & s_date & "\"
        End If
        If mfso.folderexists(fdn_bak_aft & s_date & "\") = False Then
            mfso.CreateFolder fdn_bak_aft & s_date & "\"
        End If

        If mfso.FileExists(fdn_bak_bef & s_date & "\" & fln) Then
            Kill fdn_bak_bef & s_date & "\" & fln
        End If
        mfso.copy_file flfp_pr, fdn_bak_bef & s_date & "\" & fln



        '=================================模板，Z:\24_Temp\PA_Logs\V1.2\TEMPLATE\010c1612_Purchase Requisition(20170503).xlsm
        '=================================本地模板，D:\VBA\EXCEL_MODULE\PR\V1.2\TEMPLATE\010c1612_Purchase Requisition(20170503).xlsm
        '取最新的拷贝到本地然后打开，
        Dim fdn_NewT_Net As String
        Dim fdn_NewT_LOC As String
        Dim fln_NewT As String

        fdn_NewT_Net = "Z:\24_Temp\PA_Logs\V1.2\TEMPLATE\"
        fdn_NewT_LOC = "D:\VBA\EXCEL_MODULE\PR\V1.2\TEMPLATE\"
        'fln_NewT = "010c1612_Purchase Requisition(20170503).xlsm"
        fln_NewT = "010c1612_Purchase Requisition(20170503)_M.xlsm"
        If mfso.folderexists(fdn_NewT_LOC) = False Then mfso.CreateFolder fdn_NewT_LOC
        If mfso.folderexists(fdn_NewT_Net) = False Then mfso.CreateFolder fdn_NewT_Net
        If mfso.FileExists(fdn_NewT_LOC & fln_NewT) = False Then
            If mfso.FileExists(fdn_NewT_Net & fln_NewT) = False Then
                Single_V0_to_V1_M = "No Template Exists! " & Chr(10) & fdn_NewT_Net & fln_NewT
                Close_wb2 wb
                Exit Function
            Else
                mfso.copy_file fdn_NewT_Net & fln_NewT, fdn_NewT_LOC & fln_NewT
            End If
        Else
            If mfso.FileExists(fdn_NewT_Net & fln_NewT) Then
                If mfso.Datelastmodify(fdn_NewT_Net & fln_NewT) > mfso.Datelastmodify(fdn_NewT_LOC & fln_NewT) Then
                    Kill fdn_NewT_LOC & fln_NewT
                    mfso.copy_file fdn_NewT_Net & fln_NewT, fdn_NewT_LOC & fln_NewT
                End If
            End If
        End If

        '===========================================================模板预先删除全部宏，并另存
        Dim wb_new As Workbook
        Application.DisplayAlerts = False

        Application.AskToUpdateLinks = False

        Application.AutomationSecurity = msoAutomationSecurityForceDisable


        If open_wb(wb_new, fdn_NewT_LOC & fln_NewT) Then
        Else
            Single_V0_to_V1_M = "Can not open template!"
            MsgBox Single_V0_to_V1_M
            Exit Function
        End If
        wb_new.SaveAs fdn_bak_aft & s_date & "\" & fln

        '=======================开始转换格式






        Dim wsf As Worksheet
        Dim wst As Worksheet
        Set wsf = wb.Worksheets("PA")
        Set wst = wb_new.Worksheets("PA.")

        wst.Range("E3") = wsf.Range("C3")
        wst.Range("N3") = wsf.Range("M3")
        wst.Range("J7") = wsf.Range("G13")
        wst.Range("E7") = wsf.Range("G7")
        wst.Range("P7") = wsf.Range("O7")
        wst.Range("D10") = wsf.Range("C10")
        wst.Range("G7") = wsf.Range("G10")
        wst.Range("C7") = wsf.Range("B7")
        wst.Range("M7") = wsf.Range("M7")



        '取最新的拷贝到本地然后打开，
        Dim i_count As Integer
        Dim str_total As String

        i_last = wsf.UsedRange.Rows(wsf.UsedRange.Rows.Count).Row
        Dim stra As String
        Dim strb As String
        Dim strc As String
        i_count = 0

        Dim dbl1 As Double
        Dim dbl2 As Double
        Dim dbl3 As Double


        For i = 21 To i_last


            stra = Trim(wsf.Range("B" & i))
            strb = Trim(wsf.Range("I" & i))
            strc = Trim(wsf.Range("K" & i))

            If Len(stra & strb & strc) > 0 Then
                wst.Range("B" & i - 1) = wsf.Range("B" & i)
                wst.Range("C" & i - 1) = wsf.Range("C" & i)
                wst.Range("E" & i - 1) = wsf.Range("D" & i)
                wst.Range("F" & i - 1) = wsf.Range("E" & i)
                wst.Range("G" & i - 1) = wsf.Range("F" & i)
                wst.Range("H" & i - 1) = wsf.Range("G" & i)


                wst.Range("I" & i - 1) = wsf.Range("H" & i)
                wst.Range("J" & i - 1) = wsf.Range("I" & i)
                wst.Range("K" & i - 1) = wsf.Range("J" & i)
                wst.Range("L" & i - 1) = wsf.Range("K" & i)
                wst.Range("M" & i - 1) = wsf.Range("L" & i)
                wst.Range("Q" & i - 1) = wsf.Range("M" & i)

                wst.Range("O" & i - 1) = wsf.Range("N" & i)
                wst.Range("P" & i - 1) = wsf.Range("O" & i)

                'O,颜色，备注一并复制
                If Not wsf.Range("O" & i).comment Is Nothing Then
                    wst.Range("P" & i - 1).AddComment wsf.Range("O" & i).comment.Text
                End If
                wst.Range("P" & i - 1).Interior.Color = wsf.Range("O" & i).Interior.Color
                'O,颜色，备注一并复制
                i_count = i_count + 1

                my_CDBL strb, dbl1
                my_CDBL strc, dbl2
                my_CDBL str_total, dbl3

                str_total = CStr(dbl3 + dbl1 * dbl2)


                wst.Range("L" & i - 1) = dbl2

            End If





        Next

        '================================allow only give total price xuefneg.gao@thyssenkrupp.com 20170528
        Dim str_oldtotal As String
        Dim dbl_old As Double
        str_oldtotal = wsf.Range("N19").Text
        my_CDBL str_oldtotal, dbl_old
        If dbl_old > 0 Then
            str_total = CStr(dbl_old)
        End If
        '=================================allow only give total price xuefneg.gao@thyssenkrupp.com 20170528

        wst.Range("E18") = i_count
        wst.Range("L18") = "CNY"
        wst.Range("N18") = str_total

        '设置打印区域
        wst.PageSetup.PrintArea = "$C$1:$P$" & i_count + 20


        '设置打印区域


    End If


    flfp = wb.Fullname

    flfp2 = wb_new.Fullname


    wb_new.SaveAs Filename:=wb_new.Fullname, WriteResPassword:="TKSY"

    '插入代码

    '插入代码


    wb_new.Close

    Close_wb2 wb








    '删掉原始文件，并复制新作的文件
    'Kill flfp
    mfso.deletefile flfp

    mfso.copy_file flfp2, flfp

    Application.AutomationSecurity = msoAutomationSecurityLow

End Function
Function Get_RangeVal(ws As Worksheet, rgn As String) As String
    Application.DisplayAlerts = False

    On Error GoTo Error1:
    Get_RangeVal = ws.Range(rgn)
    Exit Function
Error1:
    Get_RangeVal = ""
End Function
Function get_rowscount(wsf As Worksheet) As Integer
    Dim i As Integer
    If wsf.UsedRange.Rows(wsf.UsedRange.Rows.Count).Row > 32700 Then
        get_rowscount = 20
        For i = 21 To 10000
            If Len(Trim(wsf.Range("B" & i)) & Trim(wsf.Range("C" & i))) > 0 Then
                get_rowscount = i
            End If
            If Len(Trim(wsf.Range("B" & i + 1)) & Trim(wsf.Range("C" & i + 1))) = 0 And Len(Trim(wsf.Range("B" & i + 2)) & Trim(wsf.Range("C" & i + 2))) = 0 Then
                Exit For
            End If
        Next
    Else
        get_rowscount = wsf.UsedRange.Rows(wsf.UsedRange.Rows.Count).Row
    End If
End Function


'带容错的，文字转换为数值
Function my_CDBL(s_in As String, ByRef ff As Double) As Boolean
    my_CDBL = True
    On Error GoTo Errorhand
    ff = 0
    ff = CDbl(s_in)
    Exit Function
Errorhand:
    my_CDBL = False
End Function

Function winshuttle_studio_pr(wb As Workbook) As Boolean
    Dim wb_template As Workbook
    Set wb_template = wb.Application.Workbooks.Open("Z:\24_Temp\PA_Logs\V1.3\Draft_PurchaseApplication_Studio_V100.xlsm")
    If ws_exist(wb, "WinshuttleStudio") = False Then wb_template.Worksheets("WinshuttleStudio").Copy wb.Worksheets(wb.Worksheets.Count)
    If ws_exist(wb, "PA") = True Then If wb.ActiveSheet.Name <> "PA" Then wb.Worksheets("PA").Activate
    '只留两个表格，"PA" 和 "WinshuttleStudio"
    Dim ws As Worksheet
    wb.Application.DisplayAlerts = False

    For Each ws In wb.Worksheets
        If ws.Name = "PA" Then
            ws.Columns("Q:Q").Delete Shift:=xlToLeft
        End If
        If ws.Name <> "PA" And ws.Name <> "WinshuttleStudio" Then
            ws.Delete
        End If
    Next

    wb.Application.DisplayAlerts = True

    wb_template.Close 0
End Function

Sub test()
    'If majjl.my_findwindow("WinshuttleStudioAddin") > 0 Then
    '    majjl.my_actwindow "WinshuttleStudioAddin"''
    ''
    '        majjl.L_CLICK_WIN "WinshuttleStudioAddin", 425, 96
    ''        majjl.delay 10000
    '       '再次点ＲＵＮ'

    ' End If
    
    'sendmail
    Dim para4 As String
    Dim para3 As String
    
para4 = "GXF"
para3 = "PRnumb1"

If mokc_email.Item(para4) Is Nothing Then
mokc_email.Add para4, para4
mokc_email.Item(para4).Add para3 & " wb.Fullname", para3 & " wb.Fullname"
Else
mokc_email.Item(para4).Add para3 & " wb.Fullname", para3 & " wb.Fullname"
End If

'sendmail

    send_email
    


End Sub


Sub PRU()
    'ONLY ONCE!

    Dim mfso As New CFSO

    If init_pic = False Then Exit Sub

    Dim usname As String
    usname = Environ("Computername")
    Dim s_prto As String

    Dim myCon      As New ADODB.Connection
    Dim myRst      As New ADODB.Recordset
    Dim myFileName As String
    Dim myTblName  As String
    Dim myKey      As String
    Dim mySht      As Worksheet
    Dim i          As Long
    Dim j          As Long
    Dim str1 As String
    Dim str4    As String

    Dim L_key As Long



    Dim ExcelApp As Excel.Application
    Dim ExcelWB As Excel.Workbook


    Set ExcelApp = GetObject(, "Excel.application")


    Set ExcelWB = Nothing
    myFileName = "HTML_Data.mdb"
    myTblName = "PR_TOBEUPLOAD"

    For i = 1 To 2

        majjl.delay 1000

        DoEvents

        myCon.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "Z:\24_Temp\PA_Logs\HTML\mdb\" & myFileName & ";"
        myCon.Execute "SELECT * FROM " & myTblName


        With myRst
            .Index = "PrimaryKey"
            myRst.Open Source:=myTblName, ActiveConnection:=myCon, _
                                    CursorType:=adOpenKeyset, LockType:=adLockOptimistic, _
                                    Options:=adCmdTableDirect




            Do While Not .EOF


                str1 = .Fields(3).Value


                If mfso.FileExists(str1) And .Fields(1).Value = "TOBEDONE" Then

                    If Read_Only(str1) = False Then


                        .Fields(1).Value = "DOING_" & usname
                        .Update
                        L_key = .Fields(0).Value


                        .Close
                        myCon.Close






                        s_prto = Pr_U(str1)






                        '重新打开连接


                        myCon.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=" & "Z:\24_Temp\PA_Logs\HTML\mdb\" & myFileName & ";"
                        myCon.Execute "SELECT * FROM " & myTblName



                        .Index = "PrimaryKey"
                        .Open Source:=myTblName, ActiveConnection:=myCon, _
                                    CursorType:=adOpenKeyset, LockType:=adLockOptimistic, _
                                    Options:=adCmdTableDirect


                        myRst.Seek L_key




                        '重新打开连接


                        .Fields(4).Value = s_prto
                        .Fields(1).Value = "DONE"

                    Else
                        .Fields(1).Value = "READ ONLY"
                        .Update


                    End If




                ElseIf mfso.FileExists(str1) = False Then
                    .Fields(1).Value = "NOT_EXIST"
                ElseIf .Fields(1).Value = "DOING_" & usname Then
                    '因为只有一台电脑能够上传，所以意外中断之后，还是需要重新上传的
                    .Fields(1).Value = "TOBEDONE"

                End If







                .MoveNext

            Loop


            .Close
        End With
        myCon.Close










        Set myRst = Nothing
        Set myCon = Nothing
        ' If Not ExcelApp Is Nothing Then
        ' For j = 1 To ExcelApp.Workbooks.Count
        ' If ExcelApp.Workbooks(j).Name <> "PR_UPLOADING_20190524.xlsm" And ExcelApp.Workbooks(j).Name <> "AnJianJingLing_WinShuttle_20190524.xlsm" Then
        '  ExcelApp.Workbooks(j).Saved = True
        ' ExcelApp.Workbooks(j).Close
        ' End If

        ' Next
        ' End If




    Next



End Sub



Private Function Read_Only(str1 As String) As Boolean
    '本函数用于判断一个电子表格是否为只读，是只读则返回Ｔｒｕｅ并且在　Ｄ：＼ＥＲＲＯＲ＼Ｅｒｒｏｒ．ｔｘｔ　记录，并打开文件夹
    Dim f As String
    f = "D:\ERROR\error.txt"
    Dim mfso As New CFSO
    Dim wb As Workbook
    Workbooks.Application.ScreenUpdating = False
    'Set wb = Workbooks.Open(Filename:=str1, WriteResPassword:="TKSY")
    Set wb = Workbooks.Open(Filename:=str1, WriteResPassword:="TKSY", UpdateLinks:=False)
    If wb.ReadOnly = True Then
        Read_Only = True
        If mfso.FileExists(f) = False Then
            Open f For Output As #1
            Print #1, str1 & " READ ONLY " & now()
            Close #1
        Else
            Open f For Append As #1
            Print #1, str1 & " READ ONLY " & now()
            Close #1
        End If
        Shell "explorer.exe D:\ERROR\", vbNormalFocus
    Else
        Read_Only = False
    End If
    'wb.Close
    wb.Close 0
    Workbooks.Application.ScreenUpdating = True
End Function


Sub pru_and_prc()


    Dim wshShell As Object
    Set wshShell = CreateObject("WScript.Shell")
    Dim i As Integer
    For i = 1 To 2000
        PRU
        Application.StatusBar = "restart at:" & Format(now() + CDate("00:05:00"), "YYYY-MM-DD HH:MM:SS")
        send_email
        majjl.delay 300000
        wshShell.SendKeys "{NUMLOCK}"
        majjl.delay 500
        wshShell.SendKeys "{NUMLOCK}"
    Next

End Sub
Private Sub send_email()
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Dim ws As Worksheet
    Set ws = wb.ActiveSheet
    Dim para1 As String, para2 As String, para3 As String, para4 As String
    Dim str1 As String, str2 As String, str3 As String, str4 As String
    Dim j As Integer
    Dim i As Integer
    If mokc_email.Count > 0 Then
        For i = 1 To mokc_email.Count
            str1 = mokc_email.Item(i).key
            para1 = get_para_rg(ws.Range("A2:Z2"), str1, "N")
            If para1 = "" Then
            para1 = get_para_rg(ws.Range("A2:Z2"), str1, "Y")
            End If
            para2 = ""
            If para1 Like "*@thyssenkrupp.com" Then
                For j = 1 To mokc_email.Item(i).Count
                    para2 = para2 & mokc_email.Item(i).Item(j).key & Chr(10)
                Next
            End If
            SendMail para1, "PR UP LOADING FINISH", para2, ""
        Next
        mokc_email.ClearAll
    End If
End Sub

