Attribute VB_Name = "Excel_VBA"
Option Explicit

Function open_wb(ByRef wb As Workbook, ByVal flfp As String, Optional b_readonly As Boolean = False, Optional b_nomacro As Boolean = False, Optional psw As String = "") As Boolean
    '==========================================================
    'Open File(*.xls*):  Microsoft Excel
    '==========================================================
    open_wb = False

    Dim i As Integer
    Dim fln, flp As String
    fln = Right(flfp, Len(flfp) - InStrRev(flfp, "\"))
    flp = Left(flfp, Len(flfp) - Len(fln))
    Dim temp_b As Boolean
    temp_b = False
    For i = 1 To Workbooks.Count
        If Workbooks(i).Name = fln Then
            temp_b = True
            Set wb = Workbooks(i)
            Exit For
        End If
    Next
    If temp_b = False Then
        If Dir(flp & fln) <> "" Then

            On Error GoTo Error1:

        If b_nomacro = True Then
         Application.AutomationSecurity = msoAutomationSecurityForceDisable
        End If
        
            If b_readonly = False Then
                If psw = "" Then
                Set wb = Workbooks.Open(flp & fln)
                Else
                Set wb = Workbooks.Open(Filename:=flp & fln, WriteResPassword:=psw)
                
                End If
                
            Else
                
               If psw = "" Then
                Set wb = Workbooks.Open(flp & fln, False, True)
                Else
                  Set wb = Workbooks.Open(Filename:=flp & fln, UpdateLinks:=False, ReadOnly:=True, WriteResPassword:=psw)
                End If
                
            End If
            
        If b_nomacro = True Then
         Application.AutomationSecurity = msoAutomationSecurityLow
        End If


            temp_b = True
        End If
    End If
    open_wb = temp_b
    Exit Function
Error1:
    MsgBox "open_wb function:" + Err.Description
    Err.Clear
    Exit Function

End Function

Function ws_exist(ByRef wb As Workbook, ByVal wsn As String) As Boolean
    '==========================================================
    'Check ws Exist
    '==========================================================
    On Error GoTo Errorhand
    ws_exist = True
    Dim ws As Worksheet
    Set ws = wb.Worksheets(wsn)
    Exit Function
Errorhand:
    ws_exist = False
End Function

Function get_ws(ByRef wb As Workbook, ByVal wsname As String) As Worksheet
    On Error GoTo Errorhand
    Dim i As Integer
    Dim havewsT As Boolean
    havewsT = False
    For i = 1 To wb.Worksheets.Count
        If wb.Worksheets(i).Name = wsname Then
            Set get_ws = wb.Worksheets(i)
            havewsT = True
        End If
    Next
    If havewsT = False Then
        wb.Sheets.Add(after:=wb.Sheets(wb.Sheets.Count)).Name = wsname
        Set get_ws = wb.Worksheets(wsname)
    End If
    Exit Function
Errorhand:
    If Err.Number <> 0 Then MsgBox "get_ws function: " + Err.Description
    Err.Clear
End Function




Function open_wb2(ByRef wb As Workbook, ByVal flfp As String, Optional passw As String = "") As Boolean
'==========================================================
'���´����д� workbook
'==========================================================
open_wb2 = False

   Dim app As Object
   Set app = CreateObject("Excel.application")
    app.AutomationSecurity = msoAutomationSecurityForceDisable
    
   app.Visible = True
   
  
       
    'ActiveWindow.WindowState = xlMaximized
    
                        
   'Application.AskToUpdateLinks = False
   app.AskToUpdateLinks = False
   
Dim i As Integer
Dim fln, flp As String
fln = Right(flfp, Len(flfp) - InStrRev(flfp, "\"))
flp = Left(flfp, Len(flfp) - Len(fln))
Dim temp_b As Boolean
temp_b = False
For i = 1 To app.Workbooks.Count
If app.Workbooks(i).Name = fln Then
temp_b = True
Set wb = app.Workbooks(i)
Exit For
End If
Next
If temp_b = False Then
If Dir(flp & fln) <> "" Then

On Error GoTo Error1:
'Set wb = app.Workbooks.Open(Filename:=flp & fln, WriteResPassword:=passw)
Set wb = app.Workbooks.Open(Filename:=flp & fln, UpdateLinks:=False, WriteResPassword:=passw)

temp_b = True
End If
End If
open_wb2 = temp_b
     app.WindowState = xlMaximized
      
       
   
 app.AutomationSecurity = msoAutomationSecurityLow
         
            
Exit Function
Error1:
    MsgBox "open_wb2 function:" + Err.Description
    Err.Clear
    Exit Function
    
End Function

Function Close_wb2(ByRef wb As Workbook) As Boolean
    '==========================================================
    '���´����д� workbook
    '==========================================================
    On Error GoTo Errorhand
    Dim app As Object
    Set app = wb.Application
    If wb.Application.Workbooks.Count = 1 Then
        wb.Close
        app.Quit
        Set app = Nothing
    End If
    Exit Function
Errorhand:
    MsgBox "Close_wb2 function:" + Err.Description
    Err.Clear
End Function



Function GetColName(ByVal intCol As Long) As String
    '�к�ת����
    If InStr(CStr(Application.Version), "11") > 0 And intCol >= 1 And intCol <= 256 Then
        GetColName = Split(Workbooks(1).Worksheets(1).Cells(1, intCol).Address, "$")(1)
    ElseIf InStr(CStr(Application.Version), "12") > 0 And intCol >= 1 And intCol <= 16384 Then
        GetColName = Split(Workbooks(1).Worksheets(1).Cells(1, intCol).Address, "$")(1)

    ElseIf InStr(CStr(Application.Version), "14") > 0 And intCol >= 1 And intCol <= 16384 Then
        GetColName = Split(Workbooks(1).Worksheets(1).Cells(1, intCol).Address, "$")(1)

    Else

        GetColName = "Error"
    End If
End Function

