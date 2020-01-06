Attribute VB_Name = "MOD_Email"
Public Declare Function SetTimer Lib "user32" _
        (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerfunc As Long) As Long
Public Declare Function KillTimer Lib "user32" _
        (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



Function WinProcA(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal SysTime As Long) As Long

    KillTimer 0, idEvent
    DoEvents
    Sleep 100
    '使用Alt+S发送邮件，这是本文的关键之处，免安全提示自动发送邮件全靠它了
    Application.SendKeys "%s"
End Function


' 发送单个邮件的子程序
Sub SendMail(ByVal to_who As String, ByVal subject As String, ByVal body As String, ByVal attachement As String)


    Dim objOL As Object
    Dim itmNewMail As Object
    '引用Microsoft Outlook 对象
    Set objOL = CreateObject("Outlook.Application")
    Set itmNewMail = objOL.CreateItem(olMailItem)
    With itmNewMail
        .subject = subject  '主旨
        .body = body   '正文本文
       .To = to_who  '收件者
       If Len(attachement) > 0 Then
        .Attachments.Add attachement '附件，如果你不需要发送附件，可以把这一句删掉即可，Excel中的第四列留空，不能删哦
       End If
        
       .Display  '启动Outlook发送窗口
        SetTimer 0, 0, 0, AddressOf WinProcA
   End With
   Set objOL = Nothing
   Set itmNewMail = Nothing
End Sub




'批量发送邮件
Sub BatchSendMail()


    Dim rowCount, endRowNo
   endRowNo = Cells(1, 1).CurrentRegion.Rows.Count
    '逐行发送邮件
    For rowCount = 1 To endRowNo
       SendMail Cells(rowCount, 1), Cells(rowCount, 2), Cells(rowCount, 3), Cells(rowCount, 4)
   Next
End Sub


