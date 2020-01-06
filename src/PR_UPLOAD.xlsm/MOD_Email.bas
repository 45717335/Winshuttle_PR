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
    'ʹ��Alt+S�����ʼ������Ǳ��ĵĹؼ�֮�����ⰲȫ��ʾ�Զ������ʼ�ȫ������
    Application.SendKeys "%s"
End Function


' ���͵����ʼ����ӳ���
Sub SendMail(ByVal to_who As String, ByVal subject As String, ByVal body As String, ByVal attachement As String)


    Dim objOL As Object
    Dim itmNewMail As Object
    '����Microsoft Outlook ����
    Set objOL = CreateObject("Outlook.Application")
    Set itmNewMail = objOL.CreateItem(olMailItem)
    With itmNewMail
        .subject = subject  '��ּ
        .body = body   '���ı���
       .To = to_who  '�ռ���
       If Len(attachement) > 0 Then
        .Attachments.Add attachement '����������㲻��Ҫ���͸��������԰���һ��ɾ�����ɣ�Excel�еĵ��������գ�����ɾŶ
       End If
        
       .Display  '����Outlook���ʹ���
        SetTimer 0, 0, 0, AddressOf WinProcA
   End With
   Set objOL = Nothing
   Set itmNewMail = Nothing
End Sub




'���������ʼ�
Sub BatchSendMail()


    Dim rowCount, endRowNo
   endRowNo = Cells(1, 1).CurrentRegion.Rows.Count
    '���з����ʼ�
    For rowCount = 1 To endRowNo
       SendMail Cells(rowCount, 1), Cells(rowCount, 2), Cells(rowCount, 3), Cells(rowCount, 4)
   Next
End Sub


