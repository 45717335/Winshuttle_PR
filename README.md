# Winshuttle_PR
Upload Purchase Requestion
## how to install
* Unzip, PR_UPLOAD.7z useing the password:"PASSWORD"
* or recreate the PR_UNLOAD.xlsm Useing the VBA code in scr,
## 更新
* 20200115 增加读取.txt 文件中 email信息的函数
```VBA
Private Function get_email_address(fln As String) As String
```
[函数连接](https://github.com/45717335/Winshuttle_PR/blob/master/src/PR_UPLOAD.xlsm/MOD_PR_Uploading.bas)
* 20200106 增加发送邮件
完成上传后发送邮件给相关人员，使用函数 
```vba
Sub SendMail(ByVal to_who As String, ByVal subject As String, ByVal body As String, ByVal attachement As String)
```
[函数连接](https://github.com/45717335/Winshuttle_PR/blob/master/src/PR_UPLOAD.xlsm/MOD_Email.bas)
