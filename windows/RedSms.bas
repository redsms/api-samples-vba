Attribute VB_Name = "RedSms"
Option Explicit

Private Const ApiBase = "https://cp.redsms.ru/api/"

Private Function Setting(ByVal name)
    Setting = Environ(CStr(name))
    If Len(Setting) = 0 Then Err.Raise 5, "RedSms", "Set " & name
End Function

Private Function BytesHex(ByVal bytes)
    Dim i, result
    result = ""
    For i = 0 To UBound(bytes)
        result = result & Right("0" & Hex(bytes(i)), 2)
    Next
    BytesHex = LCase(result)
End Function

Private Function Signature(ByVal value)
    Dim utf8 As Object, md5 As Object, bytes() As Byte
    Set utf8 = CreateObject("System.Text.UTF8Encoding")
    Set md5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    bytes = utf8.GetBytes_4(CStr(value))
    Signature = BytesHex(md5.ComputeHash_2((bytes)))
End Function

Private Function FormValue(ByVal value)
    Dim utf8 As Object, bytes() As Byte, i As Long, result As String
    Set utf8 = CreateObject("System.Text.UTF8Encoding")
    bytes = utf8.GetBytes_4(CStr(value))
    result = ""
    For i = 0 To UBound(bytes)
        result = result & "%" & Right("0" & Hex(bytes(i)), 2)
    Next
    FormValue = result
End Function

Private Function Request(ByVal method, ByVal path, ByVal values, ByVal body)
    Dim http As Object, ts As String
    ts = CStr(CLng(Date)) & "-" & CStr(Timer)
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open method, ApiBase & path, False
    http.Option(6) = False
    http.SetTimeouts 5000, 5000, 15000, 15000
    http.SetRequestHeader "login", Setting("REDSMS_LOGIN")
    http.SetRequestHeader "ts", ts
    http.SetRequestHeader "sig", Signature(values & ts & Setting("REDSMS_API_KEY"))
    If method = "POST" Then
        http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        http.Send body
    Else
        http.Send
    End If
    If http.Status <> 200 Then Err.Raise 5, "RedSms", "HTTP " & http.Status & ": " & http.ResponseText
    Request = http.ResponseText
End Function

Public Sub ClientInfo()
    Debug.Print Request("GET", "client/info", "", "")
End Sub

Public Sub SendSms(ByVal messageText)
    Dim sender, recipient, body
    If Len(messageText) = 0 Then Err.Raise 5, "RedSms", "Message is empty"
    sender = Setting("REDSMS_FROM")
    recipient = Setting("REDSMS_TO")
    Dim phonePattern As Object
    Set phonePattern = CreateObject("VBScript.RegExp")
    phonePattern.Pattern = "^[+]?[1-9][0-9]{1,14}$"
    If Not phonePattern.Test(recipient) Then Err.Raise 5, "RedSms", "Enter one international phone number"
    body = "from=" & FormValue(sender) & "&route=sms&text=" & FormValue(messageText) & "&to=" & FormValue(recipient)
    Debug.Print Request("POST", "message", sender & "sms" & messageText & recipient, body)
End Sub

Public Sub MessageStatus(ByVal uuid)
    Dim pattern, fields
    Set pattern = CreateObject("VBScript.RegExp")
    pattern.Pattern = "^[0-9A-Fa-f]{8}(-[0-9A-Fa-f]{4}){3}-[0-9A-Fa-f]{12}$"
    If Not pattern.Test(uuid) Then Err.Raise 5, "RedSms", "Invalid message UUID"
    fields = "uuid,status,errCode"
    Debug.Print Request("GET", "message/" & uuid & "?fields=" & FormValue(fields), fields, "")
End Sub
