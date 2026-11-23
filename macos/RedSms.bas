Attribute VB_Name = "RedSmsMac"
Option Explicit

Private accountLogin As String
Private accountKey As String

Public Sub ConfigureRedSms()
    accountLogin = InputBox("RED SMS login")
    accountKey = InputBox("RED SMS API key (stored for this Excel session only)")
End Sub

Public Sub ClientInfo()
    Debug.Print ApiRequest("info")
End Sub

Public Sub SendSms()
    Dim sender As String, recipient As String
    sender = InputBox("Registered sender name")
    If Len(sender) = 0 Then Exit Sub
    recipient = InputBox("Your test phone number in international format")
    If Len(recipient) = 0 Then Exit Sub
    If MsgBox("Send one paid test SMS to " & recipient & "?", vbOKCancel) <> vbOK Then Exit Sub
    Debug.Print ApiRequest("send", sender & vbLf & recipient)
End Sub

Public Sub MessageStatus()
    Dim uuid As String
    uuid = Trim$(InputBox("Message UUID"))
    If Len(uuid) = 0 Then Exit Sub
    Debug.Print ApiRequest("status", uuid)
End Sub

Private Function ApiRequest(ByVal action As String, Optional ByVal fields As String = "") As String
    If Len(accountLogin) = 0 Or Len(accountKey) = 0 Then
        ApiRequest = "Run ConfigureRedSms first."
        Exit Function
    End If
    On Error GoTo ScriptError
    ApiRequest = AppleScriptTask("RedSms.applescript", "request", _
        action & vbLf & accountLogin & vbLf & accountKey & vbLf & fields)
    Exit Function
ScriptError:
    ApiRequest = "Check RedSms.applescript in the Excel Application Scripts folder. " & Err.Description
End Function
