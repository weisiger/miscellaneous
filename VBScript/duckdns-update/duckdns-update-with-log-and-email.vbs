Option Explicit

REM https://www.duckdns.org/spec.jsp

Call DuckDNSUpdate()

Sub DuckDNSUpdate()
	On Error Resume Next
	Dim objRequest
	Dim URL
	Dim NotifyOnResponseText
	Dim NotifyEmailAddress
	Dim NotifySubject

	REM
	REM must change URL and NotifyEmailAddress below
	REM
	URL = "https://www.duckdns.org/update?domains=CHANGEME-HOSTNAME&token=CHANGEME-TOKEN&verbose=true"
	NotifyOnResponseText = "UPDATED"
	NotifyEmailAddress = "CHANGEME-EMAIL-ADDRESS"
	NotifySubject = "DUCK DNS CHANGE ALERT"

	Set objRequest = CreateObject("Microsoft.XMLHTTP")
	objRequest.open "GET", URL , false
	objRequest.Send
	If objRequest.Status = 200 Then
		Call WriteToFile("duckdns_update_log.txt",Now & " - " & objRequest.responseText & " - " & URL & vbCrLf)
		If InStr(objRequest.responseText, NotifyOnResponseText) Then
			Call WriteToFile("duckdns_update_log.txt","--> change detected, sending email..." & vbCrLf)
			Call SendEmail(NotifyEmailAddress,NotifySubject,objRequest.responseText)
		End If
	End If
	Set objRequest = Nothing
End Sub
Sub WriteToFile(sFilePath,sText)
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
	Dim objFSO
	Dim objTextStream

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objTextStream = objFSO.OpenTextFile(sFilePath, ForAppending, True, TristateFalse)

	objTextStream.write sText
	objTextStream.Close 

	set objTextStream = Nothing
	set objFSO = Nothing
end Sub
Sub SendEmail(sToAddress,sTextSubject,sTextBody)
	Dim MyEmail
	Set MyEmail = CreateObject("CDO.Message")

	REM change email from address
	MyEmail.From = "CHANGEME-EMAIL-FROM-ADDRESS"
	MyEmail.Subject = sTextSubject
	MyEmail.To = sToAddress
	MyEmail.TextBody = sTextBody

	REM value of 1 for local pickup, 2 for send using network
	MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2

	REM SMTP Server
	MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "127.0.0.1"

	REM SMTP Port
	MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 

	MyEmail.Configuration.Fields.Update
	MyEmail.Send

	set MyEmail = Nothing
	Call WriteToFile("duckdns_update_log.txt","    email sent to " & sToAddress & vbCrLf)
end Sub