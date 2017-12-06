'#############################################################################################
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Name: Include Code from another file
' By: Greg Upton
' Date: 06/14/09
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Include("\\Server\Share\File") ' Path to code file

Sub Include(sInstFile)
	Dim f, s, oFSO
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	On Error Resume Next
	If oFSO.FileExists(sInstFile) Then
		Set f = oFSO.OpenTextFile(sInstFile)
		s = f.ReadAll
		f.Close
		ExecuteGlobal s
	End If
	On Error Goto 0
	Set f = Nothing
	Set oFSO = Nothing
End Sub
'####################################################################################################

Include("C:\00000directory_project\env.vbs") ' Path to code file

Dim ToAddress
Dim MessageSubject
Dim MessageBody
Dim MessageAttachment

Dim ol, ns, newMail

Set objArgs = Wscript.Arguments
if objArgs(0) <> "" then
' Display the first 2 command-line arguments
	For III = 0 to 0
		'Wscript.Echo "File Selected: " & objArgs(III)
	Next
	arg1 = objArgs(0)
	'arg2 = objArgs(1)
end if

ToAddress = arg1    'You can change this to your email address
MessageSubject = "MESSAGE FOR"
MessageBody = InputBox("Enter message ", "Enter a msg")

Set ol = WScript.CreateObject("Outlook.Application")
Set ns = ol.getNamespace("MAPI")
ns.logon mailfrom, pass, true, false
Set newMail = ol.CreateItem(olMailItem)
newMail.Subject = MessageSubject
newMail.Body = MessageBody & vbCrLf

' To Validate the recipient
Set myRecipient = ns.CreateRecipient(ToAddress)
myRecipient.Resolve
If Not myRecipient.Resolved Then
    MsgBox "Unknown Recipient"
Else
    newMail.Recipients.Add(myRecipient)
    newMail.Send
End If
Set ol = Nothing