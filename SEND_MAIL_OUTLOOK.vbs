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
'' SIG '' Begin signature block
'' SIG '' MIID8wYJKoZIhvcNAQcCoIID5DCCA+ACAQExCzAJBgUr
'' SIG '' DgMCGgUAMGcGCisGAQQBgjcCAQSgWTBXMDIGCisGAQQB
'' SIG '' gjcCAR4wJAIBAQQQTvApFpkntU2P5azhDxfrqwIBAAIB
'' SIG '' AAIBAAIBAAIBADAhMAkGBSsOAwIaBQAEFLHADx4VKfqd
'' SIG '' /wI/OKeWNbcuB72aoIICGzCCAhcwggGAoAMCAQICEIig
'' SIG '' ohjssziCQLgC4RnUrh0wDQYJKoZIhvcNAQEFBQAwGDEW
'' SIG '' MBQGA1UEAxMNUm9uYWxkIEJsYW5jbzAeFw0xNzAxMDEw
'' SIG '' NTAwMDBaFw0yMzAxMDEwNTAwMDBaMBgxFjAUBgNVBAMT
'' SIG '' DVJvbmFsZCBCbGFuY28wgZ8wDQYJKoZIhvcNAQEBBQAD
'' SIG '' gY0AMIGJAoGBANuG7UVFH9ARXhRqrHl+Z8XGqD7Are2u
'' SIG '' x2ksTzsQtVsfF4c/smclBWvYy/tgRx3N41mYLJtlYmGt
'' SIG '' LuqSa3rgJNf2MCsGzCm2+L/pXet1DnJ8sKTiED/9ZXtF
'' SIG '' pdSd/i3jyxVTlLj9hNaEvTqx6QnQQ9qe0ww4EDLAWbu3
'' SIG '' tqR/IooTAgMBAAGjYjBgMBMGA1UdJQQMMAoGCCsGAQUF
'' SIG '' BwMDMEkGA1UdAQRCMECAEOsjZ+Gf5VyZWesnccFYdw+h
'' SIG '' GjAYMRYwFAYDVQQDEw1Sb25hbGQgQmxhbmNvghCIoKIY
'' SIG '' 7LM4gkC4AuEZ1K4dMA0GCSqGSIb3DQEBBQUAA4GBAH5t
'' SIG '' bFvwi+ZCAcxrs8KhHGxE5O4jpJDgiVmkY4YnHtdZeRU1
'' SIG '' 8TKMZjCp10mwb8MH3+C109JGzrnNNNY1bHnKbBxJG9nl
'' SIG '' GpX0d9EYCt6aUMUB4Z5WAQ3aWTMGOiJ89B/Gal7Rb9pt
'' SIG '' M4DbmYYnaExGVWrU1Hc17euRGp/sbQcd18czMYIBRDCC
'' SIG '' AUACAQEwLDAYMRYwFAYDVQQDEw1Sb25hbGQgQmxhbmNv
'' SIG '' AhCIoKIY7LM4gkC4AuEZ1K4dMAkGBSsOAwIaBQCgcDAQ
'' SIG '' BgorBgEEAYI3AgEMMQIwADAZBgkqhkiG9w0BCQMxDAYK
'' SIG '' KwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYB
'' SIG '' BAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUCNWCuz8D5HmT
'' SIG '' cq5gIQ8XP2n5z0MwDQYJKoZIhvcNAQEBBQAEgYC1Irp+
'' SIG '' VxVEVzkLYvllLHb9YzI61oa0aMTh3EVNTgI4l6hcfGTD
'' SIG '' +GwDcQMS8Bu1B8zB4RoI2MPir3B8B2zX3nvFgSboVxGq
'' SIG '' qvEs3sZBKGR9p7aCNdmM9cykUEfn9m9F0xi23LnRDDnC
'' SIG '' KXG5XyuM0H7D1CZmMfAZGRaJwlHg0+vHog==
'' SIG '' End signature block
