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

Set objArgs = Wscript.Arguments
if objArgs(0) <> "" then
' Display the first 2 command-line arguments
	For III = 0 to 0
		'Wscript.Echo "File Selected: " & objArgs(III)
	Next
	arg1 = objArgs(0)
	'arg2 = objArgs(1)
end if

mymessage = InputBox("Enter message ", "Enter a msg")

WScript.Echo EMail( "MESSAGE FROM <" & mailfrom & ">", _
                    "MESSAGE TO <" & arg1 & ">", _
                    "Saludos", _
                    "TEXT!" & vbCrLf & "TEST, R", _
                    mymessage, _
                    "", _
                    server, _
                    port,_
					mailfrom,_
					pass)

Function EMail( myFrom, myTo, mySubject, myTextBody, myHTMLBody, myAttachment, mySMTPServer, mySMTPPort, myuser, mypass )
' This function sends an e-mail message using CDOSYS
'
' Arguments:
' myFrom       = Sender's e-mail address ("John Doe <jdoe@mydomain.org>" or "jdoe@mydomain.org")
' myTo         = Receiver's e-mail address ("John Doe <jdoe@mydomain.org>" or "jdoe@mydomain.org")
' mySubject    = Message subject (optional)
' myTextBody   = Actual message (text only, optional)
' myHTMLBody   = Actual message (HTML, optional)
' myAttachment = Attachment as fully qualified file name, either string or array of strings (optional)
' mySMTPServer = SMTP server (IP address or host name)
' mySMTPPort   = SMTP server port (optional, default 25)
'
' Returns:
' status message
'
' Written by Rob van der Woude
' http://www.robvanderwoude.com

    ' Standard housekeeping
    Dim i, objEmail

    ' Use custom error handling
    On Error Resume Next

    ' Create an e-mail message object
    Set objEmail = CreateObject( "CDO.Message" )
		
    ' Fill in the field values
    With objEmail
        .From     = myFrom
        .To       = myTo
        ' Other options you might want to add:
        ' .Cc     = ...
        ' .Bcc    = ...
        .Subject  = mySubject
        .TextBody = myTextBody
        .HTMLBody = myHTMLBody
        If IsArray( myAttachment ) Then
            For i = 0 To UBound( myAttachment )
                .AddAttachment Replace( myAttachment( i ), "\", "\\" ),"",""
            Next
        ElseIf myAttachment <> "" Then
            .AddAttachment Replace( myAttachment, "\", "\\" ),"",""
        End If
        If mySMTPPort = "" Then
            mySMTPPort = 25
        End If
        With .Configuration.Fields
            .Item( "http://schemas.microsoft.com/cdo/configuration/sendusing"      ) = 2
            .Item( "http://schemas.microsoft.com/cdo/configuration/smtpserver"     ) = mySMTPServer
            .Item( "http://schemas.microsoft.com/cdo/configuration/smtpserverport" ) = mySMTPPort
            .Item( "http://schemas.microsoft.com/cdo/configuration/smtpauthenticate" ) = 1
            .Item( "http://schemas.microsoft.com/cdo/configuration/sendusername" ) = myuser
            .Item( "http://schemas.microsoft.com/cdo/configuration/sendpassword" ) = mypass
            .Update
        End With
        ' Send the message
        .Send
    End With
    ' Return status message
    If Err Then
        EMail = "ERROR " & Err.Number & ": " & Err.Description
        Err.Clear
    Else
        EMail = "Message sent ok"
    End If

    ' Release the e-mail message object
    Set objEmail = Nothing
    ' Restore default error handling
    On Error Goto 0
End Function
'' SIG '' Begin signature block
'' SIG '' MIID8wYJKoZIhvcNAQcCoIID5DCCA+ACAQExCzAJBgUr
'' SIG '' DgMCGgUAMGcGCisGAQQBgjcCAQSgWTBXMDIGCisGAQQB
'' SIG '' gjcCAR4wJAIBAQQQTvApFpkntU2P5azhDxfrqwIBAAIB
'' SIG '' AAIBAAIBAAIBADAhMAkGBSsOAwIaBQAEFBIBA3pvSUSN
'' SIG '' vh4y/gW3Ohd+5cIPoIICGzCCAhcwggGAoAMCAQICEIig
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
'' SIG '' BAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUhXS2Ef/+QgXP
'' SIG '' AyfJN5rpMLx1j2gwDQYJKoZIhvcNAQEBBQAEgYCMVHZF
'' SIG '' gunelmONbpvMuXWvbxHn7+un7F2gXh5H6dMbxjBT1CO4
'' SIG '' gv098c9rIICGGpM9NbIKJYhOkJhNgCiMeW8/UOVbfDjh
'' SIG '' mLUbcw0mGiWnjrKqhDpRAqjdx8euE7EPbRH3JzUbayMU
'' SIG '' H0xb3Nt+d12LL8nfJ/QFJGMzqKnGYqLLqg==
'' SIG '' End signature block
