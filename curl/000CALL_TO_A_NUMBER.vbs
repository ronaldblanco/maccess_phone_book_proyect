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
	For III = 0 to 1
		'Wscript.Echo "File Selected: " & objArgs(III)
	Next
	arg1 = objArgs(0)
	arg2 = objArgs(1)
	'Set myFile = objFSO.OpenTextFile(filePath, ForReading, True)
	'Set myTemp= objFSO.OpenTextFile(filePath & ".tmp", ForWriting, True)
end if

Dim objShell
Set objShell = WScript.CreateObject("WScript.Shell")
objShell.Run cmdcall & arg1 & " " & arg2



'' SIG '' Begin signature block
'' SIG '' MIID8wYJKoZIhvcNAQcCoIID5DCCA+ACAQExCzAJBgUr
'' SIG '' DgMCGgUAMGcGCisGAQQBgjcCAQSgWTBXMDIGCisGAQQB
'' SIG '' gjcCAR4wJAIBAQQQTvApFpkntU2P5azhDxfrqwIBAAIB
'' SIG '' AAIBAAIBAAIBADAhMAkGBSsOAwIaBQAEFK+mwzrfOTqV
'' SIG '' w8z4pMP8qkoazmeJoIICGzCCAhcwggGAoAMCAQICEIig
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
'' SIG '' BAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUgKeVma8L3gGT
'' SIG '' KLnI33tVeudM1PcwDQYJKoZIhvcNAQEBBQAEgYBmpqKn
'' SIG '' HaydTyruMONp2RmihjWUGYXizeKsMp1Xv+9XBVUh4ZaA
'' SIG '' kjs04P22oF+GMXQkdfaxM58sJD/uCw81yVl9UODcPGgC
'' SIG '' qlkMBwJqJP/JX6Ep/P+9QFhMSYggq/iEnG9+eE+I8jyX
'' SIG '' jW4rWDURKgg/H437BcCl6w+XyBE5NVW36g==
'' SIG '' End signature block
