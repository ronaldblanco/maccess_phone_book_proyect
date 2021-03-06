#region --- Au3Recorder generated code Start (v3.3.9.5 KeyboardLayout=00000409)  ---

#region --- Internal functions Au3Recorder Start ---
Func _Au3RecordSetup()
Opt('WinWaitDelay',100)
Opt('WinDetectHiddenText',1)
Opt('MouseCoordMode',0)
Local $aResult = DllCall('User32.dll', 'int', 'GetKeyboardLayoutNameW', 'wstr', '')
If $aResult[1] <> '00000409' Then
  MsgBox(64, 'Warning', 'Recording has been done under a different Keyboard layout' & @CRLF & '(00000409->' & $aResult[1] & ')')
EndIf

EndFunc

Func _WinWaitActivate($title,$text,$timeout=0)
	WinWait($title,$text,$timeout)
	If Not WinActive($title,$text) Then WinActivate($title,$text)
	WinWaitActive($title,$text,$timeout)
EndFunc

_AU3RecordSetup()
#endregion --- Internal functions Au3Recorder End ---


_WinWaitActivate("Program Manager","")
MouseClick("right",922,178,2)
_WinWaitActivate("X-Lite","")

$i = 1
$iValue = ""
for $j = 1 to (StringLen(String($CmdLine[$i])))
	$iValue = StringMid (String($CmdLine[$i]), $j , 1)


 Select
        Case $iValue = "1"
            MouseClick("right",54,200,1)
        Case $iValue = "2"
            MouseClick("right",156,200,1)
		Case $iValue = "3"
            MouseClick("right",253,200,1)
		Case $iValue = "4"
            MouseClick("right",54,262,1)
		Case $iValue = "5"
            MouseClick("right",156,262,1)
		Case $iValue = "6"
            MouseClick("right",253,262,1)
		Case $iValue = "7"
            MouseClick("right",54,317,1)
		Case $iValue = "8"
            MouseClick("right",156,317,1)
		Case $iValue = "9"
            MouseClick("right",253,317,1)
		Case $iValue = "*"
            MouseClick("right",54,375,1)
		Case $iValue = "0"
            MouseClick("right",156,375,1)
		Case $iValue = "#"
            MouseClick("right",253,375,1)
        Case Else ; If nothing matches then execute the following.
            MsgBox($MB_SYSTEMMODAL, "", "No preceding case was True.")
EndSelect

Next

MouseClick("right",259,144,1)

_WinWaitActivate("Program Manager","")
MouseMove(907,190)
MouseDown("right")
MouseMove(909,190)
MouseUp("right")
MouseClick("right",909,190,1)
#endregion --- Au3Recorder generated code End ---
