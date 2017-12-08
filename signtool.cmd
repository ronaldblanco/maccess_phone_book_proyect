call "env.cmd"
cd ..
cd ..
cd ..
cd ..
C:
cd %foldersign%

SignTool sign /f %foldercert% %folder%env.vbs
SignTool sign /f %foldercert% %folder%SEND_MAIL.vbs
SignTool sign /f %foldercert% %folder%SEND_MAIL_OUTLOOK.vbs
SignTool sign /f %foldercert% %folder%X_LITE_CALL.exe
SignTool sign /f %foldercert% %folder%X_LITE_CALL_CODE.exe
SignTool sign /f %foldercert% %folder%curl\000CALL_TO_A_NUMBER.vbs
pause