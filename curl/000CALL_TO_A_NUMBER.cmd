call "C:\00000directory_project\curl\env.cmd"
echo %url%

set arg1=%1
set arg2=%2
set str2=%url%
set str3="&destination="
cd C:\OpenSSL\bin\_Firmar_xml_ec
curl %str2%%arg1%%str3%%arg2%