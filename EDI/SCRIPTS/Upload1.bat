@echo off

IF EXIST C:\PROEDI\XMITOUT\ATTOUT.EDI DEL C:\PROEDI\XMITOUT\ATTOUT.EDI >NUL

IF EXIST C:\PROEDI\XMITOUT\ATTOUT.DAT DEL C:\PROEDI\XMITOUT\ATTOUT.DAT >NUL

IF EXIST C:\PROEDI\EDIIN\ATTIN.DAT DEL C:\PROEDI\EDIIN\ATTIN.DAT >NUL

If Not exist \\dc-file-svr\ProEDI\Outbound\asnout.edi goto noasn

copy \\dc-file-svr\ProEDI\Outbound\asnout.edi C:\proediw\flatout\

copy \\dc-file-svr\ProEDI\Outbound\invout.edi C:\proediw\flatout\

If exist c:\proedi\flatout\invout.edi goto yesinv
If not exist c:\proedi\flatout\invout.edi goto noinvout

:yesinv

goto INVOUT

:INVOUT

goto AUDITW

:AUDITW

CD C:\PROEDIW

AUDITW 2,INVOICE,,INVOUT.EDI,,810,,

:noinvout

AUDITW 2,SHIP856,,ASNOUT.EDI,,856,,

:noasn 

AUDITW 4,MIPC,,,N,Y,,,,P,ATTOUT.EDI,W

REM PUT MINIMAL MTA MESSAGE ENVELOPE ON FILE
If Not Exist C:\PROEDI\XMITOUT\ATTOUT.EDI goto here
COPY C:\PROEDI\TOATT.EDI + C:\PROEDI\XMITOUT\ATTOUT.EDI C:\PROEDI\XMITOUT\ATTOUT.DAT
REM These are the commands to archive the files

:here
If Not Exist c:\proedi\flatout\asnout.edi goto no856 

:no856

goto 810OUT

:810OUT
If Not Exist c:\proedi\flatout\invout.edi goto NO810OUT

goto end

:NO810OUT

:end

exit




