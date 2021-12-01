@echo off

IF EXIST C:\PROEDIW\FLATIN\IN830.EDI Goto next1

If not exisxt C:\PROEDIW\FLATIN\in830.edi Goto next2

:next1

If Exist \\dc-file-svr\ProEdi\Inbound\in830.edi Del \\dc-file-svr\ProEdi\Inbound\in830.edi

Copy C:\PROEDIW\FLATIN\IN830.EDI \\DC-FILE-SVR\ProEDI\Inbound

:next2

IF EXIST C:\PROEDIW\FLATIN\IN850.EDI Goto next3

If not exisxt C:\PROEDIW\FLATIN\in850.edi Goto next4

:next3

If Exist \\dc-file-svr\ProEdi\Inbound\in850.edi Del \\dc-file-svr\ProEdi\Inbound\in850.edi

Copy C:\PROEDIW\FLATIN\IN850.EDI \\DC-FILE-SVR\ProEDI\Inbound

:next4

IF EXIST C:\PROEDIW\FLATIN\IN862.EDI Goto next5

If not exisxt C:\PROEDIW\FLATIN\in862.edi Goto next6

:next5

If Exist \\dc-file-svr\ProEdi\Inbound\in862.edi Del \\dc-file-svr\ProEdi\Inbound\in862.edi

Copy C:\PROEDIW\FLATIN\IN862.EDI \\DC-FILE-SVR\ProEDI\Inbound

:next6

IF EXIST C:\PROEDIW\EDIIN\UA0000.000 Goto next7

If not exisxt C:\PROEDIW\EDIIN\UA0000.000 Goto end

:next7

If Exist \\dc-file-svr\ProEdi\Inbound\UA0000.000 Del \\dc-file-svr\ProEdi\Inbound\UA0000.000

Copy C:\PROEDIW\EDIIN\UA0000.000 \\DC-FILE-SVR\ProEDI\Inbound

:end

exit
