@echo off

CD C:\PROEDIW

ECHO CREATING THE 862 FLAT FILE FOR FUSION IMPORT.

AUDITW 1,SHIP862,,IN862.EDI,,862,N,,

:end

exit