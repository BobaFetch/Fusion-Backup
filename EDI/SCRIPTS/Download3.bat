@echo off

CD C:\PROEDIW

ECHO CREATING THE 830 FLAT FILE FOR FUSION IMPORT.

AUDITW 1,SHIP830,,IN830.EDI,,830,N,,

:end

exit