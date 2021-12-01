@echo off
echo.|time|find "Current" >cu##ent.bat
echo set time=%%3> current.bat
call cu##ent.bat
:: At this point the time is in the environment but has colons in it
:: CU??ENT.BAT are no longer needed and will be overwritten below
echo = | choice /c=%time%= cu##ent.bat > current.bat
echo @echo off> cu##ent.bat
echo set time=>> cu##ent.bat
echo :LOOP>> cu##ent.bat
echo shift>> cu##ent.bat
echo if "%%1"=="]?" goto DONE>> cu##ent.bat
echo if "%%1"==":" goto COLON>> cu##ent.bat
echo goto NOTCOLON>> cu##ent.bat
echo :COLON>> cu##ent.bat
echo set time=%%time%%.>> cu##ent.bat
echo goto LOOP>> cu##ent.bat
echo :NOTCOLON>> cu##ent.bat
echo set time=%%time%%%%1>> cu##ent.bat
echo goto LOOP>> cu##ent.bat
echo :DONE>> cu##ent.bat
call current.bat
del cu??ent.bat > nul
