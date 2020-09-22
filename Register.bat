cls
@echo off
@echo Registering DNS.ocx to '%SystemRoot%' . . .
copy DNS.ocx %SystemRoot%\ /y
@echo Finished.
@echo ---------
@echo If this doesn't work, go into your '\Windows\System32' folder and paste DNS.ocx onto the 'regsvr32.exe' program.