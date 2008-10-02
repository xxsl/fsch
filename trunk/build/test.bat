@ECHO OFF
REM call the app that can set the errorlevel
compiler.exe index
REM higher errorlevels must be tested first
IF ERRORLEVEL 3 GOTO e3
IF ERRORLEVEL 2 GOTO e2
IF ERRORLEVEL 1 GOTO e1

REM Add here code for Errorlevel = 0
echo 0
GOTO End

:e1
echo 1
REM ...
GOTO End

:e2
echo 2
REM ...
GOTO End

:e3
echo 3
REM ...

:End
