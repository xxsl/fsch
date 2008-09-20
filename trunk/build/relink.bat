rem relink to use stdout
del shell.exe
del compiler.exe
copy fcshWrapper.exe shell.exe /y
LINK.EXE /EDIT /SUBSYSTEM:CONSOLE c:\work\trunk\build\fcshWrapper.exe
copy fcshWrapper.exe compiler.exe /y
del fcshWrapper.exe
