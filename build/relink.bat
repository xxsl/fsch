rem relink to use stdout
del server.exe
del client.exe
copy fcshWrapper.exe server.exe /y
LINK.EXE /EDIT /SUBSYSTEM:CONSOLE c:\work\trunk\build\fcshWrapper.exe
copy fcshWrapper.exe client.exe /y
del fcshWrapper.exe
