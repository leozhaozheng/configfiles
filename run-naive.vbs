DIM objShell
set objShell=wscript.createObject("wscript.shell")
iReturn=objShell.Run("cmd.exe /C D:\Apps\naiveproxy\naive.exe D:\Apps\naiveproxy\config.json", 0, TRUE)