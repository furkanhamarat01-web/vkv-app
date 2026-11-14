Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("Wscript.Shell")
dir = fso.GetParentFolderName(WScript.ScriptFullName)
pythonw = dir & "\python\pythonw.exe"
script = dir & "\app\baslat.py"
If Not fso.FileExists(pythonw) Then
  MsgBox "Uygulama dosyalari eksik. Klasorun tamamini Windows masaustune kopyalayin.", 16, "VKV Is Takip"
  WScript.Quit 1
End If
If Not fso.FolderExists(dir & "\Veriler") Then
  fso.CreateFolder dir & "\Veriler"
End If
env = shell.Environment("PROCESS")
env("IS_MERKEZI_HOME") = dir & "\Veriler"
env("TCL_LIBRARY") = dir & "\python\tcl\tcl8.6"
env("TK_LIBRARY") = dir & "\python\tcl\tk8.6"
env("PYTHONUTF8") = "1"
env("PYTHONDONTWRITEBYTECODE") = "1"
env("PYTHONIOENCODING") = "utf-8"
shell.CurrentDirectory = dir & "\python"
shell.Run """" & pythonw & """ """ & script & """", 0, False
