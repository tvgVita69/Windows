Set objShell = CreateObject("Shell.Application")
Set objDesktop = objShell.NameSpace(&H0)
desktopDir = objDesktop.Self.Path & "\"
Set objDesktop = Nothing
Set objShell = Nothing
Set objWShell = CreateObject("WScript.Shell")
Set objSC = objWShell.CreateShortcut(desktopDir & "���������.lnk")
objSC.Description = "���������"
objSC.IconLocation = "\\192.168.0.100\netlogon\label\Text.ico"
objSC.TargetPath = "\\192.168.0.2\Documents\Clinica\��������\���� 6\������� 608\���������"
objSC.Save
Set objWShell = Nothing