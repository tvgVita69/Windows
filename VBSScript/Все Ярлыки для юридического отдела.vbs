Dim SavePsw 
Set SavePsw = CreateObject("WScript.Shell")
SavePsw.Run "cmd.exe /c chcp 1251&&net use \\labserver /user:moscow\�������������_ begin",0,true
SavePsw.Run "cmd.exe /c chcp 1251&&net use \\192.168.0.2 /user:moscow\�������������_ begin",0,true
SavePsw.Run "cmd.exe /c chcp 1251&&net use \\win2kserv.moscow.onclinic.microsoft.com /user:moscow\�������������_ begin",0,true
SavePsw.Run "cmd.exe /c chcp 1251&&net use \\192.168.0.100 /user:moscow\�������������_ begin",0,true
SavePsw.Run "cmd.exe /c chcp 1251&&net use /persistent:yes",0,true
Set SavePsw = Nothing



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



Set objDesktop = Nothing
Set objShell = Nothing
Set objWShell = CreateObject("WScript.Shell")
Set objSC = objWShell.CreateShortcut(desktopDir & "������� �������.lnk")
objSC.Description = "������� �������"
objSC.TargetPath = "\\192.168.0.2\Documents\Clinica\������� �������"
objSC.IconLocation = "\\192.168.0.100\netlogon\label\������� �������.ico"
objSC.Save
Set objWShell = Nothing



Set objShell = CreateObject("Shell.Application")
Set objDesktop = objShell.NameSpace(&H0)
desktopDir = objDesktop.Self.Path & "\"
Set objDesktop = Nothing
Set objShell = Nothing
Set objWShell = CreateObject("WScript.Shell")
Set objSC = objWShell.CreateShortcut(desktopDir & "������ �� �����.lnk")
objSC.Description = "���������"
objSC.IconLocation = "\\192.168.0.100\netlogon\label\user.ico"
objSC.TargetPath = "\\192.168.0.211\Froda\On_Eib_C\raspis.vbs"
objSC.WindowStyle = 1
objSC.WorkingDirectory = "\\192.168.0.211\Froda\On_Eib_C"
objSC.Save
Set objSC = Nothing
Set objWShell = Nothing

Set objShell = CreateObject("Shell.Application")
Set objDesktop = objShell.NameSpace(&H0)
desktopDir = objDesktop.Self.Path & "\"
Set objDesktop = Nothing
Set objShell = Nothing
Set objWShell = CreateObject("WScript.Shell")
Set objSC = objWShell.CreateShortcut(desktopDir & "���.lnk")
objSC.Description = "���������"
objSC.IconLocation = "\\192.168.0.100\netlogon\label\TREESURG.ICO"
objSC.TargetPath = "\\192.168.0.211\Froda\On_Eib_C\eib.vbs"
objSC.WindowStyle = 1
objSC.WorkingDirectory = "\\192.168.0.211\Froda\On_Eib_C"
objSC.Save
Set objSC = Nothing
Set objWShell = Nothing

Set objShell = CreateObject("Shell.Application")
Set objDesktop = objShell.NameSpace(&H0)
desktopDir = objDesktop.Self.Path & "\"
Set objDesktop = Nothing
Set objShell = Nothing
Set objWShell = CreateObject("WScript.Shell")
Set objSC = objWShell.CreateShortcut(desktopDir & "�����������.lnk")
objSC.Description = "���������"
objSC.IconLocation = "\\192.168.0.100\netlogon\label\blood.ico"
objSC.TargetPath = "\\192.168.0.211\Froda\On_Eib_C\pusk.vbs"
objSC.WindowStyle = 1
objSC.WorkingDirectory = "\\192.168.0.211\Froda\On_Eib_C"
objSC.Save
Set objSC = Nothing
Set objWShell = Nothing


Set objShell = CreateObject("Shell.Application")
Set objDesktop = objShell.NameSpace(&H0)
desktopDir = objDesktop.Self.Path & "\"
Set objDesktop = Nothing
Set objShell = Nothing
Set objWShell = CreateObject("WScript.Shell")
Set objSC = objWShell.CreateShortcut(desktopDir & "����� ������.lnk")
objSC.Description = "����� ������"
objSC.IconLocation = "\\192.168.0.100\netlogon\label\Edit Text.ico"
objSC.TargetPath = "\\192.168.0.2\Pub\����� ������"
objSC.Save
Set objWShell = Nothing

Set objShell = CreateObject("Shell.Application")
Set objDesktop = objShell.NameSpace(&H0)
desktopDir = objDesktop.Self.Path & "\"
Set objDesktop = Nothing
Set objShell = Nothing
Set objWShell = CreateObject("WScript.Shell")
Set objSC = objWShell.CreateShortcut(desktopDir & "��������� �� ������.lnk")
objSC.Description = "���������"
objSC.IconLocation = "\\192.168.0.100\netlogon\label\onclinic.ico"
objSC.TargetPath = "\\192.168.0.211\Froda\On_Eib_C\���������\������� �������.hta"
objSC.WindowStyle = 1
objSC.WorkingDirectory = "\\192.168.0.211\Froda\On_Eib_C\���������"
objSC.Save
Set objSC = Nothing
Set objWShell = Nothing

Set objShell = CreateObject("Shell.Application")
Set objDesktop = objShell.NameSpace(&H0)
desktopDir = objDesktop.Self.Path & "\"
Set objDesktop = Nothing
Set objShell = Nothing
Set objWShell = CreateObject("WScript.Shell")
Set objSC = objWShell.CreateShortcut(desktopDir & "������ ���������.lnk")
objSC.Description = "������ ���������"
objSC.IconLocation = "\\192.168.0.100\netlogon\label\numbers.ico"
objSC.TargetPath = "\\192.168.0.211\Froda\On_Eib_C\���������\spisok.hta"
objSC.WindowStyle = 1
objSC.WorkingDirectory = "\\192.168.0.211\Froda\On_Eib_C\���������"
objSC.Save
Set objSC = Nothing
Set objWShell = Nothing

