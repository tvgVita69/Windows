'************************************************************ 
'* ���: Shortcut_Create.vbs                                 * 
'* ����: VBScript                                           * 
'* ����������: �������� � ��������� ������ �� ������� ����� *
'************************************************************

Set objShell = CreateObject("Shell.Application")
Set objDesktop = objShell.NameSpace(&H0)
desktopDir = objDesktop.Self.Path & "\"
Set objDesktop = Nothing
Set objShell = Nothing
Set objWShell = CreateObject("WScript.Shell")
Set objSC = objWShell.CreateShortcut(desktopDir & "���������.lnk")
objSC.Description = "���������"
objSC.IconLocation = "\\192.168.0.100\netlogon\label\dllmar.ico"
objSC.TargetPath = "\\192.168.0.211\Froda\On_Eib_C\dll_mar.vbs"
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
Set objSC = objWShell.CreateShortcut(desktopDir & "����� �����.lnk")
objSC.IconLocation = "\\192.168.0.100\netlogon\label\Marketing.ico"
objSC.TargetPath = "\\192.168.0.2\Laboratoria\Marketing"
objSC.Save
Set objWShell = Nothing