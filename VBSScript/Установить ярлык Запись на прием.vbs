'************************************************************ 
'* ���: Shortcut_Create.vbs                                 * 
'* ����: VBScript                                           * 
'* ����������: �������� � ��������� ������ �� ������� ����� *
'************************************************************
Set objWShell = CreateObject("WScript.Shell")
Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(DESKTOP)
Set objDesktop = objShell.NameSpace(&H0)
desktopDir = objDesktop.Self.Path & "\"
Set objFolderItem = objDesktop.Self
Set fso = CreateObject("Scripting.FileSystemObject")
link = "������ �� �����.lnk"
If (fso.FileExists(objFolderItem.Path & "\" & LINK)=false) Then
Set objSC = objWShell.CreateShortcut(desktopDir & "������ �� �����.lnk")
objSC.Description = "���������"
objSC.IconLocation = "\\192.168.0.100\netlogon\label\user.ico"
objSC.TargetPath = "\\192.168.0.211\Froda\On_Eib_C\raspis.vbs"
objSC.WindowStyle = 1
objSC.WorkingDirectory = "\\192.168.0.211\Froda\On_Eib_C"
objSC.Save
else
'MsgBox("����� ��� ����������")
end if
Set objSC = Nothing
Set objWShell = Nothing
