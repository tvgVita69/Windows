Dim Shell, strCommand, strHost, ReturnCode, oShell, objExec, objShell
    strHost = "192.168.0.211" 
    '����������� ���������� ��� �����
Set Shell = wscript.createObject("wscript.shell") 
'������� ������ ��� ���������� ������� �������
Set oShell = CreateObject("Shell.Application") 
'��������� �� ������ �����
'Set FSO = CreateObject("Scripting.FileSystemObject")
'If FSO.FileExists("C:\miss\" & WScript.ScriptName & ".txt") Then
 '   MsgBox "����� ��� �������"
 '  WScript.Quit
'End If
'FSO.CreateTextFile "C:\miss\" & WScript.ScriptName & ".txt"

'��������� ��������� ��������
'Set WshShell = CreateObject ("WScript.Shell")
'WshShell.Run "indicator_processa.hta", 1 , TRUE

'������� ������ ��� ���������� ��������	
Set objShell = CreateObject("WScript.Shell")
Set objExec = objShell.Exec("\\192.168.0.211\froda\On_Eib_C\labos009.exe -t -c\\192.168.0.211\froda\On_Eib_C\config.fpw") 
'WshShell.Popup "���������, ���� ���������� � �����", 15, "�������� '�����������'", vbInformation
Do While true 
    if objExec.Status = 1  then
	objShell.Run "taskkill.exe /f /im labos*", 0
	Wscript.Quit 1
	 End if
    '��������� ���������� ������� �� �����
     strCommand = "ping -n 4 -l 1 " & strHost 
    '����������� ���������� ���������� �� ���� ������ �����
     ReturnCode = Shell.Run(strCommand, 0, True)
    '���������� ���
        If ReturnCode = 0 Then 
            '0 - �������� ����������������� ������� 1 - ������� �� �����������
            strCommand = true  
        Else
		oShell.ShellExecute "taskkill.exe", "/f /im labos*", , , 0  
		MsgBox  "��������!!!"& chr(10) &"����� �������� ����������. ���� ����� �� ��������." & chr(10) & "������� ��, ����� �������� ��������� ���������.",65584
		WScript.Quit()
	 End if
                    	
Loop