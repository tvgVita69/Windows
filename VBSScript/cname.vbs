set objNetwork = createobject("wscript.network") 
set WShell = WScript.CreateObject("WScript.Shell") 
set fso = createobject("scripting.FileSystemObject") 
Set WshNetwork = CreateObject("WScript.Network")
'����� ����� ������������, ���������� � ����� �� ����� �� ������
info =  "��� ����������: " & WshNetwork.ComputerName
info = info & Chr(10)
info = info & Chr(10)
info = info & "��� ������������: " & WshNetwork.UserName
info = info & Chr(10)
info = info & Chr(10)
info = info & "�����: " & WshNetwork.UserDomain
WScript.Echo info
pts2="\\192.168.0.100\netlogon\compname" '���� ���� � ����, ��� ����� ������ cname.vbs
sname="cname.vbs" 
profpath=WShell.SpecialFolders("Desktop")
compname=objNetwork.ComputerName 
LinkName= profpath & "\ " & compname & ".lnk" 
   if (fso.fileexists (LinkName) = true)  then              
	WScript.Quit
   end if   
   
   if fso.Folderexists (profpath)=true then 
      profpath=profpath
   elseif fso.Folderexists (profpath)=true then 
      profpath=profpath
   else 
      WScript.Quit 
   end if 


 
      
set CNLink=WShell.CreateShortcut(LinkName) 
CNLink.TargetPath = pts2 & "\" & sname 
CNLink.Description = "���" 
CNLink.IconLocation = "\\192.168.0.100\netlogon\compname\name_computer.ico"
CNLink.WorkingDirectory = ""
CNLink.Hotkey = "CTRL+SHIFT+1" 
CNLink.save 

