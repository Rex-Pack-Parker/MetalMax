@Echo Off
Title Citra-User�浵����    -�׿�˹.�� Ver:1.7
Mode Con: Cols=128 Lines=32
Color 0A
CD /d %~dp0
Set SFAD=%AppData%\Citra\
Set NOWD=%~dp0Data\_User\

Echo [ת��ע��]
Echo !����������򽫸��ļ�����ģ�����ļ�����������. ��Ҫ~!
Echo �����Ը��ƽ���,����ʾԴλ���滻Ŀ���ļ�λ��.
Echo [�����ļ������]
Echo  http://window.es-geom.com/
Echo  ftp://mm@s.es-geom.com/

Echo *
Echo --------[���]
Echo +��ǰλ�� %NOWD%
If Exist "%NOWD%" (Echo  ������) Else (Echo  ��Ч!)

Set RootPath=%NOWD%
Set WorkPath=%NOWD%sdmc\Nintendo 3DS\00000000000000000000000000000000\00000000000000000000000000000000\title\

Set Cheat=%RootPath%Cheats\
Set SAVE=%WorkPath%00040000
Set DLC=%WorkPath%0004008c
Set Patch=%WorkPath%0004000e
Set MOD=%RootPath%load\mods

:Main
	Pause
	Cls
	
	Echo --------[����]
	Echo +ѡ��:
	Echo  0. (���)�˳�
	Echo  1. ��� ��ǰ
	Echo  2. ���� ��ǰ�浵 '����������/�ļ���
	rem Echo  3. 
	rem Echo  4. 
	rem Echo  5. 
	rem Echo  6. 
	Echo  ---- [MM4]
	Echo  a. ��� ����ָ
	Echo  b.  ʹ���Զ�����ָ   '��Ҫ�ر�ģ����
	Echo  c.  ʹ��1.0����ָ    '��Ҫ�ر�ģ����
	Echo  d.  ʹ��1.1����ָ    '��Ҫ�ر�ģ����
	Echo  e. ��� ��ҪĿ¼
	Echo  f.  ��� �浵
	If Exist "%DLC%" (
		Echo  g.  ���� DLC   '��Ҫ�ر�ģ����
	) Else (
		Echo  g.  ���� DLC   '��Ҫ�ر�ģ����
	)
	If Exist "%DLC%" ( Echo  h.   ��� DLC)
	If Exist "%Patch%" (
		Echo  i.  ���� ����  '��Ҫ�ر�ģ����
	) Else (
		Echo  i.  ���� ����  '��Ҫ�ر�ģ����,ע�ⲹ��Ӱ�����ָ
	)
	If Exist "%Patch%" ( Echo  j.   ��� ����)
	
	Echo  k. ɾ��MOD\romfs               '��Ҫ�ر�ģ����
	Echo  l.  MOD ʹ��515�汾����        '��Ҫ�ر�ģ����
	Echo  m.  MOD ʹ��515�汾���� ����   '��Ҫ�ر�ģ����
	Echo ----ѡ��:
	Set Select=
	Set /p Select=����:
	
	If %Select%==0 GoTo Bye
	If %Select%==1 GoTo OpenNOWD
	If %Select%==2 GoTo FileListNOWD
	rem If %Select%==3 GoTo 
	rem If %Select%==4 GoTo 
	rem If %Select%==5 GoTo 
	rem If %Select%==6 GoTo 
	
	If %Select%==a GoTo Open-Cheat
	If %Select%==b GoTo To-Cheat-DIY
	If %Select%==c GoTo To-Cheat-10
	If %Select%==d GoTo To-Cheat-11
	If %Select%==e GoTo Open-Work
	If %Select%==f GoTo Open-SAVE
	If %Select%==g GoTo ED-DLC
	If %Select%==h GoTo Open-DLC
	If %Select%==i GoTo ED-Patch
	If %Select%==j GoTo Open-Patch
	If %Select%==k GoTo To-MOD-del
	If %Select%==l GoTo To-MOD-515-0
	If %Select%==m GoTo To-MOD-515-1
	
GoTo Main


:OpenSFAD
	Explorer "%SFAD%"
GoTo Main

:OpenNOWD
	Explorer "%NOWD%"
GoTo Main

:FileListSFAD
	Dir "%WorkPath1%\*lot_*" /s /t:w /o:d
GoTo Main

:FileListNOWD
	Dir "%WorkPath%\*lot_*" /s /t:w /o:d
GoTo Main

:SFAD2NOWD
	XCopy "%SFAD%%SAVE%*.*" "%NOWD%%SAVE%" /h /e /y
GoTo Main

:NOWD2SFAD
	XCopy "%NOWD%%SAVE%*.*" "%SFAD%%SAVE%" /h /e /y
GoTo Main


:Open-Cheat
	Explorer "%Cheat%"
GoTo Main

:Open-Work
	Explorer "%WorkPath%"
GoTo Main

:Open-SAVE
	Explorer "%SAVE%\000afd00\data\00000001\"
GoTo Main

:Open-DLC
	Explorer "%DLC%\000afd00\content\"
GoTo Main

:Open-Patch
	Explorer "%Patch%\000afd00\content\"
GoTo Main


:ED-DLC
	CD %WorkPath%
	If Exist "0004008c" (
		ren "0004008c" "!0004008c"
	) Else (
		ren "!0004008c" "0004008c"
	)
	CD %~dp0
GoTo Main

:ED-Patch
	CD %WorkPath%
	If Exist "0004000e" (
		ren "0004000e" "!0004000e"
	) Else (
		ren "!0004000e" "0004000e"
	)
	CD %~dp0
GoTo Main

:ED-Cheat
	CD %WorkPath%
	If Exist "0004000e" (
		
	) Else (
		
	)
	CD %~dp0
GoTo Main

:To-Cheat-DIY
	CD %Cheat%
	del /q 00040000000AFD00.txt
	mklink /h 00040000000AFD00.txt 00040000000AFD00-DIY.txt
	CD %~dp0
GoTo Main


:To-Cheat-10
	CD %Cheat%
	del /q 00040000000AFD00.txt
	mklink /h 00040000000AFD00.txt 00040000000AFD00-1.0.txt
	CD %~dp0
GoTo Main

:To-Cheat-11
	CD %Cheat%
	del /q 00040000000AFD00.txt
	mklink /h 00040000000AFD00.txt 00040000000AFD00-1.1.txt
	CD %~dp0
GoTo Main

:To-MOD-del
	rd /s /q %MOD%\00040000000AFD00\romfs
GoTo Main

:To-MOD-515-0
	CD %MOD%\00040000000AFD00\
	rd /s /q romfs
	mklink /j romfs romfs-515-0
	CD %~dp0
GoTo Main

:To-MOD-515-1
	CD %MOD%\00040000000AFD00\
	rd /s /q romfs
	mklink /j romfs romfs-515-1
	CD %~dp0
GoTo Main

:Bye
Pause