Dim CitraList, CPU_Type, CPU_TypeArr , Command
'--����
Set WSS = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
Set WMI = GetObject("Winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

Const GEOM_Title = "Metal Dogs PT-����    by:Rex.Pack v:1.0"
Const GEOM_Color = "06"

Set Obj_CL = WScript.Arguments
'msgbox typename(Obj_CL)
For Each CMS In Obj_CL
	Command = Command & " " & CMS
Next

If Not LCase(Replace(WScript.FullName, WScript.Path & "\", "")) = "cscript.exe" Then
	WSS.Run "cmd /k mode con: lines=40 cols=90 & color " & GEOM_Color & " & Title " & GEOM_Title & " & CScript //nologo " & Chr(34) & WScript.ScriptFullName & Chr(34) & Command
	WSH.Sleep 1000
	WSH.Quit
End If

NCD = Replace(WSH.ScriptFullName, WSH.ScriptName, "")
Data = NCD & "_DATA\"
AppData = "%LOCALAPPDATA%Low\24Frame,Inc_\"

For Each CMS In Obj_CL
	Select Case CMS
	Case "0":Model = 0
	Case "1":Model = 1
	Case Else
		
	End Select
Next

If Model = "" Then 
	Echo "--ִ��"
	Echo " 0:����"
	Echo " 1:�������"
	Echo "----"
	Print "ѡ��ģʽ:" : Model = Input
End If

Select Case Model
	Case "0"
		Echo "����.."
		Echo Runs("md " & PathC34(AppData))
		WSH.Sleep 500
		Echo "ɾ��AppData��MD�ļ���\�����ļ���:  " & "rd /q " &  PathC34(AppData & "METAL DOGS")
		Runs("rd /q " &  PathC34(AppData & "METAL DOGS"))
		Echo Runs("mklink /j " & PathC34(AppData & "METAL DOGS") & " " & PathC34(NCD & "_DATA\"))
		Echo "����ԱȨ��: " & PathC34(NCD & "METAL DOGS Playtest world.reg")
		Echo Runs("regedit /s " & PathC34(NCD & "METAL DOGS Playtest world.reg"))
	Case "1"
		Echo "���.."
		Runs("rd /s /q " &  PathC34(AppData & "METAL DOGS"))
	Case Else
		WSH.Quit
End Select


Print "����.": Input

Function PathC34(ByVal GEOM_FP)
	PathC34 = IIf(Left(GEOM_FP, 1) = Chr(34), "", Chr(34)) & GEOM_FP & IIf(Right(GEOM_FP, 1) = Chr(34), "", Chr(34))
End Function

Function PathCC(ByVal GEOM_FP)
	PathCC = Replace(GEOM_FP, Chr(34), "")
End Function

Function Runs(ByVal GEOM_ProgramPath)
	'On Error Resume Next
	WSS.Run "cmd /c " & GEOM_ProgramPath, 0, True
	If Err Then Runs = Err.Description Else Runs = "��� " & GEOM_ProgramPath
End Function

Function IIf(ByVal GEOM_tf, ByVal GEOM_T, ByVal GEOM_F)
	If GEOM_tf = True Then IIF = GEOM_T Else IIF = GEOM_F
End Function

Function tfFolder(ByVal Path)
	If FSO.FolderExists(PathC34(Path)) Then tfFolder = True Else tfFolder = False
End Function

Sub Echo(Text)
	WSH.Echo Text
End Sub

Sub Print(ByVal Texts)
	WScript.StdOut.Write Texts
End Sub

Sub PrintC(ByVal Texts)
	WScript.StdOut.Write Chr(13) & Texts
End Sub

Sub PrintL(ByVal Texts)
	WScript.StdOut.WriteLine(Texts)
End Sub

Function InPut()
	InPut = WScript.StdIn.ReadLine
End Function