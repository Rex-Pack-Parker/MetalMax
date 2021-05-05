Dim CitraList, CPU_Type, CPU_TypeArr , Command
'--引用
Set WSS = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
Set WMI = GetObject("Winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")


Set Obj_CL = WScript.Arguments
'msgbox typename(Obj_CL)
For Each CMS In Obj_CL
	Command = Command & " " & CMS
Next

If Not LCase(Replace(WScript.FullName, WScript.Path & "\", "")) = "cscript.exe" Then
	WSS.Run "cmd /k mode con: lines=40 cols=80 & color 0F & CScript //nologo " & Chr(34) & WScript.ScriptFullName & Chr(34) & Command
	WSH.Sleep 1000
	WSH.Quit
End If

 NCD = Replace(WSH.ScriptFullName, WSH.ScriptName, "")
 Data = NCD & "Data\"

	For Each CMS In Obj_CL
		Select Case CMS
		Case "0":Model = 0
		Case "1":Model = 1
		Case Else
			Echo "--执行"
			Echo " 0:部署"
			Echo " 1:清除链接"
			Echo "----"
			Print "选择模式:" : Model = Input
		End Select
	Next

	If Model = "" Then 
		Echo "--执行"
		Echo " 0:部署"
		Echo " 1:清除链接"
		Echo "----"
		Print "选择模式:" : Model = Input
	End If
Select Case Model
	Case "0"
		Echo "部署模拟器虚拟连接.."
	Case "1"
		Echo "清除中.."
	Case Else
		WSH.Quit
End Select

For Each List in FSO.GetFolder(NCD).subfolders
	If Not List.Name = "Data" Then
		Dim InFolder, SourceFolder
		InFolder = List.Path & "\ROM"
		SourceFolder = Data & "_ROM"
		Print "处理:" & InFolder
		If FSO.FolderExists(PathCC(InFolder)) Then Runs("rd /s /q " &  PathC34(InFolder))
		If Model = "0" Then
			PrintL "  " & Runs("mklink /j " & PathC34(InFolder) & " " & PathC34(SourceFolder))
		Else
			PrintL "  完成"
		End If
		
		For Each ListUser in FSO.GetFolder(Data & "_User").subfolders
			Dim User, UserInFolder, UserSourceFolder
			UserInFolder = List.Path & "\User\" & ListUser.Name
			UserSourceFolder = ListUser.Path
			Print "处理:" & UserInFolder
			
			If FSO.FolderExists(PathCC(UserInFolder)) Then Runs("rd /s /q " &  PathC34(UserInFolder))
			If Model = "0" Then
				PrintL "  " & Runs("mklink /j " & PathC34(UserInFolder) & " " & PathC34(UserSourceFolder))
			Else
				PrintL "  完成"
			End If
		Next
	End If
Next
Print "结束.": Input

Function PathC34(ByVal GEOM_FP)
	If InStr(GEOM_FP, " ") = 0 Then
		PathC34 = Replace(GEOM_FP, Chr(34), "")
	Else
		PathC34 = IIf(Left(GEOM_FP, 1) = Chr(34), "", Chr(34)) & GEOM_FP & IIf(Right(GEOM_FP, 1) = Chr(34), "", Chr(34))
	End If
End Function

Function PathCC(ByVal GEOM_FP)
	PathCC = Replace(GEOM_FP, Chr(34), "")
End Function

Function Runs(ByVal GEOM_ProgramPath)
	'On Error Resume Next
	WSS.Run "cmd /c " & GEOM_ProgramPath, 0, True
	If Err Then Runs = Err.Description Else Runs = "完成"
End Function

Function IIf(ByVal GEOM_tf, ByVal GEOM_T, ByVal GEOM_F)
	If GEOM_tf = True Then IIF = GEOM_T Else IIF = GEOM_F
End Function

Sub Echo(Text)
	WSH.Echo Text
End Sub

Sub Print(ByVal Texts)
	WScript.StdOut.Write Texts
End Sub

Sub PrintL(ByVal Texts)
	WScript.StdOut.WriteLine(Texts)
End Sub

Function InPut()
	InPut = WScript.StdIn.ReadLine
End Function