
'---- init
Const Title = "yuzu-部署 v:1.0"
Const Color = "06"

Set Core = CreateObject("Scripting.FileSystemObject").OpenTextFile(Replace(WSH.ScriptFullName, WSH.ScriptName, "") & "Core.GEOM", 1, False)
Execute Core.ReadAll

'---- main

Dim CitraList, CPU_Type, CPU_TypeArr

Model = ""
Data = ScriptPath & "Data\"

For Each CMS In Split(Command.All)
	Select Case CMS
	Case "0":Model = "0"
	Case "1":Model = "1"
	Case Else
	End Select
Next

Echo "--执行"
Echo " 0:部署"
Echo " 1:清除链接"

Call Main
Sub Main()
	If Model = "" Then 
		Echo ""
		CScript.PrintI "选择模式:" : Model = CScript.Input
	End If
	
	Select Case Model
	Case ""
	Case "0", "1"
		Select Case Model
			Case "0": Echo "部署模拟器虚拟连接.."
			Case "1": Echo "清除中.."
			Case Else
				'WSH.Quit
		End Select

		For Each List in FSO.GetFolder(ScriptPath).SubFolders
			Select Case List.Name
			Case "Data", "数据"
			Case Else
				Dim InFolder, SourceFolder
				InFolder = List.Path & "\ROM"
				SourceFolder = Data & "_ROM"
				CScript.PrintL "处理 ROM:	" & InFolder
				If FSO.FolderExists(PathCC(InFolder)) Then Call ReDir(PathC34(InFolder))
				If Model = "0" Then Call MKlink(PathC34(InFolder), PathC34(SourceFolder), False)
				
				InFolder = List.Path & "\User"
				SourceFolder = Data & "_User"
				CScript.PrintL "处理 User:	" & InFolder
				If FSO.FolderExists(PathCC(InFolder)) Then Call ReDir(PathC34(InFolder))
				If Model = "0" Then Call MKlink(PathC34(InFolder), PathC34(SourceFolder), False)
				'For Each ListUser in FSO.GetFolder(Data & "_User").SubFolders
				'	Dim UserInFolder, UserSourceFolder
				'	UserInFolder = List.Path & "\User\" & ListUser.Name
				'	UserSourceFolder = ListUser.Path
				'	CScript.PrintL "处理 User:	" & UserInFolder
				'	If FSO.FolderExists(PathCC(UserInFolder)) Then Call ReDir(PathC34(UserInFolder))
				'	If Model = "0" Then Call MKlink(PathC34(UserInFolder), PathC34(UserSourceFolder), False)
				'Next
			End Select
		Next
		CScript.PrintI "结束."': STD.Input
	End Select
	Model = ""
	Call Main
End Sub


'----

