
Const Title = "多项模拟器启动 v:1.0"
Const Color = "03"

Set Core = CreateObject("Scripting.FileSystemObject").OpenTextFile(Replace(WSH.ScriptFullName, WSH.ScriptName, "") & "Core.GEOM", 1, False)
Execute Core.ReadAll


'---- main

Echo "------------------------"
Echo "	"
Echo "	----	[ 路径 ]"
Echo "	" & ScriptPath
Echo ""
Echo "	----	[ 工具 ]"
Echo "	a	辅助工具(未开发)"
Echo "	b	MMXR存档备份工具"
Echo ""
Echo "	----	[ 选择启动模拟器 ]"

For Each List in FSO.GetFolder(ScriptPath).subfolders
	Select Case List.Name
	Case "数据"
	Case "Data"
	Case Else
		If File.EFi(List & "\yuzu.exe") Then
			i = i + 1
			Echo "	" & i & "	" & List.Name
			SD.Add "G" & i, List
		End If
	End Select
Next
Echo ""

SD.Add "Ra","*.vbs"
SD.Add "Rb","yuzu-MMXR存档备份.vbs"

Call Main

Sub Main()
	CScript.PrintI "	选择:": CM = CScript.Input
	Select Case True
	Case SD.Exists("G" & CM)
		If File.EFi(SD.Item("G" & CM) & "\yuzu.exe") Then
			FilePath = PathC34(SD.Item("G" & CM) & "\yuzu.exe")
			Shell FilePath, "", SD.Item("G" & CM) & "\", "", ""
		Else
			Echo "文件不存在"
		End If
	Case SD.Exists("R" & CM)
		Shell SD.Item("R" & CM), "", ScriptPath, "", ""
	End Select
	Call Main
End Sub


