
Const Title = "多项模拟器启动 v:1.2"
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
Echo "	a	辅助工具"
Echo "	b	MM4存档备份工具"
Echo ""
Echo "	----	[ 选择启动模拟器 ]"
Echo "	0	针对CPU版本的模拟器"

For Each List in FSO.GetFolder(ScriptPath).subfolders
	Select Case List.Name
	Case "Citra_CPU"
	Case "Data"
	Case Else
		If File.EFi(List & "\citra-qt.exe") Then
			i = i + 1
			Echo "	" & i & "	" & List.Name
			SD.Add "G" & i, List
		End If
	End Select
Next
Echo ""

SD.Add "Ra","Citra-辅助工具.vbs"
SD.Add "Rb","Citra-MM4存档备份.vbs"

Call Main

Sub Main()
	CScript.PrintI "	选择:": CM = CScript.Input
	Select Case True
	Case SD.Exists("G" & CM)
		If File.EFi(SD.Item("G" & CM) & "\citra-qt.exe") Then
			FilePath = PathC34(SD.Item("G" & CM) & "\citra-qt.exe")
			Shell FilePath, "", SD.Item("G" & CM) & "\", "", ""
		Else
			Echo "文件不存在"
		End If
	Case CM = "0"
		If File.EFi(ScriptPath & "\Citra_CPU\Citra-启动模拟器-CPU优化.vbs") Then
			FilePath = PathC34(ScriptPath & "\Citra_CPU\Citra-启动模拟器-CPU优化.vbs")
			Shell FilePath, "", ScriptPath & "\Citra_CPU\", "", ""
		Else
			CScript.PrintL "文件不存在"
		End If
	Case SD.Exists("R" & CM)
		Shell SD.Item("R" & CM), "", ScriptPath, "", ""
	End Select
	Call Main
End Sub


