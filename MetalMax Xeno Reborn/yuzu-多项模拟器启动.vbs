
Const Title = "����ģ�������� v:1.0"
Const Color = "03"

Set Core = CreateObject("Scripting.FileSystemObject").OpenTextFile(Replace(WSH.ScriptFullName, WSH.ScriptName, "") & "Core.GEOM", 1, False)
Execute Core.ReadAll


'---- main

Echo "------------------------"
Echo "	"
Echo "	----	[ ·�� ]"
Echo "	" & ScriptPath
Echo ""
Echo "	----	[ ���� ]"
Echo "	a	��������(δ����)"
Echo "	b	MMXR�浵���ݹ���"
Echo ""
Echo "	----	[ ѡ������ģ���� ]"

For Each List in FSO.GetFolder(ScriptPath).subfolders
	Select Case List.Name
	Case "����"
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
SD.Add "Rb","yuzu-MMXR�浵����.vbs"

Call Main

Sub Main()
	CScript.PrintI "	ѡ��:": CM = CScript.Input
	Select Case True
	Case SD.Exists("G" & CM)
		If File.EFi(SD.Item("G" & CM) & "\yuzu.exe") Then
			FilePath = PathC34(SD.Item("G" & CM) & "\yuzu.exe")
			Shell FilePath, "", SD.Item("G" & CM) & "\", "", ""
		Else
			Echo "�ļ�������"
		End If
	Case SD.Exists("R" & CM)
		Shell SD.Item("R" & CM), "", ScriptPath, "", ""
	End Select
	Call Main
End Sub


