
Const Title = "����ģ�������� v:1.2"
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
Echo "	a	��������"
Echo "	b	MM4�浵���ݹ���"
Echo ""
Echo "	----	[ ѡ������ģ���� ]"
Echo "	0	���CPU�汾��ģ����"

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

SD.Add "Ra","Citra-��������.vbs"
SD.Add "Rb","Citra-MM4�浵����.vbs"

Call Main

Sub Main()
	CScript.PrintI "	ѡ��:": CM = CScript.Input
	Select Case True
	Case SD.Exists("G" & CM)
		If File.EFi(SD.Item("G" & CM) & "\citra-qt.exe") Then
			FilePath = PathC34(SD.Item("G" & CM) & "\citra-qt.exe")
			Shell FilePath, "", SD.Item("G" & CM) & "\", "", ""
		Else
			Echo "�ļ�������"
		End If
	Case CM = "0"
		If File.EFi(ScriptPath & "\Citra_CPU\Citra-����ģ����-CPU�Ż�.vbs") Then
			FilePath = PathC34(ScriptPath & "\Citra_CPU\Citra-����ģ����-CPU�Ż�.vbs")
			Shell FilePath, "", ScriptPath & "\Citra_CPU\", "", ""
		Else
			CScript.PrintL "�ļ�������"
		End If
	Case SD.Exists("R" & CM)
		Shell SD.Item("R" & CM), "", ScriptPath, "", ""
	End Select
	Call Main
End Sub


