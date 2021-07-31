
'--����
Set GEOM_WSS = CreateObject("WScript.Shell")
Set GEOM_SA  = CreateObject("Shell.Application")
Set GEOM_FSO = CreateObject("Scripting.FileSystemObject")
Set GEOM_SD  = CreateObject("Scripting.Dictionary")

Const GEOM_Title = "����ģ��������    by:Rex.Pack v:1.2"
Const GEOM_Color = "03"

If Not LCase(Replace(WScript.FullName, WScript.Path & "\", "")) = "cscript.exe" Then
	GEOM_WSS.Run "cmd /k mode con: lines=40 cols=90 & color " & GEOM_Color & " & Title " & GEOM_Title & " & CScript //nologo " & Chr(34) & WScript.ScriptFullName & Chr(34) & Command
	WSH.Sleep 1000
	WSH.Quit
End If

NCD = Replace(WSH.ScriptFullName, WSH.ScriptName, "")
Echo "	"
Echo "	----	[ ·�� ]"
Echo "	" & NCD
Echo ""
Echo "	----	[ ���� ]"
Echo "	a	��������"
Echo "	b	MM4�浵���ݹ���"
Echo ""
Echo "	----	[ ѡ������ģ���� ]"
Echo "	0	���CPU�汾��ģ����"

For Each List in GEOM_FSO.GetFolder(NCD).subfolders
	Select Case List.Name
	Case "Citra_CPU"
	Case "Data"
	Case Else
		If tfFile(List & "\citra-qt.exe") Then
			i = i + 1
			Echo "	" & i & "	" & List.Name
			GEOM_SD.Add "G" & i, List
		End If
	End Select
Next
Echo ""

GEOM_SD.Add "Ra","Citra-��������.bat"
GEOM_SD.Add "Rb","Citra-MM4�浵����.vbs"

Call Main

Sub Main()
	Print "	ѡ��:": CM = Input
	Select Case True
	Case GEOM_SD.Exists("G" & CM)
	'If GEOM_SD.Exists("G" & CM) Then
		If tfFile(GEOM_SD.Item("G" & CM) & "\citra-qt.exe") Then
			FilePath = PathC34(GEOM_SD.Item("G" & CM) & "\citra-qt.exe")
			Shell FilePath, "", GEOM_SD.Item("G" & CM) & "\", "", ""
		Else
			PrintL "�ļ�������"
		End If
	'End If
	Case CM = "0"
		If tfFile(NCD & "\Citra_CPU\Citra-����ģ����-CPU�Ż�.vbs") Then
			FilePath = PathC34(NCD & "\Citra_CPU\Citra-����ģ����-CPU�Ż�.vbs")
			Shell FilePath, "", NCD & "\Citra_CPU\", "", ""
		Else
			PrintL "�ļ�������"
		End If
	Case GEOM_SD.Exists("R" & CM)
		Shell GEOM_SD.Item("R" & CM), "", NCD, "", ""
	End Select
	Call Main
End Sub


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
	GEOM_WSS.Run "cmd /c " & GEOM_ProgramPath, 0, True
	If Err Then Runs = Err.Description Else Runs = "���"
End Function

Function Shell(ByVal GEOM_ProgramPath, ByVal CM1, ByVal CM2, ByVal CM3, ByVal CM4)		'���г��� ShellExecute��ʽ
  CM4 = IIf(CM4 = "", 1, CM4)
  Shell = GEOM_SA.ShellExecute(GEOM_ProgramPath, CM1, CM2, CM3, IIf(IsNumeric(CM4), IIf(CM4 <= 10, CM4, 1), 1))
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

Sub PrintC(ByVal Texts)
	WScript.StdOut.Write Chr(13) & Texts
End Sub

Sub PrintL(ByVal Texts)
	WScript.StdOut.WriteLine(Texts)
End Sub

Function InPut()
	InPut = WScript.StdIn.ReadLine
End Function

Function tfFile(ByVal GEOM_FilePath)		'�ļ��Ƿ����
	tfFile = IIf(GEOM_FSO.FileExists(PathCC(GEOM_FilePath)) = False, False, True)
 End Function
 
Function tfFolder(ByVal GEOM_FilePath)		'�ļ����Ƿ����
	tfFolder = IIf(GEOM_FSO.FolderExists(PathCC(GEOM_FilePath)) = False, False, True)
 End Function