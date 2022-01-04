
'---- init
Const Title = "Citra 辅助工具 v:1.0"
Const Color = "0A"
Const SelectGAME = "重装机兵4"

Set Core = CreateObject("Scripting.FileSystemObject").OpenTextFile(Replace(WSH.ScriptFullName, WSH.ScriptName, "") & "Core.GEOM", 1, False)
Execute Core.ReadAll

'----MM4
Path_Save	= WorkPath & ID_Save
Path_Patch	= WorkPath & ID_Patch
Path_DLC	= WorkPath & ID_DLC


Set SD_Menu  = CreateObject("Scripting.Dictionary")
Set SD_Subsidiary  = CreateObject("Scripting.Dictionary")



For Each Item In Array("1 开关补丁|Patch", "2 开关DLC|DLC", "3 选择金手指文件", "4 选择MOD")
	Dim TempVal
	TempVal = Split(Item)
	SD_Menu.Add TempVal(0), TempVal(1)
Next

SD_Subsidiary.Add "Patch", Path_Patch
SD_Subsidiary.Add "DLC", Path_DLC

'--------
Echo "	 "

Echo "[信息]"
Echo " 根路径: " & ScriptPath
Echo " 金手指: " & Path_Cheat
Echo " 模  组: " & Path_MODS
Echo " 游戏名: " & SelectGAME
Echo "   存档: " & IIf(File.EFo(Path_Save), "存在", "关闭")
Echo "   补丁: " & IIf(File.EFo(Path_Patch), "存在", "关闭")
Echo "    DLC: " & IIf(File.EFo(Path_DLC), "存在", "关闭")
Echo "----"
Echo "[命令]"
Echo  "a 打开存档文件夹"
Echo  "1 切换金手指"
Echo  "2 切换模组"
Echo  "3 开关补丁"
Echo  "4 开关DLC"

Set Obj_CL = WScript.Arguments
If Obj_CL.Count = 0 Then
	'
Else
	For Each CMS In Obj_CL
		Select Case CMS
		Case "3": Call SubsidiaryOC("Patch")
		Case "4": Call SubsidiaryOC("DLC")
		End Select
	Next
End If

Call Main()



'----
Sub Main()
	Dim SelectValue
	Echo ""
	STD.PrintI " 选择:": SelectValue = STD.Input
	Select Case SelectValue
	Case "": Call EndVBS()
	Case "1": Call Select_Cheat()
	Case "2": Call Select_MOD()
	Case "3": Call SubsidiaryOC("Patch")
	Case "4": Call SubsidiaryOC("DLC")
	Case "a": Call Runs("explorer " & PathC34(Path_Save & "\" & ID_Extend & "\data\00000001\"), True)
	Case Else
		Echo "表瞎搞~"
	End Select
	Call Main()
End Sub

Sub SubsidiaryOC(ByVal Item)
	Dim TempVal
	On Error Resume Next
	TempVal = SD_Subsidiary.Item(Item)
	If TempVal = "" Then
		Echo "未知参数: " & Item
	Else
		If File.EFo(TempVal) Then
			Call ReDir(TempVal)
		Else
			Call MKlink(TempVal, TempVal & " " & Item, False)
		End If
		Sleep 400
		Echo IIf(File.EFo(TempVal), "已开启", "已关闭")
	End If
	Error.ETD()
End Sub

Sub Select_Cheat()
	Dim I, S, Text_Cheat, String_Cheat, Array_Cheat, Select_Cheat
	Echo ""
	Echo "[切换金手指]"
	Text_Cheat = "0 关闭"
	On Error Resume Next
	For Each List in FSO.GetFolder(Path_Cheat & ID_KEY).Files
		I = I + 1
		Text_Cheat = Text_Cheat & vbCrLf & I & " " & List.Name
		'Echo "> " & List.Path
		String_Cheat = String_Cheat & vbCrLf & List.Name & " " & List.Path
	Next
	Array_Cheat = Split(String_Cheat, vbCrLf)
	
	S = InputBox(Text_Cheat,"选择字体", 1)
	CheatFS = Path_Cheat & ID_KEY & ".txt"
	Select Case S
	Case ""
	Case 0: ReFile CheatFS
	Case Else
		Select_Cheat = Split(Array_Cheat(S))
		If File.EFi(CheatFS) Then Call ReFile(CheatFS)
		Call MKlink(CheatFS, Select_Cheat(1), True)
		Echo "设置为:" & Select_Cheat(0)
	End Select
	Error.ETD()
	Echo "----"
End Sub

Sub Select_MOD()
	Dim I, S, Text_MODS, String_MODS, Array_MODS, Select_MODS
	Echo ""
	Echo "[切换模组字体]"
	Text_MODS = "0 关闭"
	On Error Resume Next
	For Each List in FSO.GetFolder(Path_MODS).SubFolders
		I = I + 1
		Select Case LCase(List.Name)
		Case "romsf"
		Case Else
			Text_MODS = Text_MODS & vbCrLf & I & " " & List.Name
			'Echo "> " & List.Path
			String_MODS = String_MODS & vbCrLf & List.Name & " " & List.Path
		End Select
	Next
	Array_MODS = Split(String_MODS, vbCrLf)
	
	S = InputBox(Text_MODS,"选择字体", 1)
	ROMFS = Path_MOD & ID_KEY & "\romfs"
	Select Case S
	Case ""
	Case 0: ReDir ROMFS
	Case Else
		Select_MODS = Split(Array_MODS(S))
		If File.EFo(ROMFS) Then Call ReDir(ROMFS)
		Call MKlink(ROMFS, Select_MODS(1), "")
		Echo "设置为:" & Select_MODS(0)
	End Select
	Error.ETD()
	Echo "----"
End Sub

