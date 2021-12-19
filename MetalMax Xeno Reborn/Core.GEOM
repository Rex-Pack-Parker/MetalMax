
Set MGMT = GetObject("winmgmts:")
Set WMI = GetObject("Winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set WSS = CreateObject("WScript.Shell")
Set SA  = CreateObject("Shell.Application")
Set FSO = CreateObject("Scripting.FileSystemObject")
Set SD  = CreateObject("Scripting.Dictionary")

With WSS
	Set SES_P = .Environment("Process")
	Set SES_S = .Environment("System")
	Set SES_U = .Environment("User")
End With

Const CL_UAC     = "-UAC"
Const CL_UACR    = "-UACR"
Const CL_CMDR	 = "-CMDR"
Const CL_MGMT    = "-MGMT"
Const CL_About   = "-About"
Const CL_InfoTip = "-InfoTip"
Const CL_Host32  = "-Host:32"
Const CL_Host64  = "-Host:64"

Const HostWS = "WSCRIPT.EXE"
Const HostCS = "CSCRIPT.EXE"

Const vbSplit = 8204
Const CntStr  = "0123456789 ABCDEFGHIJKLMNOPQRSTUVWXYZ abcdefghijklmnopqrstuvwxyz `~!@#$%^&*()[]\;',.{}:""<>? +-*/"

Set STD			= New BaseSTD
Set File		= New BaseFile
Set Host		= New BaseHost
Set Error		= New BaseError
Set CScript		= New BaseCScript
Set Command		= New BaseCommand
Set Windows		= New BaseWindows

CScript.Color = Color
CScript True, ""


Dim ScriptPath, RootPath, WorkPath
Dim Path_Cheat, Path_MODF, Path_MODS
Dim GetGAMEIDs, ID_KEY, ID_Save, ID_Patch, ID_DLC

ScriptPath = Replace(WSH.ScriptFullName, WSH.ScriptName, "")

'---- Path
RootPath 	= ScriptPath & "Data\_User\"
WorkPath 	= RootPath & "sdmc\Nintendo 3DS\00000000000000000000000000000000\00000000000000000000000000000000\title\"

Path_Cheat	= RootPath & "Cheats\"
Path_MOD	= RootPath & "load\mods\"
Path_MODS	= Path_MOD & "bcfnt\"

'---- Body
Echo "------------------------"
Echo "	" & Title
Echo "------------------------"
Echo "  关于"
Echo "   基于 Core.GEOM"
Echo "   by: 雷克斯.派(Rex.Pack)"
Echo "	站点"
Echo "	 http://window.es-geom.com "
Echo "	 http://metalmax.fans "
Echo "------------------------"



'----收录游戏
Set SD_GAME  = CreateObject("Scripting.Dictionary")
								'KEY, Save, Patch, DLC, Extend
SD_GAME.Add "重装机兵4", Array("00040000000AFD00", "00040000", "0004000e", "0004008c", "000afd00")

'----选择游戏
GetGAMEIDs = SD_GAME.Item(SelectGAME)
If VarType(GetGAMEIDs) = vbSplit Then
	ID_KEY		= GetGAMEIDs(0)
	ID_Save		= GetGAMEIDs(1)
	ID_Patch	= GetGAMEIDs(2)
	ID_DLC		= GetGAMEIDs(3)
	ID_Extend	= GetGAMEIDs(4)
End If


'--------
'----辅助

Sub EndVBS() '退出脚本
	WScript.Quit
End Sub

Function IIf(ByVal TF, ByVal T, ByVal F) '标准IIf函数
	If TF = True Then IIf = T Else IIf = F
End Function

Function Runs(ByVal GEOM_ProgramPath, ByVal ViewWindow)
	'On Error Resume Next
	Call WSS.Run(GEOM_ProgramPath, IIf(ViewWindow = True, 1, vbHide))
	Runs = Error.EIF(True, Err.Description)
End Function

Function Shell(ByVal ProgramPath, ByVal CM1, ByVal CM2, ByVal CM3, ByVal CM4)		'运行程序 ShellExecute方式
  CM4 = IIf(CM4 = "", 1, CM4)
  Shell = SA.ShellExecute(ProgramPath, CM1, CM2, CM3, IIf(IsNumeric(CM4), IIf(CM4 <= 10, CM4, 1), 1))
End Function

Function PathC34(ByVal FP) '判断路径中是否含有空格,没有则去除"号 有则添加上""
	If InStr(FP, " ") = 0 Then
		PathC34 = Replace(FP, Chr(34), "")
	Else
		PathC34 = IIf(Left(FP, 1) = Chr(34), "", Chr(34)) & FP & IIf(Right(FP, 1) = Chr(34), "", Chr(34))
	End If
End Function

Function PathCC(ByVal FP) '去除路径的""号,无论有没有空格
	PathCC = Replace(FP, Chr(34), "")
End Function

Function Echo(ByVal Text) '
	WSH.Echo "  " & Text
End Function

Sub Sleep(ByVal MS)
	WSH.Sleep MS
End Sub

Function URLDecode(ByVal GEOM_strURL) 'URL转码
	Dim AC_I
	If InStr(GEOM_strURL, "%") = 0 Then
		URLDecode = GEOM_strURL
		Exit Function
	End If
	For AC_I =  1 To Len(GEOM_strURL)
		If Mid(GEOM_strURL, AC_I, 1) = "%" Then
		If Eval("&H" & Mid(GEOM_strURL, AC_I + 1, 2)) > 127 Then
			URLDecode = URLDecode & Chr(Eval("&H" & Mid(GEOM_strURL, AC_I + 1, 2) & Mid(GEOM_strURL, AC_I + 4, 2)))
			AC_I = AC_I + 5
			Else
			URLDecode = URLDecode & Chr(Eval("&H" & Mid(GEOM_strURL, AC_I + 1, 2)))
			AC_I = AC_I + 2
			End If
			Else
		URLDecode = URLDecode & Mid(GEOM_strURL, AC_I, 1)	   
		End If	
	Next
End Function

'----
Sub MKlink(ByVal F_To, ByVal F_Path, ByVal FF)
	Dim TempVal
	TempVal = "/c mklink " & IIf(FF, "/h", "/j") & " " & PathC34(F_To) & " " & PathC34(F_Path)
	Echo TempVal
	Call Runs("CMD " & TempVal, False)
End Sub

Sub ReDir(ByVal F_Path)
	Dim TempVal
	TempVal = "/c rd /s /q " & PathC34(F_Path)
	Echo TempVal
	Call Runs("CMD " & TempVal, False)
End Sub

Sub ReFile(ByVal F_Path)
	Dim TempVal
	TempVal = "/c del " & PathC34(F_Path)
	Echo TempVal
	Call Runs("CMD " & TempVal, False)
End Sub

 '--------
'CScript宿主
Class BaseCScript
	Private Data_Color, Data_Mode_Lines, Data_Mode_Cols
	
	Private Sub Class_Initialize()
		Color = "F0"
		Mode_Lines = 48
		Mode_cols = 128
	End Sub
	Private Sub Class_Terminate()
		
	End Sub
	
	Property Let Color(ByVal GEOM_TempDataStr)
		Data_Color = GEOM_TempDataStr
	End Property
	Property Get Color()
		Color = Data_Color
	End Property
	
	Property Let Mode_Lines(ByVal GEOM_TempDataStr)
		Data_Mode_Lines = GEOM_TempDataStr
	End Property
	Property Get Mode_Lines()
		Mode_Lines = Data_Mode_Lines
	End Property
	
	Property Let Mode_Cols(ByVal GEOM_TempDataStr)
		Data_Mode_Cols = GEOM_TempDataStr
	End Property
	Property Get Mode_Cols()
		Mode_Cols = Data_Mode_Cols
	End Property
	
	Function Help()
		Help = "同CMD中的Color"
	End Function
	
	'改用CScript宿主模式重新启动, CSave = True Then 保留CMD, Bit 32/64
	Public Default Function Run(ByVal CSave, ByVal GEOM_Bit)
		'If Windows.Bit = 32 Then GEOM_Bit = 32
		If GEOM_Bit = "" Then GEOM_Bit = Host.Bit
		If Host.HostType = True And GEOM_Bit = Host.Bit Then
			Run = "相同模式."
		Else
			GEOM_TempDataStr = Windows.SysPath(GEOM_Bit)
			If GEOM_TempDataStr = False Then
				Run = GEOM_Bit & "位系统路径不正确."
			Else
				CM = Split("     ")
				CM(0) = "CMD /" & IIf(CSave = True, "k", "c")
				CM(1) = " mode con:" & " lines=" & Data_Mode_Lines & " cols=" & Data_Mode_Cols
				CM(2) = " & Color " & Data_Color & " & " & "Title " & Title & " & "
				CM(3) = GEOM_TempDataStr & "CScript.exe //nologo " & PathC34(WSH.ScriptFullName) & " "
				CM(4) = Eval("CL_Host" & GEOM_Bit) & " " & _
						Command.All
				
				Runs Join(CM), True
				'Run = True
				EndVBS
			End If
		End If
	End Function
	
	'输出函数
	Sub PrintI(ByVal Texts)
		WScript.StdOut.Write Texts
	End Sub
	
	'输出函数,覆盖行,空白长度
	Sub PrintC(ByVal Texts, ByVal LenNum)
		WScript.StdOut.Write Chr(13) & Texts & String(LenNum, " ")
	End Sub
	
	'输出函数,+换行
	Sub PrintL(ByVal Texts)
		WScript.StdOut.WriteLine(Texts)
	End Sub
	
	'输入函数
	Function InPut()
		InPut = Trim(WScript.StdIn.ReadLine)
	End Function
 End Class

Class BaseHost '宿主信息
	'返回当前宿主程序名
	Function Name()
		Name = UCase(Replace(WScript.FullName, WScript.Path & "\", ""))
	End Function
	
	'返回当前宿主位
	Function Bit()
		GEOM_TempDataStr = Right(WScript.Path, 2)
		Select Case GEOM_TempDataStr
			Case "32", "64": Bit = Array(False, 32, 64)(Windows.Bit \ CInt(GEOM_TempDataStr))
			Case Else: Bit = "非官方宿主."
		End Select
		GEOM_TempDataStr = Empty
	End Function
	
	Function HostType()
		Select Case Host.Name
			Case HostWS: HostType = False
			Case HostCS: HostType = True
			Case Else: HostType = ""
		End Select
	End Function
	
	'指定Host,类型与位,运行vbs
	Public Default Function Run(ByVal GEOM_HostType, ByVal GEOM_Bit, ByVal GEOM_VBSFile, ByVal GEOM_Command, ByVal If_GEOM, ByVal tfRetain)
		Select Case UCase(GEOM_HostType)
		Case "W", "WSCRIPT", "WSCRIPT.EXE", "C", "CSCRIPT", "CSCRIPT.EXE"
			GEOM_TempDataStr = Windows.SysPath(GEOM_Bit)
			If GEOM_TempDataStr = False Then
				Run = GEOM_Bit & "Bit,路径不存在."
			Else
				GEOM_TempDataStr = GEOM_TempDataStr & UCase(Left(GEOM_HostType, 1)) & "Script.exe"
				If File.EFi(GEOM_TempDataStr) Then
					If PathCC(ScriptFullName) = PathCC(GEOM_VBSFile) And Not tfHostBit = True Then
						Run = "防止重复运行自身."
					Else
						CM = Split("  ")
						CM(0) = PathC34(GEOM_TempDataStr)
						CM(1) = IIf(If_GEOM, PathC34(GEOM_FullName) & " " & Command.GEOM(""), "") _
									& " " & Eval("CL_Host" & GEOM_Bit) _
									& " " & PathC34(GEOM_VBSFile) _
									& " " & Command("") _
									& " " & GEOM_Command
						Run = Shell(CM(0), CM(1), "", "", "")
						GEOM_TempDataStr = Empty
						If Not tfRetain = True Then EndVBS
					End If
				Else
					Run = GEOM_TempDataStr & ":Host不存在."
				End If
			End If
		Case Else: Run = "非支持的Host."
		End Select
	End Function
 End Class

Class BaseFile '文件系统
	Private GEOM_FSO, GEOM_GF, GEOM_NowFN, GEOM_Tristate
	
	Private Sub Class_Initialize()
		Set GEOM_FSO = CreateObject("Scripting.FileSystemObject")
		GEOM_Tristate = 0
	End Sub
	Private Sub Class_Terminate()
		
	End Sub
	
	Public Function Obj()
		Set Obj = GEOM_FSO
	End Function
	
	Public Property Let OpenType(ByVal GEOM_TempDataStr)
		GEOM_Tristate = GEOM_TempDataStr
	End Property
	
	Public Default Property Get OpenType()
		OpenType = GEOM_Tristate
	End Property
	
	Function Add(ByVal GEOM_FP)				'创建文件
		On Error Resume Next
		GEOM_FSO.CreateTextFile(PathCC((GEOM_FP)))
		Add = Error.EIF(True, Err.Description)
	End Function
	
	Function ReadAll(ByVal GEOM_FP)			'读取文件所有内容
		On Error Resume Next
		Set GEOM_GF = GEOM_FSO.OpenTextFile(PathCC((GEOM_FP)), 1, False, OpenType)
		GEOM_TempData = GEOM_GF.ReadAll
		ReadAll = Error.EIF(GEOM_TempData, Err.Description)
		GEOM_TempData = Empty
	End Function
	
	Function ReadAllx(ByVal GEOM_FP)			'读取文件所有内容(脱离核心)
		On Error Resume Next
		Set GEOM_GF = GEOM_FSO.OpenTextFile(Replace(GEOM_FP, """", ""), 1, False, OpenType)
		GEOM_TempData = GEOM_GF.ReadAll
		If Err = 0 Then ReadAllx = GEOM_TempData Else ReadAllx = Err.Description: Err.Clear
		GEOM_TempData = Empty
	End Function
	
	Function ReadLine(ByVal GEOM_FP)			'读取文件一行内容
		On Error Resume Next
		If IIf((GEOM_FP) = GEOM_NowFN, True, False) Then
			GEOM_TempData = IIf(GEOM_GF.AtEndOfStream, True, Replace(GEOM_GF.ReadLine, vbLf, ""))
		Else
			Set GEOM_GF = GEOM_FSO.OpenTextFile(PathCC((GEOM_FP)), 1, False, OpenType)
			GEOM_TempData = GEOM_TempData & GEOM_GF.ReadLine & vbCrLf
			GEOM_NowFN = (GEOM_FP)
		End If
		ReadLine = Error.EIF(GEOM_TempData, Err.Description)
		GEOM_TempData = Empty
	End Function
	
	Function Write(ByVal GEOM_FP, ByVal GEOM_TempDataStr)	'写入文件
		'On Error Resume Next
		Set GEOM_GF = GEOM_FSO.OpenTextFile(PathCC((GEOM_FP)), 2, True, OpenType)
		GEOM_GF.Write GEOM_TempDataStr
		Write = Error.EIF(True, Err.Description)
		GEOM_TempDataStr = Empty
	End Function
	
	Function WriteLine(ByVal GEOM_FP, ByVal GEOM_TempDataStr)	'写入文件尾部
		'On Error Resume Next
		If IIf((GEOM_FP) = GEOM_NowFN, True, False) Then
			GEOM_GF.WriteLine GEOM_TempDataStr & vbCrLf
		Else
			Set GEOM_GF = GEOM_FSO.OpenTextFile(PathCC((GEOM_FP)), 8, True, OpenType)
			GEOM_GF.WriteLine GEOM_TempDataStr & vbCrLf
			GEOM_NowFN = (GEOM_FP)
		End If
		WriteLine = Error.EIF(True, Err.Description)
		GEOM_TempDataStr = Empty
	End Function
	
	Sub Skip(ByVal GEOM_Num)
		GEOM_GF.Skip GEOM_Num
	End Sub
	
	Sub SkipLine()
		GEOM_GF.SkipLine
	End Sub
	
	Sub Close()
		GEOM_NowFN = Empty
		GEOM_GF.Close
	End Sub
	
	Function ReName(ByVal GEOM_FP, ByVal GEOM_NN)	'文件重命名
		On Error Resume Next
		Set GEOM_GF = GEOM_FSO.GetFile((GEOM_FP))
		GEOM_GF.Name = GEOM_NN
		ReName = Error.EIF(True, Err.Description)
	End Function
	
	Function Delete(ByVal GEOM_FP)				'删除指定文件
		On Error Resume Next
		GEOM_FSO.DeleteFile (GEOM_FP)
		Delete = Error.EIF(True, Err.Description)
	End Function
	
	Function CopyFile(ByVal GEOM_FPA, ByVal GEOM_FPB, ByVal GEOM_TF)		'复制文件
		On Error Resume Next
		GEOM_FSO.CopyFile (GEOM_FPA), _
						(GEOM_FPB), _
						IIf(GEOM_TF = False, False, True)
		Copy = Error.EIF(True, Err.Description)
	End Function
	
	Function CopyFolder(ByVal GEOM_FPA, ByVal GEOM_FPB, ByVal GEOM_TF)		'复制文件夹
		On Error Resume Next
		GEOM_FSO.CopyFolder (GEOM_FPA), _
							(GEOM_FPB), _
							IIf(GEOM_TF = False, False, True)
		Copy = Error.EIF(True, Err.Description)
	End Function
	
	Function Move(ByVal GEOM_FPA, ByVal GEOM_FPB)		'移动文件
		On Error Resume Next
		GEOM_FSO.MoveFile (GEOM_FPA), (GEOM_FPB)
		Move = Error.EIF(True, Err.Description)
	End Function
	
	Function EFi(ByVal FilePath)
		EFi = GEOM_FSO.FileExists(PathCC(FilePath))
	End Function
	
	Function EFo(ByVal FolderPath)
		EFo = GEOM_FSO.FolderExists(PathCC(FolderPath))
	End Function
 End Class

Class BaseError '错误辅助
	'Private GEOM_tfClear
	
	Private Sub Class_Initialize()
	'	GEOM_tfClear = False
	End Sub
	Private Sub Class_Terminate()
		
	End Sub
	
	'Property Let tfEndClear(ByVal GEOM_tf)
	'	GEOM_tfClear = GEOM_tf
	'End Property
	
	'Property Get tfEndClear()
	'	tfEndClear = GEOM_tfClear
	'End Property
	
	'Private Sub Clear(ByVal tfEndClear)
	'	If tfEndClear Then Err.Clear
	'End Sub
'----以下需要先设置 On Error Resume Next----
	'除错版IIF,有无错误对应返回
	Function EIF(ByVal GEOM_T, ByVal GEOM_F)
		EIF = IIf(Err.Number = 0, GEOM_T, GEOM_F): Err.Clear
	End Function

	'当发现错误时,执行指定代码
	Sub ERC(ByVal GEOM_StrCode)
		IF Err.Number <> 0 Then ExeCute GEOM_StrCode: Err.Clear
	End Sub
	
	Sub Clear()
		Err.Clear
	End Sub
	
	'当发现错误时,提示并返回自定消息
	Function ETD(ByVal GEOM_TempData)
		If Err.Number = 0 Then
			ETD = True
		Else
			Err.Clear
			Echo GEOM_TempData
			ETD = GEOM_TempData
		End If
	End Function
	
	'当发现错误时,提示并返回错误信息
	Function ETM()
		If Err.Number = 0 Then
			ETM = True
		Else
			GEOM_TempDataStr = vbCrLf & "----:Error Msg" & _
								vbCrLf & "Source:" & Err.Source & _
								vbCrLf & "Description:" & Err.Description & _
								vbCrLf & "Number:" & Err.Number & _
								vbCrLf & "----:Error End"
			Err.Clear
			Echo GEOM_TempDataStr
			ETM = GEOM_TempDataStr
		End If
		GEOM_TempDataStr = Empty
	End Function

	'判断是否有错误
	Public Default Function Judge()
		Judge = EIF(True, False)
	End Function
 End Class


'Windows相关信息
Class BaseWindows 'Windows
	Private RootPath, SysPathArr, tfUAC
	
	Private Sub Class_Initialize()
		RootPath = SES_P.Item("WinDir")
		SysPathArr = Array(False, RootPath & "\System32\", RootPath & "\Syswow64\")
	End Sub
	
	Private Sub Class_Terminate()
		
	End Sub
	'----
	Function OS()
		For Each OSC In GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem")
			Select Case LCase(Split(OSC.Caption)(2))
			Case "95", "98", "2000", "xp", "2003": tfUAC = False
			Case "7", "8", "8.1", "10", "2008", "2012": tfUAC = True
			End Select
			OS = OSC.Caption
		Next
	End Function
	
	Function UAC()
		Call OS
		UAC = tfUAC
	End Function
	
	Function ComputerName()
		ComputerName = SES_P.Item("ComputerName")
	End Function
	
	Function Bit()
		If File.EFo(SysPathArr(2)) Then Bit = 64 Else Bit = 32
	End Function
	
	Function WinPath()
		WinPath = RootPath
	End Function
	
	Function SysPath(ByVal GEOM_Bit)
		Select Case CStr(GEOM_Bit)
		Case "32", "64": SysPath = SysPathArr(Bit \ GEOM_Bit)
		Case Else	: SysPath = True
		End Select
	End Function
 End Class

Class BaseCommand '运行参数
	Private GEOM_SD, GEOM_CL, VBS_CL, Obj_CL
	
	Private Sub Class_Initialize()
		Set GEOM_SD = CreateObject("Scripting.Dictionary")
			GEOM_SD.CompareMode = vbTextCompare
		With GEOM_SD
			For Each GEOM_TempDataStr In Array(CL_UAC, CL_UACR, CL_CMDR, CL_Host32, CL_Host64, CL_Code, CL_About, CL_InfoTip, CL_MGMT)
				.Add GEOM_TempDataStr, False
			Next: .Add "VBS", ""
			
			Set Obj_CL = WScript.Arguments
			For Each CMS In Obj_CL
				If .Exists(CMS) Then
					.Item(CMS) = True
					GEOM_CL = GEOM_CL & IIf(File.EFi(CMS), PathC34(CMS), CMS) & vbCrLf
				Else	'收集非GEOM识别的参数
					If CMS = GEOM_FullName Then
						.Add "GEOM", True
					Else
						If .Item("VBS") = "" And File.EFi(CMS) Then .Item("VBS") = PathC34(CMS) Else VBS_CL = VBS_CL & IIf(File.EFi(CMS), PathC34(CMS), CMS) & vbCrLf
					End If
					
				End If
			Next
		End With
	End Sub
	
	Private Sub Class_Terminate()
		Set GEOM_SD = NoThing
	End Sub
	
	'----
	Public Function Enabled()
		Enabled = IIf(Obj_CL.Count = 0, False, True)
	End Function
	
	Public Function Item(ByVal GEOM_CN)
		If GEOM_SD.Exists(GEOM_CN) Then Item = GEOM_SD.Item(GEOM_CN) Else Item = False
	End Function
	
	Public Function All()
		All = GEOM("") & " " & VBS("")
	End Function
	
	Public Function GEOM(ByVal GEOM_Type)
		GEOM = LineType(GEOM_CL, GEOM_Type)
	End Function
	
	Public Default Function VBS(ByVal GEOM_Type)	
		VBS = LineType(VBS_CL, GEOM_Type)
	End Function
	
	Private Function LineType(ByVal GEOM_Line, ByVal GEOM_Type)	
		Select Case LCase(GEOM_Type)
		Case "array": LineType = Split(GEOM_Line, vbCrLf)
		Case Else: LineType = Replace(GEOM_Line, vbCrLf, " ")
		End Select
	End Function
 End Class

Class BaseSTD 'CScript I\O
	Sub PrintI(ByVal Texts)
		WScript.StdOut.Write Texts
	End Sub
	
	Sub PrintC(ByVal Texts, ByVal LenNum)
		WScript.StdOut.Write Chr(13) & Texts & String(LenNum, " ")
	End Sub
	
	Sub PrintL(ByVal Texts)
		WScript.StdOut.WriteLine(Texts)
	End Sub
	
	Function InPut()
		InPut = Trim(WScript.StdIn.ReadLine)
	End Function
 End Class
