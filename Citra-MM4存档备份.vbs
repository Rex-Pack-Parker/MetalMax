
Set WSS = CreateObject("WScript.Shell")

Const GEOM_Title = "Citra MetalMax4 �浵����    by:Rex.Pack v:1.0 �����(�������ϰ���)"
Const GEOM_Color = "0D"

If Not LCase(Replace(WScript.FullName, WScript.Path & "\", "")) = "cscript.exe" Then
	WSS.Run "cmd /c mode con: lines=70 cols=100 & color " & GEOM_Color & " & Title " & GEOM_Title & " & CScript //nologo " & Chr(34) & WScript.ScriptFullName & Chr(34)
	WSH.Sleep 1000
	WSH.Quit
End If


'---- init
Dim SA(1), MAC, tfMAC, Slot(1), SlotPath(1), NowSlot, Slot_Index, NowPathItem(1), NowFileItem(1)

Set STD = New BaseSTD
Set GEOMMGMT = GetObject("winmgmts:")
Set SA(0)  = CreateObject("Shell.Application")
Set SA(1)  = CreateObject("Shell.Application")
Set RootPathItem = SA(0).NameSpace(URLDecode("ftp://mms@s.es-geom.com/"))

ScriptPath = Replace(WSH.ScriptFullName, WSH.ScriptName, "")

Slot(0) = "slot_0"
Slot(1) = "slot_1"

tfMAC = False

Set NowPathItem(0) = SA(0).NameSpace(ScriptPath & "Data\_User\sdmc\Nintendo 3DS\00000000000000000000000000000000\00000000000000000000000000000000\title\00040000\000afd00\data\00000001\")
Set NowPathItem(1) = SA(1).NameSpace(RootPathItem)

On Error Resume Next
For Each GEOMNetWorkAC In GEOMMGMT.InstancesOf("Win32_NetworkAdapterConfiguration") ' Where IPEnabled = TRUE
	MAC = Replace(GEOMNetWorkAC.MacAddress,":","")
	Err.Clear
	If Not MAC = "" Then Exit For
Next

'---- code

'�����������Ƿ��д�MAC�ļ���
For Each Temp In NowPathItem(1).Items
	If Temp.Name = MAC Then tfMAC = True: Exit For
Next

'���û�д�MAC�ļ���,���½�
If TypeName(NowPathItem(0).ParseName(MAC))= "Nothing" Then NowPathItem(0).NewFolder MAC
If Not tfMAC Then
	WSH.Sleep 500
	NowPathItem(1).CopyHere NowPathItem(0).Self.Path & "\" & MAC & "\"
End If
Set NowPathItem(1) = SA(1).NameSpace(RootPathItem).ParseName(MAC).GetFolder

'��ȡ2���浵
For Each Temp In NowPathItem(0).Items
	Select Case True
	Case LCase(Temp.Name) = Slot(0):SlotPath(0) = PathCC(Temp.Path)
	Case LCase(Temp.Name) = Slot(1):SlotPath(1) = PathCC(Temp.Path)
	End Select
Next

'----

STD.PrintL "------------------------"
STD.PrintL "	" & GEOM_Title
STD.PrintL "------------------------"
STD.PrintL "	���ID:  " & MAC
STD.PrintL "	�����ǰ�� http://metalmax.fans �а��� ���ID �ҵ���ǰ���ݹ��Ĵ浵"
STD.PrintL "------------------------"
STD.PrintL "	[ 1 ] Slot_0	" & IIf(SlotPath(0)="","��","����")
STD.PrintL "	[ 2 ] Slot_1	" & IIf(SlotPath(1)="","��","����")
STD.PrintL "------------------------"

Call Main


Sub Main()
	On Error Resume Next
	STD.PrintI "	�����ĸ�:":Slot_Index = STD.Input
	
	Select Case Slot_Index
	Case 1, 2
		If SlotPath(Slot_Index - 1) = "" Then
			Echo "  �ô浵������,����û������~"
		Else
			NowPathItem(1).CopyHere SlotPath(Slot_Index - 1)
			NowSlotName = Slot(Slot_Index - 1)
			NewSlotName = NowSlotName & " " & Replace(Date, "/","-") & "_" & Replace(Time,":","-")
			Set NowSlot =  NowPathItem(1).ParseName(NowSlotName)
			WSH.Sleep 1000
			NowSlot.Name = NewSlotName
		End If
	Case Else
		Echo "  ��Ϲ��~"
	End Select
	
	Echo Err.Description
	Err.Clear
	Call Main
End Sub

'----
'CScript�������
Class BaseSTD
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

'----����
'��׼IIf����
Function IIf(ByVal TF, ByVal T, ByVal F)
	If TF = True Then IIf = T Else IIf = F
End Function

'�ж�·�����Ƿ��пո�,û����ȥ��"�� ���������""
Function PathC34(ByVal GEOM_FP)
	If InStr(GEOM_FP, " ") = 0 Then
		PathC34 = Replace(GEOM_FP, Chr(34), "")
	Else
		PathC34 = IIf(Left(GEOM_FP, 1) = Chr(34), "", Chr(34)) & GEOM_FP & IIf(Right(GEOM_FP, 1) = Chr(34), "", Chr(34))
	End If
End Function

 'ȥ��·����""��,������û�пո�
Function PathCC(ByVal GEOM_FP)
	PathCC = Replace(GEOM_FP, Chr(34), "")
End Function

Function Echo(ByVal Text)
	WSH.Echo "  " & Text
End Function

Private Function URLDecode(ByVal GEOM_strURL)
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