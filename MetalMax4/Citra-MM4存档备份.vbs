
Const Title = "Citra MetalMax4 存档备份 v:1.0 公测版(仅限整合版用)"
Const Color = "0D"
Const SelectGAME = "重装机兵4"

Set Core = CreateObject("Scripting.FileSystemObject").OpenTextFile(Replace(WSH.ScriptFullName, WSH.ScriptName, "") & "Core.GEOM", 1, False)
Execute Core.ReadAll

'---- init
Dim FSA(1)
Dim MAC, tfMAC, Slot(1), SlotPath(1)
Dim NowSlot, Slot_Index, NowPathItem(1), NowFileItem(1)

Set FSA(0)  = CreateObject("Shell.Application")
Set FSA(1)  = CreateObject("Shell.Application")
Set RootPathItem = FSA(0).NameSpace(URLDecode("ftp://mms@s.es-geom.com/"))

Slot(0) = "slot_0"
Slot(1) = "slot_1"

tfMAC = False

Set NowPathItem(0) = FSA(0).NameSpace(ScriptPath & "Data\_User\sdmc\Nintendo 3DS\00000000000000000000000000000000\00000000000000000000000000000000\title\00040000\000afd00\data\00000001\")
Set NowPathItem(1) = FSA(1).NameSpace(RootPathItem)

'On Error Resume Next
For Each GEOMNetWorkAC In MGMT.InstancesOf("Win32_NetworkAdapterConfiguration Where IPEnabled = TRUE") ' Where IPEnabled = TRUE
	MAC = Replace(GEOMNetWorkAC.MacAddress,":","")
	Err.Clear
	If Not MAC = "" Then Exit For
Next

'---- code

If MAC = "" Then
	Echo "无法链接服务器备份存档."
	Call EndVBS
End If

'检索服务器是否有此MAC文件夹
For Each Temp In NowPathItem(1).Items
	If Temp.Name = MAC Then tfMAC = True: Exit For
Next

'如果没有此MAC文件夹,则新建
If TypeName(NowPathItem(0).ParseName(MAC))= "Nothing" Then NowPathItem(0).NewFolder MAC
WSH.Sleep 500
If Not tfMAC Then
	NowPathItem(1).CopyHere NowPathItem(0).Self.Path & "\" & MAC & "\"
	WSH.Sleep 1000
End If
Set NowPathItem(1) = FSA(1).NameSpace(RootPathItem).ParseName(MAC).GetFolder

'获取2个存档
For Each Temp In NowPathItem(0).Items
	Select Case True
	Case LCase(Temp.Name) = Slot(0):SlotPath(0) = PathCC(Temp.Path)
	Case LCase(Temp.Name) = Slot(1):SlotPath(1) = PathCC(Temp.Path)
	End Select
Next

'----

Echo "------------------------"
Echo "	你的ID:  " & MAC
Echo "	你可以前往 http://metalmax.fans 中按照 你的ID 找到先前备份过的存档"
Echo "------------------------"
Echo "	[ 1 ] Slot_0	" & IIf(SlotPath(0)="","无","存在")
Echo "	[ 2 ] Slot_1	" & IIf(SlotPath(1)="","无","存在")
Echo "------------------------"

Call Main


Sub Main()
	On Error Resume Next
	STD.PrintI "	备份哪个:":Slot_Index = STD.Input
	
	Select Case Slot_Index
	Case ""
	Case "1", "2"
		If SlotPath(Slot_Index - 1) = "" Then
			Echo "  该存档不存在,心里没点数嘛~"
		Else
			NowPathItem(1).CopyHere SlotPath(Slot_Index - 1)
			NowSlotName = Slot(Slot_Index - 1)
			NewSlotName = NowSlotName & " " & Replace(Date, "/","-") & "_" & Replace(Time,":","-")
			Set NowSlot =  NowPathItem(1).ParseName(NowSlotName)
			WSH.Sleep 1000
			NowSlot.Name = NewSlotName
		End If
	Case Else
		Echo "  表瞎搞~"
	End Select
	
	Echo Err.Description
	Err.Clear
	Call Main
End Sub

'----
