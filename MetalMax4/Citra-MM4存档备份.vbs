
Const Title = "Citra MetalMax4 �浵���� v:1.0 �����(�������ϰ���)"
Const Color = "0D"
Const SelectGAME = "��װ����4"

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
	Echo "�޷����ӷ��������ݴ浵."
	Call EndVBS
End If

'�����������Ƿ��д�MAC�ļ���
For Each Temp In NowPathItem(1).Items
	If Temp.Name = MAC Then tfMAC = True: Exit For
Next

'���û�д�MAC�ļ���,���½�
If TypeName(NowPathItem(0).ParseName(MAC))= "Nothing" Then NowPathItem(0).NewFolder MAC
WSH.Sleep 500
If Not tfMAC Then
	NowPathItem(1).CopyHere NowPathItem(0).Self.Path & "\" & MAC & "\"
	WSH.Sleep 1000
End If
Set NowPathItem(1) = FSA(1).NameSpace(RootPathItem).ParseName(MAC).GetFolder

'��ȡ2���浵
For Each Temp In NowPathItem(0).Items
	Select Case True
	Case LCase(Temp.Name) = Slot(0):SlotPath(0) = PathCC(Temp.Path)
	Case LCase(Temp.Name) = Slot(1):SlotPath(1) = PathCC(Temp.Path)
	End Select
Next

'----

Echo "------------------------"
Echo "	���ID:  " & MAC
Echo "	�����ǰ�� http://metalmax.fans �а��� ���ID �ҵ���ǰ���ݹ��Ĵ浵"
Echo "------------------------"
Echo "	[ 1 ] Slot_0	" & IIf(SlotPath(0)="","��","����")
Echo "	[ 2 ] Slot_1	" & IIf(SlotPath(1)="","��","����")
Echo "------------------------"

Call Main


Sub Main()
	On Error Resume Next
	STD.PrintI "	�����ĸ�:":Slot_Index = STD.Input
	
	Select Case Slot_Index
	Case ""
	Case "1", "2"
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
