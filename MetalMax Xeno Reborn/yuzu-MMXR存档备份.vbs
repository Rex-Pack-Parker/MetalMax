
Const Title = "Citra MetalMax Xeno Reborn �浵���� v:1.0 �����(�������ϰ���)"
Const Color = "0D"
Const SelectGAME = "MetalMax Xeno Reborn"

Set Core = CreateObject("Scripting.FileSystemObject").OpenTextFile(Replace(WSH.ScriptFullName, WSH.ScriptName, "") & "Core.GEOM", 1, False)
Execute Core.ReadAll

'---- init
Dim FSA(1)
Dim MAC, tfMAC, SaveDataPath, SelectSDP
Dim NowPathItem(1), NowFileItem(1)

Set FSA(0)  = CreateObject("Shell.Application")
Set FSA(1)  = CreateObject("Shell.Application")
Set RootPathItem = FSA(0).NameSpace(URLDecode("ftp://mms@s.es-geom.com/"))


tfMAC = False

Set NowPathItem(0) = FSA(0).NameSpace(ScriptPath & "Data\_User\nand\user\save\0000000000000000\92E3DE845B188AD79504F415E0AB7445\0100900012F22000\")
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


'----

Echo "------------------------"
Echo "	���ID:  " & MAC
Echo "	�����ǰ�� http://metalmax.fans �а��� ���ID �ҵ���ǰ���ݹ��Ĵ浵"
Echo "------------------------"

'��ȡ�浵
SelectSDP = 0
For Each Temp In NowPathItem(0).Items
	If Left(Temp.Name, 12) = "MainSaveData" Then
		SelectSDP = SelectSDP + 1
		SD.Add Temp.Name, Temp.Path
		SaveDataPath = SaveDataPath & vbCrLf & Temp.Name
		Set ATTRIB = FSO.GetFile(Temp.Path)
		Echo  SelectSDP & "  " & Temp.Name & "	" & ATTRIB.DateLastModified
	End If
Next
Set ATTRIB = Nothing
SaveDataPath = Split(SaveDataPath, vbCrLf)
Echo "------------------------"

Call Main


Sub Main()
	On Error Resume Next
	STD.PrintI "	�����ĸ�:":SelectSDP = STD.Input
	
	Select Case SelectSDP
	Case ""
	Case "0"
		Echo "dsa"
	Case Else
		NowPathItem(1).CopyHere SD.Item(SaveDataPath(SelectSDP))
		NowSlotName = SaveDataPath(SelectSDP)
		NewSlotName = "MMXR - " & NowSlotName & " " & Replace(Date, "/","-") & "_" & Replace(Time,":","-")
		Set NowSlot =  NowPathItem(1).ParseName(NowSlotName)
		WSH.Sleep 1000
		NowSlot.Name = NewSlotName
	End Select
	
	Echo Err.Description
	Err.Clear
	Call Main
End Sub

'----
