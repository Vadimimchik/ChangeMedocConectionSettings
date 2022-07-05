Const ForReading = 1 
Const ForWriting = 2

Const doBackUp = False

Const strFolder1 = "c:\Program Files\Medoc\"
Const strFolder2 = "c:\ProgramData\Medoc\"
Const strFolder3 = "c:\Medoc\"

Const strSourceFile1 = "DMF.AppServer.exe.config"
Const strSourceFile2 = "ezvit.exe.config"
Const strSourceFile3 = "station.exe.config"

Const StrSource = "192.168.1.1"
Const StrDest = "192.168.2.1"
Dim StrDate
Dim objFSO
ReDim ArrayFiles(0)

Dim lenSourceFile1
Dim lenSourceFile2
Dim lenSourceFile3

Dim newSourceFile1
Dim newSourceFile2
Dim newSourceFile3

Function ReturnStrDate()
	DateYear = CStr(DatePart("yyyy", Date))
	DateMonth = DatePart("m", Date)
	If DateMonth < 10 Then
		DateMonth = "0" + CStr(DateMonth)
	Else	
		DateMonth = CStr(DateMonth)
	End If
	DateDay = DatePart("d", Date)
	If DateDay < 10 Then
		DateDay = "0" + CStr(DateDay)
	Else	
		DateDay = CStr(DateDay)
	End If
	
	ReturnStrDate = DateYear + DateMonth + DateDay
End Function

Function isOurFile(Val)
	isGood = False
	curFileName = Right(Val, lenSourceFile1 + 1)
	If curFileName = "\" + strSourceFile1 Then
		isGood = True
	Else
		curFileName = Right(Val, lenSourceFile2 + 1)
		If curFileName = "\" + strSourceFile2 Then
			isGood = True
		Else
			curFileName = Right(Val, lenSourceFile3 + 1)
			If curFileName = "\" + strSourceFile3 Then
				isGood = True
			End If
		End If
	End If

	isOurFile = isGood
End Function

Function AddInArray(Val)
	isGood = isOurFile(Val)
	If isGood Then
		i = UBound(ArrayFiles)
		ArrayFiles(i) = Val
		ReDim Preserve ArrayFiles(i + 1)
	End If
	
	AddInArray = isGood
End Function

Sub ChangeFile(SourcePath)
	If objFSO.FileExists(SourcePath) Then
		Set File = objFSO.OpenTextFile(SourcePath, ForReading)
		Buffer = File.ReadAll
		File.Close

		index = InStr(1, Buffer, StrSource, 1)
		If index <> 0 Then
			If doBackUp Then
				DestPath = Replace(SourcePath, strSourceFile1, newSourceFile1, 1, -1, 1)
				DestPath = Replace(DestPath, strSourceFile2, newSourceFile2, 1, -1, 1)
				DestPath = Replace(DestPath, strSourceFile3, newSourceFile3, 1, -1, 1)
				objFSO.CopyFile SourcePath, DestPath, True
			End If

			Buffer = Replace(Buffer, StrSource, StrDest, 1, -1, 1)
			Set File = objFSO.OpenTextFile(SourcePath, ForWriting)
			File.Write Buffer
			File.Close
		End If
		Set File = Nothing
	End If
End Sub

Sub FindSubFolders(objFolder)
	Set colFolders = objFolder.SubFolders
	For Each objSubFolder In colFolders
		Set colFiles = objSubFolder.Files
		On Error Resume Next 
		For Each objFile In colFiles
			AddInArray(objFile.Path)
		Next
		FindSubFolders(objSubFolder)
	Next
End Sub

Sub FindFiles(objFolder)
	Set colFiles = objFolder.Files
	For Each objFile In colFiles
		AddInArray(objFile.Path)
	Next
	FindSubFolders(objFolder)
End Sub

' Sub KillProcess(ProcesssName)
	' strComputer = "."
	' Set objWMIService = GetObject("winmgmts:" _ 
		' & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	' Set colProcessList = objWMIService.ExecQuery _ 
		' ("Select * from Win32_Process Where Name = '" + ProcesssName + "'")
	' For Each objProcess in colProcessList
		' objProcess.Terminate()
	' Next
' End Sub

Sub KillProcess(ProcesssName)
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:" _ 
		& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colProcessList = objWMIService.ExecQuery _ 
		("Select * from Win32_Process Where Name like '%" + ProcesssName + "%'")
	For Each objProcess in colProcessList
		objProcess.Terminate()
	Next
End Sub

KillProcess("station.exe")
'KillProcess("UniCrypt")
StrDate = ReturnStrDate()
lenSourceFile1 = Len(strSourceFile1)
lenSourceFile2 = Len(strSourceFile2)
lenSourceFile3 = Len(strSourceFile3)
newSourceFile1 = StrDate + "_" + strSourceFile1
newSourceFile2 = StrDate + "_" + strSourceFile2
newSourceFile3 = StrDate + "_" + strSourceFile3

Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FolderExists(strFolder1) Then
	Set objFolder = objFSO.GetFolder(strFolder1)
	FindFiles(objFolder)
End If

If objFSO.FolderExists(strFolder2) Then
	Set objFolder = objFSO.GetFolder(strFolder2)
	FindFiles(objFolder)
End If

If objFSO.FolderExists(strFolder3) Then
	Set objFolder = objFSO.GetFolder(strFolder3)
	FindFiles(objFolder)
End If

For i = 0 To UBound(ArrayFiles) - 1
	ChangeFile(ArrayFiles(i))
Next

MsgBox("Завершено переконфигурирование Медков" + chr(13) + "Медок можно запускать")
