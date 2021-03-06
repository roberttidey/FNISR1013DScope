' Hexdump.vbs - v1.01 11thSeptember 2001
'Dumps first part of file in hex, If length set to 0 dumps all file
'Command line hexdump Offset BlockSize Total Bytes InFile OutFile
' -----------------------------------
' R.J.Tidey  Copyright (C) 2001 R.J.Tidey
' -----------------------------------
Option Explicit
Const MODE_ADMIN				= False 'Set the need to run script in Admin Mode
Const MODE_32BIT				= True	'Set the need to run the script in 32 bit mode
Const LOGFILENAME				= "ScopeDumpLog.txt"
Const LOGFILENAME_BACKUP		= "ScopeDumpLogBackUp.txt"
Const MAX_LOGSIZE				= 2000000
Const vbCOMMA			= ","
Const vbCOLON			= ":"
Const vbQUOTE			= """"
Const vbSPACE			= " "
Const WAIT_TOKEN		= "\WaitElevated.txt" 'used to wait for elevated script to finish
Const WAIT_TIMEOUT		= 10
Const TMP_EXT = "_$$$"
Const TEMPORARYFOLDER = 2
Const forReading = 1
Const forWriting = 2
Const forAppending = 8

Const HDRSIZE		= 1000
Const DEFAULTCOLSPERLINE	= 16
Const FORMAT_SPACED			= 1
Const FORMAT_ASCII			= 0

'Controls what is included in file
'Set the HEX values to 1 to see the raw binary data in hex format as required 
Const HDR_HEX = 1
Const BLOCK1_HEX = 0
Const BLOCK2_HEX = 0
Const BLOCK1_VALS = 1
'Set this to 1 if absolute voltage values wanted
Const SCALE_VALUES = 1

Const CHANNEL_VSCALE = "5.0:V,2.5:V,1.0:V,500:mV,200:mV,100:mV,50:mV"
Const CHANNEL_COUPLING = "DC,AC"
Const CHANNEL_PROBE = "1,10,100"
Const TIMEBASE_HSCALE ="50:S,20:S,10:S,5:S,2:S,1:S,500:mS,200:mS,100:mS,50:mS,20:mS,10:mS,5:mS,2:mS,1:mS,500:uS,200:uS,100:uS,50:uS,20:uS,10:uS,5:uS,2:uS,1:uS,500:nS,200:nS,100:nS,50:nS,20:nS,10:nS"
Const MEASURES = "Vpp,Vrms,Freq,Tim+,Tim-,Cycle,Vavg,Vmax,VMin,Vp,Duty+,Duty-"
Const TRIGTYPE = "Auto,Single,Normal"
Const TRIGEDGE = "Rising,Falling"
Const TRIGCHANNEL = "CH1,CH2"
Const SCROLL = "Fast,Slow"
Const TRIG50 = "Off,On"
Const HDR_CHGAIN = 4
Const HDR_CHPROBE = 10
Const HDR_CHCOUPLING = 8
Const HDR_TIMEBASE = 22
COnst HDR_TRIGTYPE = 26
Const HDR_TRIGEDGE = 28
Const HDR_TRIGCHANNEL = 30
Const HDR_SCREENBRIGHT = 120
Const HDR_GRIDBRIGHT = 122
Const HDR_SCROLL = 24
Const HDR_TRIG50 = 124

'******************************
'Main Script Code goes here
'******************************
	Dim Logging
	Dim ScriptPath
	Dim ScriptName
	Dim ScriptTime
	Dim TestMode 			'Set to True if manually run script
	Dim iFileName			'ApsCom file to convert
	Dim iFile
	Dim iSize
	Dim tFileName			'Temporary work file
	Dim tFile
	Dim lFileName
	Dim lFile				'Log File
	Dim oFileName			'Output Tif filename
	Dim binAcc				'Used to Access binary files
	Dim fso					'File System Object
	Dim WshShell
	DIm ColsPerLine
	Dim DataBlock()			'Holds Data Block for dumping
	Dim vScale(2)
	Dim vUnits(2)
	Dim tScale
	Dim tUnits
	Dim Index
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set WshShell = CreateObject("WScript.Shell")
	ScriptName = fso.GetFileName(WScript.ScriptFullName)
	ScriptPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\"	
	
	InitLogging
	ScriptTime = Timer()
	If RunScript(MODE_ADMIN, MODE_32BIT, "","") Then
		Initialise
		Set binAcc = CreateObject("BinaryAccess.clsBinaryAccess")
		tFileName = fso.GetSpecialFolder(TEMPORARYFOLDER).Path & "\" & fso.GetFilename(iFilename) & TMP_EXT
		lFileName = fso.GetParentFoldername(oFileName)
		If lFileName <> "" Then
			lFileName = lFileName & "\" 
		End If
		If oFileName = "" Then
			lFileName = lFileName & fso.GetBaseName(iFilename) & ".hex"
		Else
			lFileName = lFileName & oFileName
		End If
		If fso.FileExists(tFilename) Then fso.DeleteFile tFilename
		If fso.FileExists(oFilename) Then fso.DeleteFile oFilename
		
		iSize = fso.GetFile(iFileName).Size
		If iSize > 0 Then
			iFile = binAcc.OpenBinary(iFileName)
				
			Set lFile = fso.OpenTextFile(lFilename, forWriting,True)
			lFile.WriteLine  iFileName & " Size &H" & Hex(iSize) & "  " & Now()
			
			'Get and Dump the header
			binAcc.SeekBinary(iFile) = 1
			ReDim DataBlock(HDRSIZE)
			Index = binAcc.ReadBytes(iFile, DataBlock)
			If HDR_HEX = 1 Then LogHdrBlock
			DecodeHdrBlock
			If BLOCK1_HEX = 1 Then
				LogDataBlock "CH1 Data", 1000, 3000
				LogDataBlock "CH2 Data", 4000, 3000
			End If
			If BLOCK2_HEX = 1 Then
				LogDataBlock "CH1 Display", 7000, 1500
				LogDataBlock "CH2 Display", 8500, 1500
			End If
			If BLOCK1_VALS = 1 Then
				LogDataVals "CH1 Values", 1000, 3000, 0
				LogDataVals "CH2 Values", 4000, 3000, 1
			End If
			'LogDataBlock "End Data", 10000, 5000
				
			binAcc.CloseBinary() = iFile
			lFile.Close
			
		End If
		Set binAcc = Nothing
	End If
	
	Set wshShell = Nothing
	Set fso = Nothing
'End of Main Script


'*******************************************
'Sub Routines and Function calls follow here
'*******************************************

'**********************************************************************
'Routine to collect filenames and parameters, allows for manual running
'**********************************************************************
Sub Initialise()
	'This Routine collects the Filenames and parameter
	ColsPerLine = DEFAULTCOLSPERLINE
	iFileName = InputBox("Input FileName", ,"Ch1_100uS_0.5V.wav")
	oFileName = ""
End Sub

'**********************************************************************
'Function to Convert a 4 byte from an array at Offset into a long word
'**********************************************************************
Function Bytes2Long(Blk(),i)
	Dim y
	y = ((Blk(i+3) * 256  + Blk(i+2)) * 256 + Blk(i+1)) * 256 + Blk(i)
	If y <= 2147483647  Then
		Bytes2Long = y
	Else
		Bytes2Long = y - 4294967296
	End If
End Function

Function getMeasure(ch,Index)
	Dim addr
	Dim ty
	Dim val3
	Dim valDiv256
	Dim ret
	If ch = 1 Then addr = 209 Else addr = 257
	addr = addr + 4 * Index
	ty = dataBlock(addr)
	val3 = (dataBlock(addr + 1) * 256 + dataBlock(addr + 2)) * 256 + dataBlock(addr+3)
	valDiv256 = val3 / 256
	getMeasure = Cstr(ty) & "," & CStr(val3) & "," & cStr(valDiv256)
End Function

'**********************************************************************
'Routine to decode hdr block
'**********************************************************************
Function DecodeHdrBlock()
	Dim Index
	Dim scale
	
	scale = Split(CHANNEL_VSCALE,",")(DataBlock(HDR_CHGAIN + 1))
	vScale(0) = Split(scale, vbCOLON)(0)
	vUnits(0) = Split(scale, vbCOLON)(1)
	scale = Split(CHANNEL_VSCALE,",")(DataBlock(HDR_CHGAIN + 11))
	vScale(1) = Split(scale, vbCOLON)(0)
	vUnits(1) = Split(scale, vbCOLON)(1)
	scale = Split(TIMEBASE_HSCALE,",")(DataBlock(HDR_TIMEBASE + 1))
	tScale = Split(scale, vbCOLON)(0)
	tUnits = Split(scale, vbCOLON)(1)
	lFile.WriteLine "CH1 Volts:" & vScale(0) & vUnits(0) & "/div"
	lFile.WriteLine "CH1 Coupling:" & Split(CHANNEL_COUPLING,",")(DataBlock(HDR_CHCOUPLING + 1))
	lFile.WriteLine "CH1 Probe:x" & Split(CHANNEL_PROBE,",")(DataBlock(HDR_CHPROBE + 1))
	lFile.WriteLine "CH1 Volts:" & vScale(1) & vUnits(1) & "/div"
	lFile.WriteLine "CH2 Coupling:" & Split(CHANNEL_COUPLING,",")(DataBlock(HDR_CHCOUPLING + 11))
	lFile.WriteLine "CH2 Probe:" & Split(CHANNEL_PROBE,",")(DataBlock(HDR_CHPROBE + 11))
	lFile.WriteLine "TimeBase:" & tScale & tUnits &  "/div"
	lFile.WriteLine "TrigType:" & Split(TRIGTYPE,",")(DataBlock(HDR_TRIGTYPE + 1))
	lFile.WriteLine "TrigEdge:" & Split(TRIGEDGE,",")(DataBlock(HDR_TRIGEDGE + 1))
	lFile.WriteLine "TrigChannel:" & Split(TRIGCHANNEL,",")(DataBlock(HDR_TRIGCHANNEL + 1))
	lFile.WriteLine "ScreenBright:" & CStr(DataBlock(HDR_SCREENBRIGHT + 1))
	lFile.WriteLine "GridBright:" & CStr(DataBlock(HDR_GRIDBRIGHT + 1))
	lFile.WriteLine "ScrollSpeed:" & Split(SCROLL,",")(DataBlock(HDR_SCROLL + 1))
	lFile.WriteLine "TRIG50:" & Split(TRIG50,",")(DataBlock(HDR_TRIG50 + 1))

	For Index = 0 to 11
		lFile.WriteLine "Measure Ch1:" & Split(MEASURES,",")(Index) & " = " & getMeasure(1, Index)
	Next
	For Index = 0 to 11
		lFile.WriteLine "Measure Ch2:" & Split(MEASURES,",")(Index) & " = " & getMeasure(2, Index)
	Next
	lFile.WriteLine
End Function

'**********************************************************************
'Routine to log a hdr block in Hex format, ColsPerLine entries per line
'**********************************************************************
Function LogHdrBlock()
	Dim Index
	Dim Columns
	Dim Ch
	Dim AsciiString
	Dim MaxIndex
	
	binAcc.SeekBinary(iFile) = 1
	ReDim DataBlock(HDRSIZE)
	Index = binAcc.ReadBytes(iFile, DataBlock)
	MaxIndex = CLng(UBound(DataBlock) / ColsPerLine) * ColsPerLine
	For Index = 0 To MaxIndex - 1
		If Index Mod ColsPerLine = 0 Then
			lFile.Write Right("00000000" & Hex(Index),4) & " "
		End If
		If Index < UBound(DataBlock) Then
			Ch = DataBlock(Index + 1)
			lFile.Write Right("00" & Hex(Ch),2)
		Else
			Ch = 0
			lFile.Write "  "
		End If
		If FORMAT_SPACED = 1 Then lFile.Write " "
		If (Ch > 31) And (Ch < 128) Then
			AsciiString = AsciiString & Chr(Ch)
		Else
			AsciiString = AsciiString & "."
		End If
		Columns = Columns + 1
		If Columns >= ColsPerLine  Then
			If FORMAT_ASCII = 1 Then lFile.Write " " & AsciiString
			AsciiString = ""
			lFile.WriteLine
			Columns = 0
		Else
			If FORMAT_SPACED = 1 Then
				If Columns Mod 8 = 0 Then
					lFile.Write " - "
				End If
			End If
		End If
	Next
	If Columns <> 0 Then
		lFile.WriteLine
	End If
	lFile.WriteLine
End Function

'**********************************************************************
'Routine to log a data block in Hex format, compressed to block similar values
'**********************************************************************
Function LogDataBlock(blockTitle, DataStart, DataSize)
	Dim Index
	Dim blockIndex
	Dim blockVal1
	Dim blockVal2
	Dim blockCount
	Dim Ch
	Dim AsciiString
	Dim MaxIndex
	
	lFile.WriteLine blockTitle & " " & Right("0000" & CStr(DataStart + blockIndex),4) & ":" & Right("0000" & CStr(DataSize),4)
	binAcc.SeekBinary(iFile) = DataStart + 1
	ReDim DataBlock(DataSize)
	Index = binAcc.ReadBytes(iFile, DataBlock)
	blockVal1 = DataBlock(2) * 256 + DataBlock(1)
	blockCount = 0
	blockIndex = 0
	For Index = 0 To DataSize - 2 Step 2
		blockVal2 = DataBlock(Index + 2) * 256 + DataBlock(Index + 1)
		If(Abs(blockVal2 - blockVal1) < 3) Then
			blockCount = blockCount + 1
		Else
			lFile.WriteLine Right("0000" & CStr(DataStart + blockIndex),4) & ":" & Right("0000" & CStr(blockCount),4) & ":" & Right("0000" & Hex(blockVal1),4)
			blockCount = 1
			blockIndex = Index
			blockVal1 = blockVal2
		End If
	Next
	If blockCount > 1 Then
		lFile.WriteLine Right("0000" & CStr(DataStart + blockIndex),4) & ":" & Right("0000" & CStr(blockCount),4) & ":" & Right("0000" & Hex(blockVal1),4)
	End If
End Function

'**********************************************************************
'Routine to log a data block in Values format
'**********************************************************************
Function LogDataVals(blockTitle, DataStart, DataSize, ch)
	Dim Index
	Dim blockVal
	Dim vMult
	DIm tMult
	
	lFile.WriteLine blockTitle & " " & vScale(ch) & vUnits(ch) & "/div " & tScale & tUnits &  "/div"
	lFile.WriteLine "SampleNum,SampleTime,10 values"
	binAcc.SeekBinary(iFile) = DataStart + 1
	ReDim DataBlock(DataSize)
	Index = binAcc.ReadBytes(iFile, DataBlock)
	vMult = CSng(vScale(ch))
	tMult = CSng(tScale)
	For Index = 0 To DataSize - 2 Step 2
		blockVal = DataBlock(Index + 2) * 256 + DataBlock(Index + 1)
		If Index Mod 20  = 0 Then
			lFile.WriteLine
			lFile.Write Right("0000" & CStr(Index / 2),4) & "," & CStr(Index / 100 * tMult) & ","
		End If
		lFile.Write CStr((blockVal - 200) * vMult / 50)
		If Index Mod 20 <> 18 Then
			lFile.Write ","
		End If
	Next
	lFile.WriteLine
	lFile.WriteLine
End Function


'******************************************************************
'Function to get OS64 bit status
'******************************************************************
Function CheckWin64()
	Dim procArch, Wow

	procArch = Ucase(GetEnv("PROCESSOR_ARCHITECTURE", "SYSTEM"))
	Wow      = GetEnv("PROCESSOR_ARCHITEW6432", "SYSTEM")
	If procArch <> "X86" Or Wow <> "" Then
		WriteLog "proc = " & procArch & Wow & "  Assume 64 bit OS"
		CheckWin64 = True
	Else
		WriteLog "proc = " & procArch & Wow & "  Assume 32 bit OS"
		CheckWin64 = False
	End If
End Function

'******************************************************************
'Sub to Get an environment variable
'******************************************************************
Function GetEnv(Key, EnvArea)
	Dim objEnv
	
	Set objEnv = wshShell.Environment(EnvArea )
	GetEnv = objEnv(Key)
	Set objEnv = Nothing
End Function

'******************************************************************
'Sub to initialise logging
'******************************************************************
Sub InitLogging()
	Dim lFile

	If fso.FileExists(ScriptPath & LOGFILENAME) Then
		Logging = 1 'log if file present
		On Error Resume Next
		Set lFile = fso.GetFile(ScriptPath & LOGFILENAME)
		If lFile.Size > MAX_LOGSIZE Then
			If fso.FileExists(ScriptPath & LOGFILENAME_BACKUP) Then
				fso.DeleteFile ScriptPath & LOGFILENAME_BACKUP
			End If
			lFile.Move ScriptPath & LOGFILENAME_BACKUP
		End If
	Else
		Logging = 0
	End If
End Sub

'******************************************************************
'Sub to write a log line
'******************************************************************
Sub WriteLog(Msg)
	Const RETRY_MAX = 5
	Const RETRY_INTERVAL = 1000 'Milliseconds
	Dim logFile
	Dim RetryCount
	
	If Logging <> 0 Then
		For RetryCount = 1 To RETRY_MAX
			On Error Resume Next
			Err.Clear
			Set logFile = fso.OpenTextFile(ScriptPath & "\" & LOGFILENAME, forAppending, True)
			If Err.Number = 0 Then
				logFile.WriteLine Now() & vbCOMMA & CStr(Round(Timer() - ScriptTime, 3)) & vbCOMMA & Msg
				logFile.Close
				Exit For
			End If
			WScript.sleep RETRY_INTERVAL
		Next
	End If
End Sub

'******************************************************************
'Function to run a script or hta in 32/64 bit and / or elevated priviliges mode
' 
'******************************************************************
Function RunScript(ModeAdmin, Mode32, ScriptRequested, ScriptArgs)
	Const ELEVATE_ARGUMENT = "/ELEVATE_UAC"
	Const ELEVATE_VERB	= "runas"
	Const SYS_FOLDER = "\System32\"
	Const SYS_FOLDER32 = "\SysWOW64\"
	Const RUN_SCRIPT = "wscript.exe "
	Const RUN_HTA = "mshta.exe "
	Const EXT_HTA = ".hta"
	Dim NeedToRecurse
	Dim RunScriptApp
	Dim ScriptName
	Dim ElevateVerb
	Dim ElevateOK
	Dim cmdLine
	Dim Index
	Dim oShell
	Dim waitFile
	
	If ScriptRequested <> "" Then
		ScriptName = ScriptRequested
	Else
		ScriptName = WScript.ScriptFullName
	End If
	'On Error Resume Next
	Err.Clear
	NeedToRecurse = False
	ElevateVerb = ""
	ElevateOK = True
	'Check if need to recurse to run in admin mode
	If ModeAdmin Then
		If NeedToElevate() Then
			ElevateVerb = ELEVATE_VERB
			ElevateOK = False
		End If
	End If
	'Check if need to recurse to run in 32 bit
	RunScriptApp = fso.GetSpecialFolder(0) & SYS_FOLDER
	If Mode32 Then
		If CheckWin64() Then
			RunScriptApp =  fso.GetSpecialFolder(0) & SYS_FOLDER32
			ElevateOK = False
		End If
	End If
	If LCase(Right(ScriptName,Len(EXT_HTA))) = EXT_HTA Then
		RunScriptApp = RunScriptApp & RUN_HTA
	Else
		RunScriptApp = RunScriptApp & RUN_SCRIPT
	End If
	If Not ElevateOK Then
		cmdLine = vbQUOTE & ScriptName & vbQUOTE
		If ScriptRequested <> "" Then
			cmdLine = cmdLine & vbSPACE & ScriptArgs
			ElevateOK = False
		Else
			'Recursive call so process local arguments
			If WScript.Arguments.Length = 0 Then
				WriteLog "No arguments. Must be unelevated"
				cmdLine = cmdLine & vbSPACE & ELEVATE_ARGUMENT
				ElevateOK = False
			Else
				If WScript.Arguments(WScript.Arguments.Length - 1) = ELEVATE_ARGUMENT Then
					WriteLog "Elevated Argument found"
					ElevateOK = True
				Else
					'Build the command line this script was called with
					For Index = 0 To Wscript.Arguments.Length - 1
						If Instr(Wscript.Arguments(Index), vbSPACE) > 0 Then
							cmdLine = cmdLine & vbSPACE & vbQUOTE & Wscript.Arguments(Index) & vbQUOTE
						Else
							cmdLine = cmdLine & vbSPACE & WScript.Arguments(Index)
						End If
					Next
					cmdLine = cmdLine & vbSPACE & ELEVATE_ARGUMENT
					ElevateOK = False
				End If
			End If
		End If
		If Not ElevateOK Then 
			WriteLog "Recursive elevation with " & RunScriptApp & vbSPACE & cmdLine
			Set oShell = CreateObject("Shell.Application")
			'Create token to wait for elevated to terminate
			Set waitFile = fso.CreateTextFile(ScriptPath & WAIT_TOKEN, True)
			waitFile.Close
			oShell.ShellExecute RunScriptApp, cmdLine, "", ElevateVerb, 1
			If Err.Number <> 0 Then
				WriteLog "Elevation failed. Try to carry on anyway. " & Err.description
				ElevateOK = True
			End If
			WaitForElevated
		End If
	Else
		ElevateOK = True
	End If
	Set oShell = Nothing
	RunScript = ElevateOK
End Function

Function NeedToElevate()
	Dim strComputer, oWMIService, colOSInfo, oOSProperty, strCaption, bElevate
	strComputer = "."

	Set oWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colOSInfo = oWMIService.ExecQuery("Select * from Win32_OperatingSystem")
	For Each oOSProperty in colOSInfo 
		strCaption = oOSProperty.Caption 
	Next
	If InStr(1,strCaption, "XP", vbTextCompare) Then
		bElevate = False
	ElseIf InStr(1,strCaption, "2003", vbTextCompare) Then
		bElevate = False
	ElseIf InStr(1,strCaption, "2000", vbTextCompare) Then
		bElevate = False
	Else
		'If not 200, XP, or 2003 assume we need to elevate
		bElevate = True
	End If
	Set colOSInfo = Nothing
	Set oWMIService = Nothing
	If bElevate Then
		WriteLog "OS is " & strCaption & " Need to Elevate"
	Else
		WriteLog "OS is " & strCaption & " No Need to Elevate"
	End If
	NeedToElevate = bElevate
End Function

Sub WaitForElevated()
	Dim RetryCount
	
	For RetryCount = 1 To WAIT_TIMEOUT
		If Not fso.FileExists(ScriptPath & WAIT_TOKEN) Then
			Exit For
		End If
		WScript.sleep 1000
	Next
	If fso.FileExists(ScriptPath & WAIT_TOKEN) Then
		WriteLog "Elevated script timed out"
		fso.DeleteFile(ScriptPath & WAIT_TOKEN)
	End If
End Sub	


