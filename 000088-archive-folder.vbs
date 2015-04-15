''
''	SCRIPT:
''		archive-folder.vbs
''
''
''	SCRIPT_ID:
''		88
''
'' 
''	DESCRIPTION:
''		This scriptperforms the following actions:
''		1) Collects files older than x days from a source folder
''		2) Compresses the folder to a archive file using 7za.exe
''		3) Deletes older archived files to keep x amount of files.
''
''
''	VERSION:
''		01	2015-04-15	First version
'' 
''	SUBS AND FUNCTIONs:
''		Function GetScriptPath
''		Function NumberAlign
''		Function ProperDateFs
''		Function ProperDateTime
''		Function RunCommand
''		Sub MakeFolder
''		Sub ProcessSet
''		Sub ScriptDone
''		Sub ScriptInit
''		Sub ScriptRun
'' 		Sub CollectFilesBeforeArchiving
'' 
''	---------------------------------------------------------------------------
''
''
''	Archive files from a folder older than x days last modified to an
''	archive file.
''
''	---------------------------------------------------------------------------




Option Explicit



Dim		gobjFso
Dim		gstrFolder
Dim		gintKeepDays
Dim		gstrPathArchive
Dim		gintCount
Dim		gstrPathArchiver



Call ScriptInit()
Call ScriptRun()
Call ScriptDone()



Sub ScriptInit()
	'gstrFolder = "d:\lazarus" 

	gintCount = 0
	
	gstrPathArchiver = GetProgramPath("pkzip25.exe")
	
	
	Set gobjFso = CreateObject("Scripting.FileSystemObject")
End Sub



Sub ScriptRun()
	Dim		strSets
	Dim		arrSets
	Dim		x
	
	strSets = ReadConfig("", "Sets")
	arrSets = Split(strSets, ";")
	For x = 0 To UBound(arrSets)
		Call ProcessSet(arrSets(x))
	Next
End Sub



Sub ScriptDone()
	Set gobjFso = Nothing
	WScript.Quit(0)
End Sub



Function ProperDateFs(ByVal dtmDateTime)
	''
	''	Convert a system formatted date time to a proper file system date time
	''
	''	Returns the current date time when no date time is specified by dtmDateTime
	''
	''	Returns a date time in format: YYYY-MM-DD
	''
	''	dtmDateTime 
	''
	''	blnFolder3
	''		True: 	Uses '\' as the separator char in the date: YYYY\MM\DD
	''		False:	Uses '-' as the separator char in the date: YYYY-MM-DD
	''
	Dim		strSeperator
	Dim		strResult

	strResult = ""
	
	If Len(dtmDateTime) = 0 Then
		dtmDateTime = Now()
	End If
	
	strSeperator = "-"
	strResult = NumberAlign(Year(dtmDateTime), 4) & strSeperator 
	strResult = strResult & NumberAlign(Month(dtmDateTime), 2) & strSeperator
	strResult = strResult & NumberAlign(Day(dtmDateTime), 2)
	
	ProperDateFs = strResult
End Function '' of Function ProperDateFS



Function NumberAlign(ByVal intNumber, ByVal intLen)
	'	
	'	Returns a number aligned with zeros to a defined length
	'
	'	NumberAlign(1234, 6) returns '001234'
	'
	NumberAlign = Right(String(intLen, "0") & intNumber, intLen)
End Function ' of NumberAlign



Function GetBeforeDateZip(ByVal intDays)
	'
	'	Get the date for the before option of PKZIP. -before=MMDDYYYY
	'
	
	Dim		dtBefore
	Dim		r
	
	' Make a positive number negative (x - (2 * x))
	intDays = intDays - (2 * intDays)

	' Calculate the date intDays before
	dtBefore = DateAdd("d", intDays, Now())
	
	' Return the dtBefore in MMDDYYYY format to be used by PKZIP25 -before option
	GetBeforeDateZip = NumberAlign(Month(dtBefore), 2) & NumberAlign(Day(dtBefore), 2) & NumberAlign(Year(dtBefore), 4)
End Function


Function GetBeforeDateArchive(ByVal intDays)
	'
	'	Get the date for the archive file based on the intDays ago (format: yyyyMmm)
	'
	'
	
	Dim		dtBefore
	Dim		r
	
	' Make a positive number negative (x - (2 * x))
	intDays = intDays - (2 * intDays)

	' Calculate the date intDays before
	dtBefore = DateAdd("d", intDays, Now())
	
	' Return the dtBefore in yyyyMmm to name the archive file by month.
	GetBeforeDateArchive =  NumberAlign(Year(dtBefore), 4) & "M" & NumberAlign(Month(dtBefore), 2)
End Function



Sub CollectFilesBeforeArchiving(ByVal strFolderSource, ByVal strFolderCollect, ByVal intKeepDays)
	Dim		c
	Dim		r
	
	WScript.Echo "Collecting files before archiving, please wait..."
	
	c = "robocopy.exe "
	c = c & Chr(34) & strFolderSource & Chr(34) & " "
	c = c & Chr(34) & strFolderCollect & Chr(34) & " "
	c = c & "*.* "							'' 				All files
	c = c & "/move "						'' 	/move 		the files
	c = c & "/z "							'' 	/z 			copy files in restartable mode 
	c = c & "/s "							'' 	/s 			copy sub dirs
	c = c & "/np "							'' 	/np 		no progress counter aka procent
	c = c & "/r:5 "							''	/r			Restart in 5 secs.
	c = c & "/w:10 "						'' 	/w			Wait bewteen retries for 10 sec.
	c = c & "/minlad:" & intKeepDays & " "  '' 	Not used for intKeepDays for Last Access Date (/minlad)
	
	''c = c & "/create "					''	TEST: Create 0 length files and folder stryucture
	''c = c & "/l " 						''	TEST: Testing, do only log, not actually move files.
	c = c & "/tee " 						''	TEST: Log to file and screen both.
	
	c = c & "/log:robocopy-collect.txt"
	
	WScript.Echo c
	r = RunCommand(c)
	WScript.Echo "CollectFilesBeforeArchiving=" & r
	
End Sub '' of Sub CollectFilesBeforeArchiving


Sub CompressCollectedFiles(strFolderCollect, strPathArchive)
	''
	''	Source: http://sevenzip.sourceforge.jp/chm/cmdline/switches/method.htm
	''
	
	Dim		c
	Dim		r
	
	'' 7za.exe a -r D:\archive-older-then\d_temp.7z D:\archive-older-then\d_temp\*.* 
	
	WScript.Echo
	WScript.Echo "CompressCollectedFiles()" 
	
	c = "7za.exe "
	c =	c & "a "			'' 	Add files to archive
	c = c & "-r "			''	Recurse folders 
	c = c & "-mx9 "			''	Set maximum compression. Ultra compressing
	c = c & Chr(34) & strPathArchive & Chr(34) & " "
	c = c & Chr(34) & strFolderCollect & "\*.*" & Chr(34)
	
	WScript.Echo c
	r = RunCommand(c)
	WScript.Echo "CompressCollectedFiles=" & r
	
	If r = 0 Then
		WScript.Echo "INFO: Compression successful, deleting collection folder: " & strFolderCollect
		Call DeleteFolder(strFolderCollect)
	Else
		WScript.Echo "ERROR: Compression of " & strFolderCollect & " failed with code: " & r
	End If
End Sub '' of Sub CompressCollectedFiles



Sub DeleteFolder(filespec)
	'//////////////////////////////////////////////////////////////////////////////
	'//
	'//	DeleteFolder() -- Delete a folder specified
	'//
	'//	filespec	The name of the folder to delete.
	'//

   	Dim fso
   	
   	Set fso = CreateObject("Scripting.FileSystemObject")
   	fso.DeleteFolder filespec, True
   	Set fso = Nothing
End Sub



Sub ProcessSet(ByVal strSet)
	Dim		stFolderSource
	Dim		intActive           	'' 1=ACTIVE, 0=INACTIVE 
	Dim		strPathArchive 			'' 
	Dim		strFolderArchive   		''
	Dim		strFolderCollect
	Dim		intKeepDays				'' 
	Dim		intKeepArchives			''
	Dim		strCmd
	Dim		i
	Dim		strDateArchive
	Dim		dtmDateArchive
	Dim		intDaysBack
	Dim		strFilenameArchive
	
	WScript.Echo 
	WScript.Echo "ProcessSet(): " & strSet
	
	intActive = Int(ReadConfig(strSet, "Active"))
	If intActive = 1 Then
		stFolderSource = ReadConfig(strSet, "SourceFolder")
		strFolderArchive = ReadConfig(strSet, "FolderArchive")
		
		intKeepDays = Int(ReadConfig(strSet, "KeepDays"))
		intKeepArchives = Int(ReadConfig(strSet, "KeepArchives"))
		
		intDaysBack = intKeepDays - (2 * intKeepDays)
		dtmDateArchive = DateAdd("d", intDaysBack, Now())
		strFilenameArchive = ProperDateFs(dtmDateArchive)

		'strFolderCollect = strFolderArchive & "\" & strSet
		strFolderCollect = "D:\Temp\~" & strSet
		
		strPathArchive = strFolderArchive & "\" & strSet & "\" & strFilenameArchive & ".7z"
		
		WScript.Echo vbTab & "  Source folder : " & stFolderSource
		WScript.Echo vbTab & "Archived folder : " & strFolderArchive
		WScript.Echo vbTab & "      Keep days : " & intKeepDays
		WScript.Echo vbtab & " Collect folder : " & strFolderCollect
		WScript.Echo vbTab & "   Archive date : " & dtmDateArchive
		WScript.Echo vbTab & "   Path archive : " & strPathArchive
		WScript.Echo vbTab & "  Keep archives : " & intKeepArchives
		
		
		Call CollectFilesBeforeArchiving(stFolderSource, strFolderCollect, intKeepDays)
	
		Call CompressCollectedFiles(strFolderCollect, strPathArchive)
		
		
		
	Else
		WScript.Echo "Set " & strSet & " is not active (Active=0)"
	End If
End Sub



Function RunCommand(sCommandLine)
	'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	'//
	'//	RunCommand(sCommandLine)
	'//
	'//	Run a DOS command and wait until execution is finished before the cript can commence further.
	'//
	'//	Input
	'//		sCommandLine	Contains the complete command line to execute 
	'//
	Dim oShell
	Dim sCommand
	Dim	nReturn

	Set oShell = WScript.CreateObject("WScript.Shell")
	sCommand = "CMD /c " & sCommandLine
	' 0 = Console hidden, 1 = Console visible, 6 = In tool bar only
	'LogWrite "RunCommand(): " & sCommandLine
	nReturn = oShell.Run(sCommand, 6, True)
	Set oShell = Nothing
	RunCommand = nReturn 
End Function '' RunCommand



Function ReadConfig(ByVal sSection, ByVal sSetting)
	''
	''	Verbeterde versie 2009-01-13
	''
	''	Reads a setting from a .conf file and returns the value.
	''
	''	Name the file.conf same as the script name but with a conf extension.
	''	
	''	Layout config conf file:
	''	' comment line
	''	[Section1]
	''	Name=Whatever is the biatch
	''	Name=Perry
	''	
	''	[Section2]
	''	Name=Adrian
	''
	''	[Section3]
	''	Name=Jill
	''	------------------
	''
	''	Example looping for more entries:
	''	Dim	x
	''	Dim	bAgain
	''	Dim	sLogEntry
	''	bAgain = True
	''	x = 1
	''	Do
	''		sLogEntry = ConfigReadSetting("LogEntry" & x)
	''	
	''		If IsEmpty(sLogEntry) Then
	''			bAgain = False
	''		Else
	''		WScript.Echo x & ": [" & sLogEntry & "]"
	''		End If
	''		x = x + 1
	''	Loop Until bAgain = False
	''
	''	Remark: Convert strings to Integers for numbers
	''		n = Int(ConfigReadSetting("", "Number"))
	''
	Const	FOR_READING = 1		'== Read mode for config file. Read only
	Const	SEP = "="			'== The char =that seperates the setting and it value
	
	Dim	oFso
	Dim	sPath
	Dim	oFile		
	Dim	bFoundValue		'== Boolean for found value
	Dim	sLine			'== Lime buffer tpo hold the complete line
	Dim	bInSection		'== Is the value in asection
	Dim	bFoundSection	'== Is the section found
	Dim	sReturn			'== Return value from this function
	
	bFoundValue = False
	bInSection = False
	bFoundSection = False
	
	Set oFso = CreateObject("Scripting.FileSystemObject")
	
	''	Replace the .vbs for a .conf extension from the script path
	sPath = Replace(WScript.ScriptFullName, ".vbs", ".conf")
	
	Set oFile = oFso.OpenTextFile(sPath, FOR_READING)
	
	'== Surround the section text with square brackets
	sSection = "[" & sSection & "]"
	
	'WScript.Echo "sSection="&sSection
	
	if sSection = "[]" then 
		'== No section is specified, returns the first occurance of sSetting
		do
			sLine = oFile.ReadLine
			'WScript.Echo vbTab & "sLine:"&vbTab&sLine
			if (InStr(sLine, SEP) > 0) and (Left(sLine, 1) <> "'") Then
				'WScript.Echo "Normale regel"
				If InStr(sLine, sSetting) > 0 Then
					sReturn = Right(sLine, Len(sLine) - InStr(sLine, SEP))
					bFoundValue = True
				End If
			end if
		loop until (bFoundValue = true) or (oFile.AtEndOfStream = true)
	else
		'== Section specitied. First search for the section.
		sLine = ""
		do
			sLine = oFile.ReadLine
		
			if sSection = sLine then 
				bFoundSection = true
			end if
			
			'== Only return a value if:
			'== 1) in the line is a seperator char (InStr(sLine, SEP) > 0)
			'== 2) the line is not a comment (Left(sLine, 1) <> "'")
			'== 3) are we in the specified section (bFoundSection = true)
			if (InStr(sLine, SEP) > 0) and (Left(sLine, 1) <> "'") and (bFoundSection = true) Then
				If InStr(sLine, sSetting) > 0 Then
					sReturn = Right(sLine, Len(sLine) - InStr(sLine, SEP))
					bFoundValue = True
				End If
			end if
		loop until (bFoundValue = true) or (oFile.AtEndOfStream = true)
	end if

	'== Close the file
	oFile.Close
	Set oFile = Nothing
	
	ReadConfig = sReturn
End Function '== ReadConfig



Sub MakeFolder(ByVal sNewFolder)
	''
	'' MakeFolder(strNewFolder)
	''
	'' Create a folder structure.
	''
	'' Parameters:
	''	sNewFolder	Contains the path of the folder structure
	''			e.g. C:\This\Is\A\New\Folder or
	''			\\server\share\folder\folder
	''
	'	Added
	'		When the path contains a file name (d:\folder\file.ext)
	'		It will be deleted first.
	'	
	'' Returns:
	''	True		Folder created.
	''	False		Folder could not be created.
	''
	Dim	arrFolder
	Dim	c
	Dim	intCount
	Dim	intRootLen
	Dim	objFSO
	Dim	strCreateThis
	Dim	strPathToCreate
	Dim	strRoot
	Dim	x
	Dim	bReturn

	bReturn = False

	' If the sNewFolder contains a file name (d:\folder\file.ext)
	' Return only the path and delete file.ext from the sNewFolder.
	If InStrRev(sNewFolder, ".") > 0 Then
		sNewFolder = Left(sNewFolder, InStrRev(sNewFolder, "\") - 1)
	End If
		
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	If objFSO.FolderExists(sNewFolder) = False Then
		'' WScript.Echo "Folder " & sNewFolder & " does not exists, creating it."
		If Right(sNewFolder, 1) = "\" Then
			sNewFolder = Left(sNewFolder, Len(sNewFolder) - 1)
		End If

		If Mid(sNewFolder, 2, 1) = ":" Then
			' Path contains a drive letter (e.g. 'D:')
			intRootLen = 2 
			strPathToCreate = Right(sNewFolder, Len(sNewFolder) - intRootLen)
			strRoot = Left(sNewFolder, intRootLen)
		Else
			' Path contains a share name (e.g. '\\server\share')
			intCount = 0
			intRootLen = 0
			For intRootLen = 1 To Len(sNewFolder)
				c = Mid(sNewFolder, intRootLen, 1)
				If c = "\" Then
					intCount = intCount + 1
				End If
				If intCount = 4 Then
					Exit For
				End If
			Next
			intRootLen = intRootLen - 1
			strPathToCreate = Right(sNewFolder, Len(sNewFolder) - intRootLen)
			strRoot = Left(sNewFolder, intRootLen)
		End If

		arrFolder = Split(strPathToCreate, "\")
		strCreateThis = strRoot
		
		For x = 1 To UBound(arrFolder)
			strCreateThis = strCreateThis & "\" & arrFolder(x)
		
			's = s & "\" & arrFolder(x)
			If Not objFSO.FolderExists(strCreateThis) Then
				On Error Resume Next
				objFSO.CreateFolder strCreateThis
				If Err.Number <> 0 Then
					WScript.Echo "MakeFolder: Error: Can't create " & strCreateThis
				End If
			End If
		Next
	End If

	Set objFSO = Nothing
End Sub '' MakeFolder



Function GetProgramPath(sProgName)
	'==
	'==	Locates a command line program in the path of the user,
	'==	or in the current folder where the script is started.
	'==
	'==	Returns:
	'==		Path to program when found
	'==		Blank string when program is not found
	'==
	Dim	oShell
	Dim	sEnvPath
	Dim	oColVar
	Dim	aPath
	Dim	sScriptPath
	Dim	sScriptName
	Dim	x
	Dim	oFso
	Dim	sPath
	Dim	sReturn

	sReturn = "GetProgramPath() COULD NOT FIND " & sProgName

	Set oFso = CreateObject("Scripting.FileSystemObject")
	Set oShell = CreateObject("WScript.Shell")
	
	sScriptPath = WScript.ScriptFullName
	sScriptName = WScript.ScriptName

	sScriptPath = Left(sScriptPath, Len(sScriptPath) - Len(sScriptName))
	
	'=
	'=	Build the path string like:
	'=		folder;folder;folder;...
	'=
	'=	Place the current folder first in line. So it will find the file first when
	'=	it is in the same folder as the script.
	'=
	sEnvPath = sScriptPath & ";" & oShell.ExpandEnvironmentStrings("%PATH%")
	
	'WScript.Echo sEnvPath
	aPath = Split(sEnvPath, ";")
	For x = 0 To UBound(aPath)
		If Right(aPath(x), 1) <> "\" Then
			aPath(x) = aPath(x) & "\"
		End If
		
		'WScript.Echo x & ": " & aPath(x)
		sPath = aPath(x) & sProgName
		'WScript.Echo sPath
		If oFso.FileExists(sPath) = True Then
			sReturn = sPath
			Exit For
		End If
		
	Next
	
	Set oShell = Nothing
	Set oFso = Nothing
	'= Return the string with double quotes enclosed. For paths with spaces.
	'GetProgramPath = Chr(34) & sReturn & Chr(34)
	'= 2011-02-16 Removed the Chr(34); was not working.
	GetProgramPath = sReturn
End Function '' GetProgramPath


Function RunCommand(sCommandLine)
	''
	''	RunCommand(sCommandLine)
	''
	''	Run a DOS command and wait until execution is finished before the script can commence further.
	''
	''	Input
	''		sCommandLine	Contains the complete command line to execute 
	''
	Dim oShell
	Dim sCommand
	Dim	nReturn

	Set oShell = WScript.CreateObject("WScript.Shell")
	sCommand = "CMD /c " & sCommandLine
	' 0 = Console hidden, 1 = Console visible, 6 = In tool bar only
	'LogWrite "RunCommand(): " & sCommandLine
	nReturn = oShell.Run(sCommand, 6, True)
	Set oShell = Nothing
	RunCommand = nReturn 
End Function '' RunCommand

''	EOS