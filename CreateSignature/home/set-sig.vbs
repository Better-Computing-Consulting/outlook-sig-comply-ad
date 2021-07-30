On Error Resume Next
usrNetFolder = "\\fpas\Apps\UserSignatures\" & CreateObject("Wscript.Network").userName & "\"
htmFile = usrNetFolder & "firmsignature.htm"
jpgFile = usrNetFolder & "lysr_logo.jpg"
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(htmFile) Then
	dstFolder = CreateObject("Shell.Application").Namespace(&H1a&).Self.Path & "\Microsoft\Signatures\"
	If Not fso.FolderExists(dstFolder) Then
		fso.CreateFolder dstFolder
	End if
	fso.CopyFile htmFile, dstFolder
	fso.CopyFile jpgFile, dstFolder
	Set objShell = CreateObject("WScript.Shell")
	objShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\11.0\Common\MailSettings\NewSignature", "firmsignature", "REG_SZ"
	On Error Resume Next
	objShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Office\11.0\Common\MailSettings\ReplySignature"
	'objShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\11.0\Common\MailSettings\ReplySignature", "firmsignature", "REG_SZ"	

	const HKEY_CURRENT_USER = &H80000001
	Set objreg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
	strKeyPath = "Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\"
	If strProfile = "" Then
		objreg.GetStringValue HKEY_CURRENT_USER, strKeyPath, "DefaultProfile", strProfile
	End If
	myArray = StringToByteArray("firmsignature", True)
	strKeyPath = strKeyPath & strProfile & "\9375CFF0413111d3B88A00104B2A6676"
	objreg.EnumKey HKEY_CURRENT_USER, strKeyPath, arrProfileKeys
	For Each subkey In arrProfileKeys
		strsubkeypath = strKeyPath & "\" & subkey
		objreg.SetBinaryValue HKEY_CURRENT_USER, strsubkeypath, "New Signature", myArray
		objreg.DeleteValue HKEY_CURRENT_USER, strsubkeypath, "Reply-Forward Signature"
		'objreg.SetBinaryValue HKEY_CURRENT_USER, strsubkeypath, "Reply-Forward Signature", myArray
	Next
End If 
Public Function StringToByteArray(Data, NeedNullTerminator)
	Dim strAll
	strAll = StringToHex4(Data)
	If NeedNullTerminator Then
		strAll = strAll & "0000"
	End If
	intLen = Len(strAll) \ 2
	ReDim arr(intLen - 1)
	For i = 1 To Len(strAll) \ 2
		arr(i - 1) = CByte("&H" & Mid(strAll, (2 * i) - 1, 2))
	Next
	StringToByteArray = arr
End Function
Public Function StringToHex4(Data)
	Dim strAll
	For i = 1 To Len(Data)
		' get the four-character hex for each character
		strChar = Mid(Data, i, 1)
		strTemp = Right("00" & Hex(AscW(strChar)), 4)
		strAll = strAll & Right(strTemp, 2) & Left(strTemp, 2)
	Next
	StringToHex4 = strAll
End Function