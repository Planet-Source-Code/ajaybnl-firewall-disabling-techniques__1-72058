Attribute VB_Name = "V"
 
 ' ------------------------------------------------
 ' Récupération des informations
 ' 'Version', 'Type', 'Copyright' et 'Description'
 ' d'un fichier DLL, OCX, EXE ou DRV
 '
 ' Création: webcyril - Février 2001
 ' url: http://www.webcyril.fr.st
 ' ------------------------------------------------
 Option Explicit

 ' API
 Public Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" _
 (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, _
 lpData As Any) As Long
 Public Declare Function GetFileVersionInfoSize Lib "version.dll" Alias _
 "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
 Public Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock _
 As Any, ByVal lpSubBlock As String, lplpBuffer As Long, puLen As Long) As Long
 Public Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (ByVal lpString1 _
 As Any, ByVal lpString2 As Any) As Long
 Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, _
 Source As Any, ByVal Length As Long)

 Public Type VS_FIXEDFILEINFO
 dwSignature As Long
 dwStrucVersion As Long
 dwFileVersionMS As Long
 dwFileVersionLS As Long
 dwProductVersionMS As Long
 dwProductVersionLS As Long
 dwFileFlagsMask As Long
 dwFileFlags As Long
 dwFileOS As Long
 dwFileType As Long
 dwFileSubtype As Long
 dwFileDateMS As Long
 dwFileDateLS As Long
 End Type

 Private Const VFT_APP = &H1
 Private Const VFT_DLL = &H2
 Private Const VFT_DRV = &H3
 Private Const VFT_VXD = &H5

 Public Function HIWORD(ByVal dwValue As Long) As Long
 Dim hexstr As String
 hexstr = Right("00000000" & Hex(dwValue), 8)
 HIWORD = CLng("&H" & Left(hexstr, 4))
 End Function

 Public Function LOWORD(ByVal dwValue As Long) As Long
 Dim hexstr As String
 hexstr = Right("00000000" & Hex(dwValue), 8)
 LOWORD = CLng("&H" & Right(hexstr, 4))
 End Function

 ' Swap de 2 valeurs de type 'byte' avec XOR
 Public Sub SwapByte(byte1 As Byte, byte2 As Byte)
 byte1 = byte1 Xor byte2
 byte2 = byte1 Xor byte2
 byte1 = byte1 Xor byte2
 End Sub

 ' Creation d'une chaine Hexadecimale pour représenter un nombre
 Public Function FixedHex(ByVal hexval As Long, ByVal nDigits As Long) As String
 FixedHex = Right("00000000" & Hex(hexval), nDigits)
 End Function

 Public Sub GetVersionInfo(ByVal sFileName As String, sVersion As String, sType As String, sCopyright As String, sDescription As String)
 Dim vffi As VS_FIXEDFILEINFO ' version info structure
 Dim buffer() As Byte ' buffer for version info resource
 Dim pData As Long ' pointer to version info data
 Dim nDataLen As Long ' length of info pointed at by pData
 Dim cpl(0 To 3) As Byte ' buffer for code page & language
 Dim cplstr As String ' 8-digit hex string of cpl
 Dim retval As Long ' generic return value

 ' Contrôle si le fichier contient des informations
 ' récupérables.
 nDataLen = GetFileVersionInfoSize(sFileName, pData)
If nDataLen = 0 Then
 Exit Sub
End If

 ' Récupération de la 'Version' du fichier
 ' ---------------------------------------
 ' Make the buffer large enough to hold the version info resource.
 ReDim buffer(0 To nDataLen - 1) As Byte
 ' Get the version information resource.
 retval = GetFileVersionInfo(sFileName, 0, nDataLen, buffer(0))

 ' Get a pointer to a structure that holds a bunch of data.
 retval = VerQueryValue(buffer(0), "\", pData, nDataLen)
 ' Copy that structure into the one we can access.
 CopyMemory vffi, ByVal pData, nDataLen
 ' Display the full version number of the file.
 sVersion = Trim(Trim(Str(HIWORD(vffi.dwFileVersionMS))) & "." & _
 Trim(Str(LOWORD(vffi.dwFileVersionMS))) & "." & _
 Trim(Str(HIWORD(vffi.dwFileVersionLS))) & "." & _
 Trim(Str(LOWORD(vffi.dwFileVersionLS))))

 ' Récupération du 'Type' de fichier
 ' ---------------------------------
 Select Case vffi.dwFileType
 Case VFT_APP
 sType = "Application"
 Case VFT_DLL
 sType = "Dynamic Link Library (DLL)"
 Case VFT_DRV
 sType = "Device Driver"
 Case VFT_VXD
 sType = "Virtual Device Driver"
 Case Else
 sType = "Unknown"
 End Select

 ' Récupération du 'Copyright' du fichier
 ' --------------------------------------
 ' Before reading any strings out of the resource, we must first determine the code page
 ' and language. The code to get this information follows.
 retval = VerQueryValue(buffer(0), "\VarFileInfo\Translation", pData, nDataLen)
 ' Copy that information into the byte array.
 CopyMemory cpl(0), ByVal pData, 4
 ' It is necessary to swap the first two bytes, as well as the last two bytes.
 SwapByte cpl(0), cpl(1)
 SwapByte cpl(2), cpl(3)
 ' Convert those four bytes into a 8-digit hexadecimal string.
 cplstr = FixedHex(cpl(0), 2) & FixedHex(cpl(1), 2) & FixedHex(cpl(2), 2) & _
 FixedHex(cpl(3), 2)
 ' cplstr now represents the code page and language to read strings as.

 ' Read the copyright information from the version info resource.
 retval = VerQueryValue(buffer(0), "\StringFileInfo\" & cplstr & "\LegalCopyright", _
 pData, nDataLen)
 ' Copy that data into a string for display.
 sCopyright = Space(nDataLen)
 retval = lstrcpy(sCopyright, pData)
sCopyright = Replace(sCopyright, Chr(0), "")
 ' Récupération de la 'Description' du fichier
 ' -------------------------------------------
 retval = VerQueryValue(buffer(0), "\StringFileInfo\" & cplstr & "\FileDescription", _
 pData, nDataLen)
 sDescription = Space(nDataLen)
 retval = lstrcpy(sDescription, pData)
 sDescription = Replace(sDescription, Chr(0), "")
 End Sub
