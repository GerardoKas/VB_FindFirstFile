Attribute VB_Name = "VBWorkshop"
Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long

Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long


Public Const MAX_PATH = 260
Public Const MAXDWORD = &HFFFF
Public Const INVALID_HANDLE_VALUE = -1
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1

Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100

Type FILETIME
dwLowDateTime As Long
dwHighDateTime As Long
End Type

Type WIN32_FIND_DATA
dwFileAttributes As Long
ftCreationTime As FILETIME
ftLastAccessTime As FILETIME
ftLastWriteTime As FILETIME
nFileSizeHigh As Long
nFileSizeLow As Long
dwReserved0 As Long
dwReserved1 As Long
cFileName As String * MAX_PATH
cAlternate As String * 14
End Type
Public Function StripNulls(OriginalStr As String) As String
If (InStr(OriginalStr, Chr(0)) > 0) Then

OriginalStr = Left(OriginalStr, _
InStr(OriginalStr, Chr(0)) - 1)
End If
StripNulls = OriginalStr
End Function

'======================================
Function FindFilesAPI(path As String, SearchStr As String, FileCount As Integer, Files() As String)

Dim FileName As String ' Walking filename variable...
Dim DirName As String ' SubDirectory Name
Dim dirNames() As String ' Buffer for directory name entries
Dim nDir As Integer ' Number of directories in this path
'Dim Files() As String
Dim i As Integer ' For-loop counter...
Dim hSearch As Long ' Search Handle. Extension
Dim WFD As WIN32_FIND_DATA
Dim Cont As Integer

If Right(path, 1) <> "\" Then path = path & "\"
' Search for subdirectories.
nDir = 0
ReDim dirNames(nDir)
Cont = True
hSearch = FindFirstFile(path & "*", WFD)

If hSearch <> INVALID_HANDLE_VALUE Then
Do While Cont
DirName = StripNulls(WFD.cFileName)
' Ignore the current and encompassing directories.
If (DirName <> ".") And (DirName <> "..") Then
' Check for directory with bitwise comparison.
If FILE_ATTRIBUTE_DIRECTORY And GetFileAttributes(path & DirName) Then
dirNames(nDir) = DirName
'DirCount = DirCount + 1
nDir = nDir + 1
ReDim Preserve dirNames(nDir)
End If
End If
Cont = FindNextFile(hSearch, WFD) 'Get next subdirectory.
Loop
Cont = FindClose(hSearch)
End If

' Walk through this directory and sum file sizes.
'===================================================
hSearch = FindFirstFile(path & SearchStr, WFD)
Cont = True

If hSearch <> INVALID_HANDLE_VALUE Then
While Cont
FileName = StripNulls(WFD.cFileName)
If (FileName <> ".") And (FileName <> "..") Then
FindFilesAPI = FindFilesAPI + (WFD.nFileSizeHigh * _
MAXDWORD) + WFD.nFileSizeLow
ReDim Preserve Files(1, FileCount)
Files(0, FileCount) = path
Files(1, FileCount) = FileName

FileCount = FileCount + 1
'From1.List1.AddItem FileName 'path & FileName   'Directamente añada el archivo a la lista
End If
Cont = FindNextFile(hSearch, WFD) ' Get next file
Wend
Cont = FindClose(hSearch)
End If

' If there are sub-directories...

If nDir > 0 Then
' Recursively walk into them...
For i = 0 To nDir - 1
FindFilesAPI = FindFilesAPI + FindFilesAPI(path & _
dirNames(i) & "\", SearchStr, FileCount, Files())
Next i
End If

End Function    'FIND  FILES API





