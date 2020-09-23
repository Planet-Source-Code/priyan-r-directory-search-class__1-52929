VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsdirsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'======================================================
' This class allows you to search a directory for the given pattern
' This class uses windows API to search so this is fast
' Programmed By Priyan

' Visit me at www.priyan.tk

' Mail me at admin@priyan.tk

'You can use this in your projects leave the above comments there
'======================================================
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Const MAX_PATH = 260
Const MAXDWORD = &HFFFF
Const INVALID_HANDLE_VALUE = -1
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100


Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type


Private Type WIN32_FIND_DATA
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
Dim search_subfolders As Boolean
Dim files As New Collection
Dim cancel_ As Boolean
Dim pause_ As Boolean
Dim searchpatterns As New Collection
Public Event searching(ByVal path$)
Public Event found(ByVal path$, file$)

Private Sub searchfiles(ByVal path$)
Dim handle&
Dim data As WIN32_FIND_DATA
Dim dirname$, filename$
Dim dirNames As New Collection
Dim cont As Boolean, i%
If Right(path, 1) <> "\" Then path = path & "\"
RaiseEvent searching(path)
'find first file
    handle = FindFirstFile(path & "*", data)
    cont = True
    If handle <> INVALID_HANDLE_VALUE Then
        Do While cont
        DoEvents
            dirname = StripNulls(data.cFileName)
             ' Ignore the current and .. directories.
            If (dirname <> ".") And (dirname <> "..") Then
                ' if it is a folder
                If GetFileAttributes(path & dirname) And FILE_ATTRIBUTE_DIRECTORY Then
                        dirNames.Add dirname
                End If
            End If
            cont = FindNextFile(handle, data)  'Get next subdirectory.
            checkpaused 'check is paused
            If cancel_ = True Then Exit Sub
        Loop
        cont = FindClose(handle)
    End If
    'find the files in the current folder
    'searches the  current folder for each search pattern
    For i = 1 To searchpatterns.Count
        handle = FindFirstFile(path & searchpatterns(i), data)
        cont = True
        If handle <> INVALID_HANDLE_VALUE Then
            While cont
                DoEvents
                filename = StripNulls(data.cFileName)
                ' Ignore the current and .. directories.
                    If (filename <> ".") And (filename <> "..") Then
                        If Not GetFileAttributes(path & filename) And FILE_ATTRIBUTE_DIRECTORY Then
                               'if there is two patterns to search then checks for duplicates
                               If additem(path & filename) = True Then
                                    RaiseEvent found(path, filename)
                               End If
                        End If
                    End If
                cont = FindNextFile(handle, data) ' Get next file
                checkpaused 'check is paused
                If cancel_ = True Then Exit Sub
            Wend
            cont = FindClose(handle)
        End If
    Next
    If searchinsubfolders = True Then
        'walk through every sub directories
        For i = 1 To dirNames.Count
            'recursively search each folder
            checkpaused 'check is paused
            If cancel_ = True Then Exit Sub
            searchfiles path & dirNames(i)
            DoEvents
        Next
    End If
End Sub
Private Function StripNulls(str As String) As String
    If (InStr(str, Chr(0)) > 0) Then
        str = Left(str, InStr(str, Chr(0)) - 1)
    End If
        StripNulls = str
End Function

Private Sub checkpaused()
    'check paused
    If pause = True Then
        Do
            If pause = False Or cancel_ = True Then Exit Sub
            DoEvents
        Loop
    End If
End Sub

Public Sub findfiles(ByVal path$, ByVal search$)
    Set files = Nothing
    Set searchpatterns = Nothing
    'cancel all the current searchs
    cancel_ = True
    'now reset cancel
    cancel_ = False
    pause_ = False
    Dim obj, arr() As String
    'splits each search pattern and adds its to searchpatterns collection
    arr = Split(search, "|")
    For Each obj In arr
        searchpatterns.Add obj
    Next
    searchfiles path
End Sub
' adds a filename to the collection if the file not exists it the collection
Private Function additem(ByVal filename$) As Boolean
On Error GoTo ext:
'an error occures the file allready in the collection
files.Add filename, filename
additem = True
Exit Function
ext:
End Function
Public Property Get filesfound() As Collection
    Set filesfound = files
End Property


Public Sub Cancel()
cancel_ = True
End Sub

Public Property Get pause() As Boolean
    pause = pause_
End Property

Public Property Let pause(ByVal vNewValue As Boolean)
    pause_ = vNewValue
End Property


Public Property Get searchinsubfolders() As Boolean
    searchinsubfolders = search_subfolders
End Property

Public Property Let searchinsubfolders(ByVal vNewValue As Boolean)
    search_subfolders = vNewValue
End Property

Private Sub Class_Initialize()
search_subfolders = True
End Sub


