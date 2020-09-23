Attribute VB_Name = "modFiles"
Option Explicit

'this part is not by myself. JvB

Public Enum PathTypes
    FileName = 1
    JustName = 2
    FileExtension = 3
    FilePath = 4
    Drive = 5
    LastFolder = 6
    FirstFolder = 7
    LastFolderAndFileName = 8
    DriveAndFirstFolder = 9
    FullPath = 10
End Enum


'Um einen kompletten Verzeichnispfad zu erstellen
Declare Function MakePath Lib "imagehlp.dll" Alias _
    "MakeSureDirectoryPathExists" (ByVal lpPath As String) As Long


Public Function GetPath(ByVal Path As String, Optional ByVal PathType As _
    PathTypes = 1) As String
Dim strPath As String
Dim ThisType As PathTypes
Dim i As Integer
Dim j As Integer

strPath = Path

If InStr(strPath, "\") = 0 And InStr(strPath, ".") > 0 And InStr(strPath, ":") _
    = 0 Then
    ThisType = FileName
ElseIf InStrRev(strPath, "\") = Len(strPath) And Len(strPath) > 3 Then
    ThisType = FilePath
ElseIf Len(strPath) = 3 And Mid(strPath, 2, 2) = ":\" Then
    ThisType = Drive
ElseIf Len(strPath) = 2 And Mid(strPath, 2, 1) = ":" Then
    ThisType = Drive
ElseIf InStrRev(strPath, "\") > InStrRev(strPath, ".") Then
    ThisType = JustName
ElseIf InStr(strPath, "\") > 0 And InStr(strPath, ".") > 0 Then
    ThisType = FullPath
Else
'    MsgBox "Cannot determine the type of the path"
    Exit Function
End If

Select Case PathType
    Case 1
        If ThisType = FullPath Or ThisType = JustName Then
            GetPath = Right(strPath, Len(strPath) - InStrRev(strPath, "\"))
        ElseIf ThisType = FileName Then
            GetPath = strPath
        End If
    Case 2
        If ThisType = FullPath Then
            strPath = StrReverse(strPath)
            i = InStr(strPath, ".") + 1
            j = InStr(strPath, "\")
            strPath = Mid(strPath, i, j - i)
            GetPath = StrReverse(strPath)
        ElseIf ThisType = FileName Then
            GetPath = Left(strPath, InStrRev(strPath, ".") - 1)
        ElseIf ThisType = JustName Then
            GetPath = Right(strPath, Len(strPath) - InStrRev(strPath, "\"))
        End If
    Case 3
        If ThisType = FullPath Or ThisType = FileName Then
            GetPath = Right(strPath, Len(strPath) - InStrRev(strPath, "."))
        End If
    Case 4
        If ThisType = FullPath Or ThisType = JustName Then
            strPath = Left(strPath, InStrRev(strPath, "\") - 1)
        ElseIf ThisType = FilePath Then
            strPath = Left(strPath, Len(strPath) - 1)
        End If
        If Left(strPath, 1) = "\" Then
            strPath = Right(strPath, Len(strPath) - 1)
        End If
        GetPath = strPath
    Case 5
        If ThisType = FilePath Or ThisType = FullPath Or ThisType = Drive Or _
            ThisType = JustName Then
            If Mid(strPath, 2, 1) = ":" Then
                GetPath = Left(strPath, 2)
            End If
        End If
    Case 6
        If ThisType = FullPath Or ThisType = JustName Or ThisType = FilePath _
            Then
            strPath = Left(strPath, InStrRev(strPath, "\") - 1)
            GetPath = Right(strPath, Len(strPath) - InStrRev(strPath, "\"))
        End If
    Case 7
        If Mid(strPath, 2, 1) <> ":" And Left(strPath, 1) <> "\" Then
            strPath = "\" & strPath
        End If
        If ThisType = FullPath Or ThisType = JustName Or ThisType = FilePath _
            Then
            strPath = Right(strPath, Len(strPath) - InStr(strPath, "\"))
            If InStr(strPath, "\") = 0 Then
                Exit Function
            End If
            GetPath = Left(strPath, InStr(strPath, "\") - 1)
        End If
    Case 8
        If ThisType = FullPath Or ThisType = JustName Then
            strPath = Left(strPath, InStrRev(strPath, "\") - 1)
            GetPath = Right(strPath, Len(strPath) - InStrRev(strPath, "\"))
            GetPath = GetPath & Right(Path, Len(Path) - InStrRev(Path, "\") + 1)
        End If
    Case 9
        If ThisType = FullPath Or ThisType = JustName Or ThisType = FilePath _
            Then
            If Mid(strPath, 2, 1) = ":" Then
                strPath = Right(strPath, Len(strPath) - InStr(strPath, "\"))
                GetPath = Left(Path, 3) & Left(strPath, InStr(strPath, "\") - 1)
            End If
        End If
    Case 10
        GetPath = strPath
End Select

End Function



Function FileExists(Path As String) As Boolean
  Const NotFile = vbDirectory + vbVolume
  On Error Resume Next
  
  FileExists = (GetAttr(Path) And NotFile) = 0
  
End Function

Function DirExists(Path As String) As Boolean
  On Error Resume Next
  DirExists = CBool(GetAttr(Path) And vbDirectory)
End Function
