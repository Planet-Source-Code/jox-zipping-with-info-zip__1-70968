VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Zip and Unzip with Info-Zipp Dll - by Jörg von Busekist"
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdClearMessages 
      Caption         =   "Clear"
      Height          =   1095
      Left            =   8400
      TabIndex        =   15
      Top             =   7560
      Width           =   855
   End
   Begin VB.TextBox txtMessages 
      Height          =   1095
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   14
      Top             =   7560
      Width           =   7935
   End
   Begin VB.TextBox txtTargetZipFile 
      Height          =   285
      Left            =   2760
      TabIndex        =   12
      Text            =   "H:\Dokumente und Einstellungen\Admin.JOX\Desktop\Unzip-Beispiel\example.zip"
      Top             =   1200
      Width           =   6375
   End
   Begin VB.TextBox txtFileToZip 
      Height          =   285
      Left            =   2760
      TabIndex        =   10
      Text            =   "H:\Dokumente und Einstellungen\Admin.JOX\Desktop\Unzip-Beispiel\mFileI.bas"
      Top             =   480
      Width           =   6375
   End
   Begin VB.CommandButton cmdZipZip 
      Caption         =   "Zip file"
      Height          =   615
      Left            =   480
      TabIndex        =   9
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton cmdZipDeleteSelected 
      Caption         =   "Delete selected files"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4800
      TabIndex        =   8
      Top             =   3720
      Width           =   4095
   End
   Begin VB.CommandButton cmdZipExtractSelected 
      Caption         =   "Extract selected files"
      Enabled         =   0   'False
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   3720
      Width           =   3735
   End
   Begin VB.CommandButton cmdZipUnzip 
      Caption         =   "Extract all"
      Height          =   615
      Left            =   480
      TabIndex        =   6
      Top             =   2760
      Width           =   1815
   End
   Begin MSComctlLib.ListView lsvZip 
      Height          =   3135
      Left            =   240
      TabIndex        =   4
      Top             =   4320
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   5530
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "imlZip"
      SmallIcons      =   "imlZip"
      ColHdrIcons     =   "imlZip"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtExtractDir 
      Height          =   285
      Left            =   2760
      TabIndex        =   3
      Text            =   "H:\Dokumente und Einstellungen\Admin.JOX\Desktop\Unzip-Beispiel\example\"
      Top             =   2880
      Width           =   6375
   End
   Begin VB.TextBox txtZipfile 
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Text            =   "H:\Dokumente und Einstellungen\Admin.JOX\Desktop\Unzip-Beispiel\example.zip"
      Top             =   2280
      Width           =   6375
   End
   Begin VB.CommandButton cmdZipShow 
      Caption         =   "Show Zip file content"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   2040
      Width           =   1815
   End
   Begin MSComctlLib.ImageList imlZip 
      Left            =   8520
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   9360
      Y1              =   1740
      Y2              =   1740
   End
   Begin VB.Label Label4 
      Caption         =   "Name of the generated Zip-File"
      Height          =   255
      Left            =   2760
      TabIndex        =   13
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "File to zip"
      Height          =   255
      Left            =   2760
      TabIndex        =   11
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Directory where to extract"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Zip file to show or to extract"
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   2040
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClearMessages_Click()

txtMessages.Text = ""

End Sub

Private Sub Form_Load()

'Problem with the DLLs? You can get the DLLs here:
'ftp://ftp.dante.de/tex-archive/tools/zip/info-zip/WIN32/
'files: unz552dN.zip and zip232dN.zip  (08-15-2008)
'there may newer versions

'Rename the Dlls if necessary (=> no DLLs on Planet-Source-Code.com allowed)
If Not FileExists(App.Path & "\unzip32.dll") Or Not FileExists(App.Path & "\zip32.dll") Then
    Call MsgBox("The program didn't found the Info-Zip Dlls in the program path." _
                & vbCrLf & "You can get these DLLs here:" _
                & vbCrLf & "" _
                & vbCrLf & "ftp://ftp.dante.de/tex-archive/tools/zip/info-zip/WIN32/" _
                & vbCrLf & "" _
                & vbCrLf & "or at: http://www.info-zip.org" _
                & vbCrLf & "" _
                & vbCrLf & "files: unz552dN.zip and zip232dN.zip  (08-15-2008)" _
                & vbCrLf & "" _
                & vbCrLf & "there may newer versions" _
                , vbInformation, "No Dlls ?")
End If

'Defaultpaths in the textboxes
txtFileToZip.Text = App.Path & "\*.*"
txtTargetZipFile.Text = App.Path & "\example.zip"
txtZipfile.Text = App.Path & "\example.zip"
txtExtractDir.Text = App.Path & "\example\"

End Sub

Private Sub cmdZipZip_Click()

Dim arFiles() As String   'to hold the filenames
Dim ZipPath As String     'to hold the ZipPath

'If you want to zip all Files in a given directory then set
'arFiles(0) = "*.*" . If Recursion = 2 then it will zip also
'all the files in subdirectories relative to the Zippath.
'If you want to zip all files of a choosen folder including also the
'folder itself (as relative in the zip file) then set
'arFiles(0) = "Foldertozip\*"

'Fill the array with file paths
ReDim Preserve arFiles(1)        'if you put in more files you have to ReDim bigger
arFiles(0) = GetPath(txtFileToZip.Text, FileName)

'Path to zip
ZipPath = GetPath(txtFileToZip.Text, FilePath)

'Call the Zipfunction in modVBZip
Call ZipFiles(txtTargetZipFile.Text, arFiles, ZipPath, 2)

MsgBox "Done."

End Sub

Private Sub cmdZipShow_Click()

'needs an ImageList on the form and its assignment to the listview

Dim i As Long
Dim sIcon As String
Dim itmX As ListItem
Dim Dateinamen() As String      'String array to get the result

lsvZip.ListItems.Clear

'Routine in modVBUnzip
Dateinamen = ZipShow(txtZipfile.Text, txtExtractDir.Text)

'Show all files in the listview
For i = 0 To UBound(Dateinamen) - 2
    If Right(Dateinamen(i), 1) <> "/" Then 'Because the AddIcon routine supports only files
        sIcon = AddIconToImageList(Dateinamen(i), imlZip, "DEFAULT") 'sometimes the Zip-Routine gives back also folders in the form "folder/"
        Set itmX = lsvZip.ListItems.Add(, Dateinamen(i), Dateinamen(i), , sIcon)
        itmX.Icon = sIcon
    End If
Next i
lsvZip.Sorted = True

cmdZipExtractSelected.Enabled = True
cmdZipDeleteSelected.Enabled = True

End Sub

Private Sub cmdZipUnzip_Click()

Call ZipExtractAll(txtZipfile.Text, txtExtractDir.Text)

MsgBox "Done"

End Sub

Private Sub cmdZipExtractSelected_Click()

Dim itmX As ListItem
Dim i As String
Dim arFilesToExtract() As String

i = 0
For Each itmX In lsvZip.ListItems
    If itmX.Selected = True Then
        i = i + 1
        ReDim Preserve arFilesToExtract(i)
        arFilesToExtract(i - 1) = itmX.Text
    End If
Next itmX

Call ZipExtractSingleFiles(txtZipfile.Text, txtExtractDir.Text, arFilesToExtract)

MsgBox "Done"

End Sub

Private Sub cmdZipDeleteSelected_Click()

Dim itmX As ListItem
Dim i As String
Dim arFilesToDelete() As String

i = 0
For Each itmX In lsvZip.ListItems
    If itmX.Selected = True Then
        i = i + 1
        ReDim Preserve arFilesToDelete(i)
        arFilesToDelete(i - 1) = itmX.Text
    End If
Next itmX

Call DeleteFilesFromZip(txtZipfile.Text, arFilesToDelete)

Call cmdZipShow_Click

End Sub

