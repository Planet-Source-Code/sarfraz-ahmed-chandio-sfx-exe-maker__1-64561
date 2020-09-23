VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SFX Maker"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnAbout 
      Caption         =   "&About"
      Height          =   375
      Left            =   1913
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton btnCreate 
      Caption         =   "&Make EXE"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton btnPath 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4440
      TabIndex        =   0
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox txtPath 
      Height          =   315
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Specify the file you want to make SFX of."
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sFile As String

Private Sub btnAbout_Click()
    MsgBox "Developed by: SARFRAZ AHMED CHANDIO  (sarfrazahmed_pk@yahoo.com)", vbInformation
End Sub

Private Sub btnCreate_Click()
If Trim(txtPath.Text) = "" Then Exit Sub

Dim sGivenFileData As String
Dim sLoaderFileData As String
Dim myArray() As Byte
Dim OriginalArray() As Byte 'array to store the original
Dim FF As Integer

'get the free file
FF = FreeFile

'read the resource file data
myArray = LoadResData(101, "CUSTOM")

'set Extractor.sfx file path
sFile = FormatPath(App.Path) & "Extractor.sfx"

'Create Extractor.sfx and put resource file data in it
If Not FileExists(sFile) Then
'This is the file(EXE) which will prompt the user to
'enter the password after a file is protected.
    Open sFile For Binary Access Write As #FF
        Put #FF, , myArray
    Close #FF
End If

'read the extractor.sfx file data
Open sFile For Binary Access Read As #FF
    sLoaderFileData = Space$(FileLen(sFile))
    Get #FF, 1, sLoaderFileData
Close #FF

Open txtPath.Text For Binary As #FF
    ReDim OriginalArray(0 To LOF(FF) - 1)
    Get #FF, , OriginalArray()
Close #FF




'compress now
Compress OriginalArray

Open txtPath.Text & ".Tmp" For Binary Access Write As #FF
    Put #FF, , OriginalArray
Close #FF





'read the file to protect data
Open txtPath.Text & ".Tmp" For Binary Access Read As #FF
    sGivenFileData = Space$(FileLen(txtPath.Text & ".Tmp"))
    Get #FF, 1, sGivenFileData
Close #FF

'remove ReadOnly attribute from given file
SetAttr txtPath.Text, vbNormal
Kill txtPath.Text
Kill txtPath.Text & ".Tmp"

'Create the protected exe file now
Open FormatPath(ExtractFilePath(txtPath.Text)) & "SFX_" & DeleteExtention(ExtractFileName(txtPath.Text)) & "exe" For Binary Access Write As #FF
    Put #FF, 1, sLoaderFileData & "<BFN>" & ExtractFileName(txtPath.Text) & "</BFN><BFD>" & sGivenFileData & "</BFD>"
Close #FF
        
MsgBox "Self-Extractor created successfully!!", vbInformation
txtPath.Text = ""

End Sub

Private Sub btnExit_Click()
End
End Sub

Private Sub btnPath_Click()
txtPath.Text = OpenDialog(Me, "All Files", "Open", "")
End Sub

Private Sub Form_Load()
'Enable dragging of files over the sfx maker
If Trim(Command$) <> "" Then
    txtPath.Text = GetLongName(Command)
End If
End Sub

Private Sub txtPath_Change()
If txtPath.Text = "" Then
    btnCreate.Enabled = False
Else
    btnCreate.Enabled = True
End If
End Sub
