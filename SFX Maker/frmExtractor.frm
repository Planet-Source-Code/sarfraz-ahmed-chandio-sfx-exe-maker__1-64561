VERSION 5.00
Begin VB.Form frmExtractor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compress It Self-Extractor"
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   Icon            =   "frmExtractor.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   3750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&About"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Extract File"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Created with SFX Maker v1.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   615
      TabIndex        =   3
      Top             =   120
      Width           =   2880
   End
End
Attribute VB_Name = "frmExtractor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sFile As String

Private Sub Command1_Click()
Dim sGivenFilePath As String
Dim sGivenFileData As String
Dim sThisFileData As String
Dim OriginalArray() As Byte
Dim FreeNum As Integer

FreeNum = FreeFile

On Error Resume Next

Open sFile For Binary Access Read As #1
    'Set the buffer
    sThisFileData = Space$(FileLen(sFile))
    Get #1, 1, sThisFileData
Close #1

'Gets All The Data Inside *<BFN>*
sGivenFilePath = Split(Split(sThisFileData, "<BFN>")(1), "</BFN>")(0)
'Gets All The Data Inside *<BFD>*
sGivenFileData = Split(Split(sThisFileData, "<BFD>")(1), "</BFD>")(0)

Open sGivenFilePath For Binary Access Write As #1
    Put #1, 1, sGivenFileData
Close #1

Open sGivenFilePath For Binary As #FreeNum
    ReDim OriginalArray(0 To LOF(FreeNum) - 1)
    Get #FreeNum, , OriginalArray()
Close #FreeNum

'Decompress now
Decompress OriginalArray

Open FormatPath(ExtractFilePath(sFile)) & sGivenFilePath For Binary Access Write As #FreeNum
    Put #FreeNum, , OriginalArray
Close #FreeNum

MsgBox "File extracted successfully!!", vbInformation
End

End Sub

Private Sub Command2_Click()
    MsgBox "SFX Maker developed by: SARFRAZ AHMED CHANDIO  (sarfrazahmed_pk@yahoo.com)", vbInformation
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
sFile = CurDir

If Right(CurDir, 1) <> "\" Then
    sFile = CurDir & "\" & App.EXEName & ".exe"
Else
    sFile = CurDir & App.EXEName & ".exe"
End If
End Sub

Public Function ExtractFilePath(PathName As String) As String
   Dim X As Integer
   For X = Len(PathName) To 1 Step -1
      If Mid$(PathName, X, 1) = "\" Then Exit For
   Next
   ExtractFilePath = Left$(PathName, X - 1)
End Function

Function FormatPath(Path)
'Purpose:
'         Adds a "\" at the end of the given path if it
'hasn't one. If the given path does have "\" at the end
'then this function does nothing...

If Right(Path, 1) = "\" Then
    FormatPath = Path
    Exit Function
Else
    If Trim(Path) <> "" Then
        FormatPath = Path & "\"
    End If
End If

End Function

