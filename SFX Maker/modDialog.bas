Attribute VB_Name = "modDialog"
Option Explicit

Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'Folders Show
Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const MAX_PATH = 260

Public Type BROWSEINFO
   hOwner           As Long
   pidlRoot         As Long
   pszDisplayName   As String
   lpszTitle        As String
   ulFlags          As Long
   lpfn             As Long
   lParam           As Long
   iImage           As Long
End Type


Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long

Type CHOOSECOLOR
lStructSize As Long
hwndOwner As Long
hInstance As Long
rgbResult As Long
lpCustColors As String
flags As Long
lCustData As Long
lpfnHook As Long
lpTemplateName As String
End Type


Private Type OPENFILENAME
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type
    
Private Const OFN_READONLY = &H1
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_SHOWHELP = &H10
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOLONGNAMES = &H40000 ' force no long names for 4.x modules
Private Const OFN_EXPLORER = &H80000 ' new look commdlg
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_LONGNAMES = &H200000 ' force long names for 3.x modules
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0

Public Function SaveDialog(Form1 As Form, Filter As String, Title As String, InitDir As String, DefaultFilename As String) As String

'Syntax:
'varFileName = SaveDialog(Form1, Filter, Title, InitDir, _
'DefaultFilename)

'Text1 = varFileName
  
  Dim ofn As OPENFILENAME
  Dim a As Long
  On Local Error Resume Next
  ofn.lStructSize = Len(ofn)
  ofn.hwndOwner = Form1.hWnd
  ofn.hInstance = App.hInstance
  If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"

  For a = 1 To Len(Filter)
      If Mid$(Filter, a, 1) = "|" Then Mid$(Filter, a, 1) = Chr$(0)
  Next
  ofn.lpstrFilter = Filter
  ofn.lpstrFile = Space$(254)
  Mid(ofn.lpstrFile, 1, 254) = DefaultFilename
  ofn.nMaxFile = 255
  ofn.lpstrFileTitle = Space$(254)
  ofn.nMaxFileTitle = 255
  ofn.lpstrInitialDir = InitDir
  ofn.lpstrTitle = Title
  ofn.lpstrDefExt = "pdf"
  ofn.flags = OFN_HIDEREADONLY + OFN_OVERWRITEPROMPT + OFN_CREATEPROMPT + OFN_PATHMUSTEXIST
  a = GetSaveFileName(ofn)


  If (a) Then
      SaveDialog = Trim$(ofn.lpstrFile)
  Else
      SaveDialog = ""
  End If
End Function

Public Function OpenDialog(Form1 As Form, Filter As String, Title As String, InitDir As String) As String

'Syntax:
'varFileName = OpenDialog(Form1, Filter, Title, InitDir)

'Text1 = varFileName
  
  Dim ofn As OPENFILENAME
  Dim a As Long
  On Local Error Resume Next
  ofn.lStructSize = Len(ofn)
  ofn.hwndOwner = Form1.hWnd
  ofn.hInstance = App.hInstance
  If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"

  For a = 1 To Len(Filter)
      If Mid$(Filter, a, 1) = "|" Then Mid$(Filter, a, 1) = Chr$(0)
  Next
  ofn.lpstrFilter = Filter
  ofn.lpstrFile = Space$(254)
  ofn.nMaxFile = 255
  ofn.lpstrFileTitle = Space$(254)
  ofn.nMaxFileTitle = 255
  ofn.lpstrInitialDir = InitDir
  ofn.lpstrTitle = Title
  ofn.flags = OFN_HIDEREADONLY + OFN_FILEMUSTEXIST + OFN_PATHMUSTEXIST
  a = GetOpenFileName(ofn)

  If (a) Then
      OpenDialog = Trim$(ofn.lpstrFile)
  Else
      OpenDialog = ""
  End If
End Function

'*****************************************************************************
'* ExtractFileName
'*****************************************************************************
Public Function ExtractFileName(PathName As String) As String
   Dim x As Integer
   For x = Len(PathName) To 1 Step -1
      If Mid$(PathName, x, 1) = "\" Then Exit For
   Next
   ExtractFileName = Right$(PathName, Len(PathName) - x)
End Function

'*****************************************************************************
'* ExtractFilePath
'*****************************************************************************
Public Function ExtractFilePath(PathName As String) As String
   Dim x As Integer
   For x = Len(PathName) To 1 Step -1
      If Mid$(PathName, x, 1) = "\" Then Exit For
   Next
   ExtractFilePath = Left$(PathName, x - 1)
End Function

' returns True when file exists, and False when not,
'   (returns True for directory also):
Public Function FileExists(ByVal sFileName As String) As Boolean
Dim i As Integer
On Error GoTo NotFound
    
    i = GetAttr(sFileName)
    FileExists = True
    Exit Function

NotFound:
    FileExists = False
End Function

Function IsFilePresent(FileName As String) As Boolean
'Purpose:
'Checks whether or not a file exists at specified path.

If Len(Trim(FileName)) > 0 Then
    If Dir$(FileName, vbArchive + vbDirectory _
    + vbHidden + vbNormal + vbReadOnly + vbSystem + _
    vbVolume) <> "" Then
        IsFilePresent = True
    Else
        IsFilePresent = False
    End If
End If

End Function

Function ShowColor(ByVal uObject As Object)  'Creates the color selection dialog.
Dim cc As CHOOSECOLOR
Dim Custcolor(16) As Long
Dim lReturn As Long
Dim CustomColors() As Byte

cc.lStructSize = Len(cc)
cc.hwndOwner = uObject.hWnd
cc.hInstance = App.hInstance
cc.lpCustColors = StrConv(CustomColors, vbUnicode)
cc.flags = 0
If CHOOSECOLOR(cc) <> 0 Then
ShowColor = cc.rgbResult
CustomColors = StrConv(cc.lpCustColors, vbFromUnicode)
Else
ShowColor = -1
End If
End Function

Public Sub ColorSelect(ByVal uObject As Object)  'Color selection dialog
Dim NewColor As Long
NewColor = ShowColor(uObject)
If NewColor <> -1 Then
uObject.BackColor = NewColor
Else
MsgBox "No color has been selected.", vbInformation
End If
End Sub

Private Function FilterLookup(ByVal sFilters As String, ByVal iCur As Long) As String
Dim iStart As Long
Dim iEnd As Long
Dim S As String

iStart = 1
If sFilters = "" Then Exit Function
Do
    ' Cut out both parts marked by null character
    iEnd = InStr(iStart, sFilters, vbNullChar)
    If iEnd = 0 Then Exit Function
    iEnd = InStr(iEnd + 1, sFilters, vbNullChar)
    If iEnd Then
        S = Mid$(sFilters, iStart, iEnd - iStart)
    Else
        S = Mid$(sFilters, iStart)
    End If
    iStart = iEnd + 1
    If iCur = 1 Then
        FilterLookup = S
        Exit Function
    End If
    iCur = iCur - 1
Loop While iCur
End Function

' This routine also works with open files
' and raises an error if the file doesn't exist.
Function GetAttribute(FileName As String) As String
Dim Result As String, attr As Long

attr = GetAttr(FileName)
' GetAttr also works with directories.
If attr And vbDirectory Then Result = Result & " Directory"
If attr And vbReadOnly Then Result = Result & " ReadOnly"
If attr And vbHidden Then Result = Result & " Hidden"
If attr And vbSystem Then Result = Result & " System"
If attr And vbArchive Then Result = Result & " Archive"
' Discard the first (extra) space.
GetAttribute = Mid$(Result, 2)

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

Public Function DeleteExtention(ByVal strFileName As String) As String
  While (Right(strFileName, 1) <> ".")
    strFileName = Mid(strFileName, 1, Len(strFileName) - 1)
  Wend
  DeleteExtention = strFileName
End Function

'Converts short filename to long filename
Public Function GetLongName(ByVal sShortName As String) As String
' --> (All this modules code) Obtained from : -
' ---> Microsoft's/MSDN's code
' ->  http://support.microsoft.com/default.aspx?scid=kb;EN-US;154822
' ---> The original comments were by them :
     Dim sLongName As String
     Dim sTemp As String
     Dim iSlashPos As Integer

     'Add \ to short name to prevent Instr from failing
     sShortName = sShortName & "\"

     'Start from 4 to ignore the "[Drive Letter]:\" characters
     iSlashPos = InStr(4, sShortName, "\")

     'Pull out each string between \ character for conversion
     While iSlashPos
       sTemp = Dir(Left$(sShortName, iSlashPos - 1), vbNormal + vbHidden + vbSystem + vbDirectory)
       If sTemp = "" Then
         'Error 52 - Bad File Name or Number
         GetLongName = ""
         Exit Function
       End If
       sLongName = sLongName & "\" & sTemp
       iSlashPos = InStr(iSlashPos + 1, sShortName, "\")
     Wend

     'Prefix with the drive letter
     GetLongName = Left$(sShortName, 2) & sLongName

   End Function

'Get system directory
Public Function SysDir()
    Dim Buffer As String, length As Integer, Directory As String
    Buffer = Space$(512)
    length = GetSystemDirectory(Buffer, Len(Buffer))
    Directory = Left$(Buffer, length)
    SysDir = Directory
End Function
