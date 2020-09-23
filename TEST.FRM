VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "*\Atextbox.vbp"
Begin VB.Form frmTest 
   BackColor       =   &H80000004&
   Caption         =   "CodeBox"
   ClientHeight    =   4575
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "TEST.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   6360
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox FIL 
      Height          =   870
      Hidden          =   -1  'True
      Left            =   3735
      Pattern         =   "*.exe"
      System          =   -1  'True
      TabIndex        =   4
      Top             =   2385
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Timer tPOP 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   720
      Top             =   3780
   End
   Begin VB.PictureBox P2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   1440
      Picture         =   "TEST.frx":0E42
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   3735
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2625
      Left            =   990
      ScaleHeight     =   2625
      ScaleWidth      =   1275
      TabIndex        =   1
      Top             =   675
      Width           =   1275
      Begin VB.ComboBox cbCon 
         Height          =   315
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   5370
      End
      Begin CodeBoxLib.CodeBox CodeBox1 
         Height          =   5235
         Left            =   450
         TabIndex        =   0
         Top             =   315
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   9234
      End
      Begin VB.Line Line1 
         Visible         =   0   'False
         X1              =   345
         X2              =   345
         Y1              =   0
         Y2              =   1305
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   360
      Top             =   1395
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Color           =   8388608
      DefaultExt      =   "java"
      DialogTitle     =   "JEditor"
      Filter          =   "Java source code (*.java)|*.java"
      InitDir         =   "E:\JDK 1.3\bin"
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Enabled         =   0   'False
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "&Redo"
         Enabled         =   0   'False
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "&Cut"
         Enabled         =   0   'False
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "C&opy"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuEditOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuEdtSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuCode 
      Caption         =   "&Code"
      Visible         =   0   'False
      Begin VB.Menu mnuCodeMethods 
         Caption         =   "&Methods"
         Begin VB.Menu mnuCodeMethodObj 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu mnuCodeVars 
         Caption         =   "&Properties"
         Begin VB.Menu mnuCodeVarsObj 
            Caption         =   ""
            Index           =   0
         End
      End
   End
   Begin VB.Menu mnuScript 
      Caption         =   "&Script"
      Begin VB.Menu mnuScriptImports 
         Caption         =   "&Imports..."
      End
      Begin VB.Menu mnuStats 
         Caption         =   "&Statistics"
      End
      Begin VB.Menu mnuSepDamn 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCompile 
         Caption         =   "&Compile..."
      End
      Begin VB.Menu mnuExternal 
         Caption         =   "&External"
         Begin VB.Menu mnuExternalApp 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu asd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDOS 
         Caption         =   "&DOS Shell"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpInfo 
         Caption         =   "&Information..."
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'TextBox Demo originally by Rang3r from SC4F
'Edited and improvised by Sushant Pandurangi

'I have a knack for editing things, you might think.
'Please also see the CoolMenu thing for icons in menus

'All that and more on http://sushantshome.tripod.com

Option Explicit
Dim Words As String, Errors As String
Dim Color As Long, OpColor As Long
Dim InSLComment As Boolean, InMLComment As Boolean
Dim bWaitingComment As Boolean, bWaitingCommentClose As Boolean
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Const JDKDir = "e:\jdk 1.3\bin" 'change this

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Const ColourVars = "black blue cyan darkGray gray green lightGray magenta orange pink red white yellow"
Private Const ColourFuncs = "brighter createContext darker decode equals getAlpha getBlue getColor getColorComponents getColorSpace getComponents getGreen getHSBColor getRed getRGB getRGBColorComponents getRGBComponents getTransparency hashCode HSBtoRGB RGBtoHSB toString"

Private Sub CodeBox1_KeyDown(ASCII As Integer)
On Error Resume Next
If ASCII = vbKeyUp Or ASCII = vbKeyDown Then
    If InStr(1, CodeBox1.CurrentLine, "//") > CodeBox1.CaretPixelX + 1 Or InStr(1, CodeBox1.CurrentLine, "//") = 0 Then
        InSLComment = False
    Else
        InSLComment = True
    End If
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
    Words = " private public int char void new return class String extends implements throws import final static true false if else while do for this try catch float long double boolean single synchronized instanceof "
    Errors = " NoClassDefFoundException NullPointerException FileNotFoundException IOException "
    Color = vbBlue
    OpColor = vbRed
    ChDir App.Path
    CodeBox1.Load "readme.txt"
    AddContents
    AddColourItems
    LoadExternals
End Sub

Private Sub Form_Resize()
On Error Resume Next
Picture1.Move 0, 0, ScaleWidth, ScaleHeight
CodeBox1.Move 0, 330, Picture1.ScaleWidth, Picture1.ScaleHeight - 330
cbCon.Move 0, 0, CodeBox1.Width
Line1.Y2 = Picture1.ScaleHeight
End Sub

Private Sub mnuCodeMethodObj_Click(Index As Integer)
CodeBox1.SelText = mnuCodeMethodObj(Index).Caption
End Sub

Private Sub mnuCodeVarsObj_Click(Index As Integer)
CodeBox1.SelText = mnuCodeVarsObj(Index).Caption
End Sub

Private Sub mnuCompile_Click()
On Error Resume Next
If CD.FileName = "" Then MsgBox "Please save the file first.", vbExclamation: Exit Sub
Shell FullPath(JDKDir, "javac.exe -verbose ") & Chr(34) & CD.FileName & Chr(34), vbNormalFocus
End Sub

Private Sub mnuDOS_Click()
On Error GoTo hell
Shell "C:\command.com", vbNormalFocus
Exit Sub
hell:
MsgBox Error, vbExclamation
End Sub

Private Sub mnuEditExit_Click()
End
End Sub

Private Sub mnuEditOpen_Click()
On Error GoTo hell
CD.ShowOpen
CodeBox1.Load CD.FileName
InMLComment = False: InSLComment = False
AddContents
hell:
End Sub

Private Sub mnuEditPaste_Click()
On Error Resume Next
Dim strs As String, s() As String, i As Long
strs = Clipboard.GetText(vbCFText)
strs = Replace(strs, Chr(10), "")
s = Split(strs, Chr(13))
For i = 0 To UBound(s)
CodeBox1.SelText = s(i) & vbCr
Next i
End Sub

Private Sub mnuEditSave_Click()
On Error GoTo hell
CD.ShowSave
CodeBox1.Save CD.FileName
hell:
End Sub

Private Sub codebox1_CanPopup()
PopupMenu mnuEdit
End Sub

Private Sub codebox1_KeyPress(ASCII As Integer)
If Chr$(ASCII) = "/" And Not bWaitingComment And Not InMLComment Then bWaitingComment = True: Exit Sub
If Chr$(ASCII) = "/" And bWaitingComment Then bWaitingComment = False: InSLComment = True: Exit Sub
If ASCII = 13 And InSLComment = True Then InSLComment = False: Exit Sub
If Chr$(ASCII) = "*" And bWaitingComment Then InMLComment = True: bWaitingComment = False: Exit Sub
If Chr$(ASCII) = "*" And InMLComment Then bWaitingCommentClose = True: Exit Sub
If Chr$(ASCII) = "/" And InMLComment And bWaitingCommentClose Then InMLComment = False: bWaitingCommentClose = False: Exit Sub
bWaitingComment = False: bWaitingCommentClose = False
End Sub

Private Sub codebox1_MoveCaret(Column As Long, Row As Long)
Caption = "TextBox Demo: Col " & Column & ", Line " & Row
Refresh
End Sub

Private Sub codebox1_Word(Word As CodeBoxLib.TextWord, NewLine As Boolean)
If Word Is Nothing Then Exit Sub
    On Error Resume Next
        If Right(Word.Word, 6) = "Color." Then
            MousePointer = 11
            tPOP.Enabled = True
        End If
        If InStr(Words, " " + (Word.Word) + " ") > 0 Then
            Word.Color = Color
            Word.KeyWord = True
        ElseIf InStr(Errors, " " + Word.Word + " ") Then
            Word.KeyWord = True
            Word.Color = vbRed
        Else
            Word.Color = CodeBox1.ForeColor
        End If
        If IsNumeric(Word.Word) Then Word.Color = 32768 'numbers
        If Left(Trim(Word.Word), 2) = "/*" Then InMLComment = True
        If Right(Trim(Word.Word), 2) = "*/" Then InMLComment = False: Word.Color = 128: Exit Sub
        If InMLComment Or InSLComment Then Word.Color = 128: Exit Sub
        If Left(Word.Word, 12) = "sushantshome" Then Word.Color = vbRed
End Sub

Function Spruce(lpStr As String) As String
Dim l As Long
'spruce up the string nice and shiny
lpStr = Replace(lpStr, Chr(10), "")
lpStr = Replace(lpStr, Chr(13), "")
lpStr = Replace(lpStr, vbTab, Space(1))
l = InStr(1, lpStr, "//")
If l Then Mid$(lpStr, l, Len(lpStr) - l) = Space$(l)
Do
lpStr = Replace(lpStr, "  ", " ")
Loop While InStr(1, lpStr, "  ")
Spruce = lpStr
End Function

Sub AddColourItems()
On Error Resume Next
Dim CL() As String, i As Long
CL = Split(ColourFuncs, " ")
For i = 0 To UBound(CL)
Load mnuCodeMethodObj(i)
mnuCodeMethodObj(i).Caption = CL(i)
Next i
CL = Split(ColourVars, " ")
For i = 0 To UBound(CL)
Load mnuCodeVarsObj(i)
mnuCodeVarsObj(i).Caption = CL(i)
Next i
End Sub

Private Sub mnuEditUndo_Click()
On Error Resume Next
cbCon.ZOrder vbBringToFront
cbCon.SetFocus
End Sub

Private Sub mnuExternalApp_Click(Index As Integer)
On Error Resume Next
Dim params As String
If CD.FileName = "" Then MsgBox "Please save the file first.", vbExclamation: Exit Sub
params = InputBox$("This application can not be run without command line parameters. Please specify a command line parameter, or click OK to use the suggested one.", "Parameters", Chr(34) & CD.FileName & Chr(34))
If params = "" Then Exit Sub
Shell FullPath(JDKDir, mnuExternalApp(Index).Caption) & " " & params, vbNormalFocus
End Sub

Private Sub mnuHelpInfo_Click()
On Error Resume Next
Dim asd As String
asd = FullPath(App.Path, "readme.txt")
InMLComment = False: InSLComment = False
CodeBox1.Load asd
End Sub

Private Sub mnuScriptImports_Click()
On Error Resume Next
Dim i As Long, s As String, mn As String
Dim lEnd As Long
s = CodeBox1.Text
Do While InStr(i, s, "import ") > 0
i = InStr(i + 1, s, "import ")
If i = 0 Then Exit Do
lEnd = InStr(i + 1, s, ";")
If lEnd = 0 Then GoTo n
mn = mn & Mid$(s, i + 7, lEnd - i - 7) & vbNewLine
n:
Loop
If mn = "" Then mn = "None."
MsgBox "Imports:" & Space$(20) & vbNewLine & vbNewLine & mn, vbInformation, CD.FileTitle
End Sub

Private Sub mnuStats_Click()
On Error Resume Next
Dim s As String
s = "This file contains " & CodeBox1.LineCount & " lines of code and " & cbCon.ListCount & " Members."
s = s & vbNewLine & "Timestamp: " & FileDateTime(CD.FileName)
s = s & " (" & Round(FileLen(CD.FileName) / 1024, 2) & " KB file)"
MsgBox s, vbInformation
End Sub

Private Sub tPOP_Timer()
On Error Resume Next
SetCursorPos Left / 15 + CodeBox1.CaretPixelX, Top / 15 + 930 / 15 + CodeBox1.CaretPixelY + 22
PopupMenu mnuCode, vbPopupMenuCenterAlign, CodeBox1.CaretPixelX * 15, ((CodeBox1.CaretPixelY + 16) * 15) + 330
MousePointer = 0
tPOP.Enabled = False
End Sub

Sub AddContents()
On Error Resume Next
cbCon.Clear
Dim i As Long, s As String
Dim pub As Long, pri As Long
Dim lEnd As Long
s = CodeBox1.Text
Do While InStr(i, s, "public ") > 0
i = InStr(i + 1, s, "public ")
If i = 0 Then Exit Do
lEnd = InStr(i + 1, s, "}")
If lEnd = 0 Then lEnd = Len(s)
If InStr(i + 1, s, ";") < lEnd Then lEnd = InStr(i + 1, s, ";")
If InStr(i + 1, s, vbNewLine) < lEnd Then lEnd = InStr(i + 1, s, vbNewLine)
If InStr(i + 1, s, "=") < lEnd Then lEnd = InStr(i + 1, s, "=")
If lEnd = 0 Then Exit Do
cbCon.AddItem Trim$(Mid$(s, i, lEnd - i))
pub = pub + 1
n:
Loop
i = 0
Do While InStr(i, s, "private ") > 0
i = InStr(i + 1, s, "private ")
If i = 0 Then Exit Do
lEnd = InStr(i + 1, s, "}")
If lEnd = 0 Then lEnd = Len(s)
If InStr(i + 1, s, ";") < lEnd Then lEnd = InStr(i + 1, s, ";")
If InStr(i + 1, s, vbNewLine) < lEnd Then lEnd = InStr(i + 1, s, vbNewLine)
If InStr(i + 1, s, "=") < lEnd Then lEnd = InStr(i + 1, s, "=")
If lEnd = 0 Then Exit Do
cbCon.AddItem Trim$(Mid$(s, i, lEnd - i))
pri = pri + 1
nx:
Loop
cbCon.Text = pub & " Public, " & pri & " Private Methods and Properties - " & CD.FileTitle
End Sub

Sub LoadExternals()
On Error Resume Next
FIL.Path = JDKDir
Dim i As Long
For i = 0 To FIL.ListCount - 1
Load mnuExternalApp(i)
mnuExternalApp(i).Caption = FIL.List(i)
Next i
End Sub

Function FullPath(lpPath As String, lpFile As String) As String
If Right(lpPath, 1) <> "\" Then lpPath = lpPath & "\"
FullPath = lpPath & lpFile
End Function
