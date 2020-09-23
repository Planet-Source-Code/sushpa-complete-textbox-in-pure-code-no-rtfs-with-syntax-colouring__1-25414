VERSION 5.00
Begin VB.UserControl CodeBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5055
   DrawWidth       =   2
   KeyPreview      =   -1  'True
   PropertyPages   =   "textbox.ctx":0000
   ScaleHeight     =   4545
   ScaleWidth      =   5055
   Begin VB.PictureBox Filler 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   4770
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   4230
      Width           =   240
   End
   Begin VB.PictureBox BBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillColor       =   &H00400000&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2175
      Left            =   900
      ScaleHeight     =   145
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   145
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1305
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   240
      Left            =   0
      Max             =   255
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4230
      Width           =   4750
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   4250
      Left            =   4800
      Max             =   20
      Min             =   1
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Value           =   1
      Width           =   240
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   960
      Top             =   360
   End
   Begin VB.PictureBox Canvas 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillColor       =   &H00400000&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4140
      Left            =   45
      MousePointer    =   3  'I-Beam
      ScaleHeight     =   276
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   301
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4515
   End
End
Attribute VB_Name = "CodeBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Dim WithEvents line As TextLine
Attribute line.VB_VarHelpID = -1
Public CurrentLine As String
Dim Content As Collection
Dim CaretX As Long
Dim CaretY As Long
Dim FirstVis As Long
Dim LastVis As Long
Dim LastPossibleY As Long
Dim LastPossibleX As Long
Public CaretMode As Long
Dim SellStartX As Long
Dim SellStartY As Long
Dim SellEndX As Long
Dim SellEndY As Long


Public Event Word(Word As TextWord, NewLine As Boolean)
Public Event MoveCaret(Column As Long, Row As Long)
Public Event Draw(Canvas As Object, Word As TextWord, X As Long, Y As Long)
Public Event KeyDown(ASCII As Integer)
Public Event KeyPress(ASCII As Integer)
Public Event CanPopup()

Private Sub Canvas_GotFocus()
   UserControl.SetFocus
End Sub

Private Sub Canvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
   Dim i As Long
   Dim str As String
   Dim tstr As String
   Dim s As Long
   CaretY = Int(Y / BBuffer.TextHeight("X")) + VScroll1.Value
   If CaretY > Content.Count Then CaretY = Content.Count
   If CaretY > LastPossibleY Then CaretY = LastPossibleY
   X = X + (8 * HScroll1.Value)
   str = Content(CaretY).Text
   CurrentLine = str
   CaretX = Len(str)
   For i = 0 To Len(str)
      tstr = Left(str, i)
      s = GetTabbedTextExtent(BBuffer.hdc, tstr, Len(tstr), 1, TabSize) And 65535
      If s > X Then
         CaretX = i - 1
         Exit For
      End If
   Next
   If CaretX = 0 Then HScroll1.Value = CaretX
   SellStartX = CaretX
   SellStartY = CaretY
   SellEndX = CaretX
   SellEndY = CaretY
   Render
   RaiseEvent MoveCaret(CaretX, CaretY)
End Sub


Private Sub Canvas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then RaiseEvent CanPopup
End Sub

Private Sub Canvas_Paint()
   BitBlt Canvas.hdc, 0, 0, Canvas.ScaleWidth, Canvas.ScaleHeight, BBuffer.hdc, 0, 0, vbSrcCopy
End Sub

Private Sub HScroll1_Change()
   Render
End Sub

Private Sub HScroll1_Scroll()
   Render
End Sub

Private Sub Line_Word(Word As TextWord, NewLine As Boolean)
   RaiseEvent Word(Word, NewLine)
End Sub

Private Sub Timer1_Timer()
   RenderCaret
End Sub

Private Sub UserControl_EnterFocus()
   Timer1.Enabled = True
End Sub

Private Sub UserControl_ExitFocus()
   Timer1.Enabled = False
End Sub

Private Sub UserControl_Initialize()
   Set Content = New Collection
   Dim tmp As New TextLine
   Content.Add tmp
   CaretX = 0
   CaretY = 1
   SetScroll
   Render
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode)
   On Error Resume Next
   Dim orgstr As String
   orgstr = Content(CaretY)
   CurrentLine = orgstr
   Select Case KeyCode
      Case vbKeyDown
         CaretY = CaretY + 1
         If CaretY > Content.Count Then CaretY = Content.Count
         If CaretY > LastPossibleY Then VScroll1.Value = VScroll1.Value + 1
         If CaretX > Len(Content(CaretY)) Then CaretX = Len(Content(CaretY))
         RenderCaret
         RaiseEvent MoveCaret(CaretX, CaretY)
      Case vbKeyUp
         CaretY = CaretY - 1
         If CaretY < 1 Then CaretY = 1
         If CaretY < FirstVis Then VScroll1.Value = VScroll1.Value - 1
         If CaretX > Len(Content(CaretY)) Then CaretX = Len(Content(CaretY))
         RenderCaret
         RaiseEvent MoveCaret(CaretX, CaretY)
      Case vbKeyRight
         CaretX = CaretX + 1
         If CaretX > Len(Content(CaretY)) Then CaretX = 0: UserControl_KeyDown vbKeyDown, 0
         RenderCaret
         RaiseEvent MoveCaret(CaretX, CaretY)
      Case vbKeyLeft
         CaretX = CaretX - 1
         If CaretX < 0 And CaretY > 1 Then
            UserControl_KeyDown vbKeyUp, 0
            CaretX = Len(Content(CaretY))
         ElseIf CaretX < 0 Then
            CaretX = 0
         End If
         RenderCaret
         RaiseEvent MoveCaret(CaretX, CaretY)
      Case vbKeyEnd
         CaretX = Len(Content(CaretY))
         SetScroll
         Render
         RaiseEvent MoveCaret(CaretX, CaretY)
      Case vbKeyHome
         CaretX = 0
         HScroll1.Value = 0
         Render
         RaiseEvent MoveCaret(CaretX, CaretY)
      Case vbKeyInsert
         CaretMode = 1 - CaretMode
         Render
         RaiseEvent MoveCaret(CaretX, CaretY)
      Case vbKeyDelete
         If CaretX < Len(orgstr) Then
            DelChar orgstr, CaretX + 1
            Set line = Content(CaretY)
            line.Text = orgstr
            Render
            RaiseEvent MoveCaret(CaretX, CaretY)
         ElseIf CaretY < Content.Count Then
            orgstr = orgstr + Content(CaretY + 1)
            Content.Remove (CaretY + 1)
            Set line = Content(CaretY)
            line.Text = orgstr
            Render
            RaiseEvent MoveCaret(CaretX, CaretY)
         End If
   End Select
   If CaretX = 0 Then HScroll1.Value = CaretX
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
   On Error Resume Next
   If KeyAscii < 32 And KeyAscii <> 8 And KeyAscii <> 13 Then Exit Sub
   Dim orgstr As String
   Dim Word As String
   orgstr = Content(CaretY)
   Select Case KeyAscii
      Case 8
         If CaretX = 0 And CaretY > 1 Then
            Content.Remove CaretY
            UserControl_KeyDown vbKeyUp, 0
            UserControl_KeyDown vbKeyEnd, 0
            Set line = Content(CaretY)
            line.Text = line.Text + orgstr
            Word = Left(orgstr, CaretX)
            Word = Mid(Word, InStrRev(Word, " "))
            'RaiseEvent Word(Word)
            SetScroll
            Render
         Else
            DelChar orgstr, CaretX
            Set line = Content(CaretY)
            line.Text = orgstr
            UserControl_KeyDown vbKeyLeft, 0
            RenderLine
         End If
      Case 13
         Set line = Content(CaretY)
         line.Text = Left(orgstr, CaretX)
         Set line = New TextLine
         line.Text = Mid(orgstr, CaretX + 1)
         Content.Add line, , , CaretY
         SetScroll
         CaretX = 0
         UserControl_KeyDown vbKeyDown, 0
         Render
      Case Else
         orgstr = Content(CaretY)
         Insert orgstr, Chr(KeyAscii), CaretX
         CaretX = CaretX + 1
         Set line = Content(CaretY)
         line.Text = orgstr
         Word = Left(orgstr, CaretX)
         Word = Mid(Word, InStrRev(Word, " "))
         'RaiseEvent Word(Word)
         RenderLine
         If TextWidth(line) > ScaleWidth Then
         SetScroll
         HScroll1.Value = HScroll1.Max
         End If
   End Select
   RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then RaiseEvent CanPopup
End Sub

Private Sub UserControl_Resize()
   On Error Resume Next
    Canvas.Move 45, 45, ScaleWidth - 280, ScaleHeight - 280
    BBuffer.Move 45, 45, Canvas.Width, Canvas.Height
    HScroll1.Move 0, ScaleHeight - HScroll1.Height, ScaleWidth - 225, 225
    VScroll1.Move ScaleWidth - 225, 0, 225, ScaleHeight - 225
    Filler.Move HScroll1.Width, VScroll1.Height
   Render
End Sub

Private Sub Render()
   Dim i As Long
   Dim j As Long
   Dim Pos As Long
   Dim pt As POINTAPI
   Dim xpos As Long
   Dim orgxPos As Long
   Dim rc As RECT
   Dim str As String
   Dim totstr As String
   BBuffer.Cls
   SetTextColor BBuffer.hdc, mForeColor
   For i = VScroll1.Value To Content.Count
      Set line = Content(i)
      xpos = -HScroll1.Value * 8
      orgxPos = xpos
      totstr = ""
      
      For j = 1 To line.Count
         str = line.Word(j).Word
         totstr = totstr + str
         SetTextColor BBuffer.hdc, line.Word(j).Color
         TabbedTextOut BBuffer.hdc, xpos, Pos, str, Len(str), 1, TabSize, orgxPos
         If line.Word(j).KeyWord = True Then
            RaiseEvent Draw(BBuffer, line.Word(j), xpos, Pos)
         End If
         xpos = (GetTabbedTextExtent(BBuffer.hdc, totstr, Len(totstr), 1, TabSize) And 65535) - HScroll1.Value * 8
         If xpos > BBuffer.ScaleWidth Then Exit For
      Next
         ' SetTextColor BBuffer.hdc, vbRed
         ' str = line.Text
         'xpos = -HScroll1.Value * 8
         ' TabbedTextOut BBuffer.hdc, xpos, Pos, str, Len(str), 1, TabSize, xpos
      Pos = Pos + BBuffer.TextHeight("X")
      If Pos > BBuffer.ScaleHeight - (HScroll1.Height / Screen.TwipsPerPixelY) Then Exit For
   Next
   FirstVis = VScroll1.Value
   LastVis = i - 1
   LastPossibleY = Int((BBuffer.ScaleHeight - (HScroll1.Height / Screen.TwipsPerPixelY)) / BBuffer.TextHeight("X")) + VScroll1.Value - 1
   RenderCaret
End Sub

Private Sub RenderCaret()
On Error Resume Next
   Dim xpos As Long
   Dim str As String, char As String
   Static Draw As Boolean
   Draw = Not Draw
   str = Content(CaretY).Text
   If CaretX < 0 Then CaretX = 0
   xpos = GetTabbedTextExtent(BBuffer.hdc, Left(str, CaretX), Len(Left(str, CaretX)), 1, TabSize)
   xpos = xpos - (HScroll1.Value * 8)
   xpos = xpos And 65535
   BitBlt Canvas.hdc, 0, 0, Canvas.ScaleWidth, Canvas.ScaleHeight, BBuffer.hdc, 0, 0, vbSrcCopy
   RenderSelection
   If SellStartX <> SellEndX Then Exit Sub
   Canvas.CurrentX = xpos
   Canvas.CurrentY = (CaretY - VScroll1.Value) * BBuffer.TextHeight("X")
   Canvas.ForeColor = ForeColor
   If Draw = True Then
      If CaretMode = 1 Then
      char = Mid$(str, CaretX + 1, 1)
         Canvas.Line -(Canvas.CurrentX + BBuffer.TextWidth(char), Canvas.CurrentY + BBuffer.TextHeight("X")), , BF
      Else
         Canvas.Line -(Canvas.CurrentX, Canvas.CurrentY + BBuffer.TextHeight("X"))
      End If
   End If
End Sub

Private Sub UserControl_Show()
   Render
End Sub

Private Sub VScroll1_Change()
   Static oldVal As Long
   If VScroll1.Value = oldVal + 1 Then
      ScrollUp
   ElseIf VScroll1.Value = oldVal - 1 Then
      ScrollDown
   End If
   Render
   oldVal = VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
   Render
End Sub

Private Sub Insert(ByRef orgstr As String, NewStr As String, Pos As Long)
   On Error Resume Next
   Dim lstr As String
   Dim rstr As String
   lstr = Left(orgstr, Pos)
   rstr = Mid(orgstr, Pos + 1)
   orgstr = lstr + NewStr + rstr
End Sub

Private Sub DelChar(ByRef orgstr As String, Pos As Long)
   On Error Resume Next
   Dim lstr As String
   Dim rstr As String
   lstr = Left(orgstr, Pos - 1)
   rstr = Mid(orgstr, Pos + 1)
   orgstr = lstr + rstr
End Sub

Private Sub SetScroll()
On Error Resume Next
   If VScroll1.Value > Content.Count Then VScroll1.Value = Content.Count: CaretY = VScroll1.Value
   VScroll1.Max = Content.Count - Round(Canvas.ScaleHeight / BBuffer.TextHeight("X"), 0) ' + BBuffer.TextHeight("X") / 15
   VScroll1.Enabled = (VScroll1.Max > 0)
   HScroll1.Enabled = (HScroll1.Max > 0)
End Sub

Private Sub RenderSelection()
   If SellEndX = SellStartX Then Exit Sub
   Canvas.CurrentX = (SellStartX - HScroll1.Value) * 8
   Canvas.CurrentY = (SellStartY - VScroll1.Value) * BBuffer.TextHeight("X")
   Canvas.ForeColor = RGB(0, 0, 255)
   Canvas.DrawMode = vbMergePen
   Canvas.Line -((SellEndX - HScroll1.Value) * 8, (SellEndY + 1 - VScroll1.Value) * BBuffer.TextHeight("X")), , B
   Canvas.DrawMode = vbCopyPen
End Sub

Public Sub Clear()
   Set line = New TextLine
   Set Content = New Collection
   Content.Add line
   SetScroll
   Canvas_MouseDown 1, 0, 0, 0
   Render
End Sub

Public Sub Load(File As String)
   Dim rows() As String
   Dim id As Long
   Dim Data As String
   Dim i As Long
   id = FreeFile
   Open File For Input As id
      Data = Input(LOF(id), id)
   Close id
   Data = Replace(Data, Chr(10), "")
   rows = Split(Data, Chr(13))
   Set Content = New Collection
   For i = 0 To UBound(rows)
      Set line = New TextLine
      line.Text = rows(i)
      If Left(Trim(line.Text), 2) = "//" Then line.Word(1).Color = COMMENT_COLOR
      Content.Add line
   Next
   SetScroll
   Render
End Sub

Private Sub ScrollUp()
   Dim Y As Long
   BitBlt BBuffer.hdc, 0, 0, BBuffer.ScaleWidth, BBuffer.ScaleHeight, BBuffer.hdc, 0, BBuffer.TextHeight("X"), vbSrcCopy
   Y = LastPossibleY + 1
   RenderLine Y
End Sub

Private Sub ScrollDown()
   Dim Y As Long
   BitBlt BBuffer.hdc, 0, BBuffer.TextHeight("X"), BBuffer.ScaleWidth, BBuffer.ScaleHeight, BBuffer.hdc, 0, 0, vbSrcCopy
   Y = FirstVis - 1
   RenderLine Y
End Sub

Private Sub RenderLine(Optional lineID As Long = 0)
   Dim i As Long
   Dim j As Long
   Dim Pos As Long
   Dim line As String
   Dim tmp As TextLine
   Dim pt As POINTAPI
   Dim xpos As Long
   Dim orgxPos As Long
   Dim rc As RECT
   Dim str As String
   Dim totstr As String
   SetTextColor BBuffer.hdc, mForeColor
   If lineID = 0 Then lineID = CaretY
   i = lineID
   Pos = (i - VScroll1.Value) * BBuffer.TextHeight("X")
   BBuffer.ForeColor = BBuffer.BackColor
   BBuffer.FillColor = BBuffer.BackColor
   BBuffer.Line (0, Pos)-(BBuffer.ScaleWidth, Pos + 15), , B
   
   If i <= Content.Count Then

   
   
   Set tmp = Content(i)
   xpos = -HScroll1.Value * 8
   orgxPos = xpos
   totstr = ""
   For j = 1 To tmp.Count
      str = tmp.Word(j).Word
      totstr = totstr + str
      SetTextColor BBuffer.hdc, tmp.Word(j).Color
      TabbedTextOut BBuffer.hdc, xpos, Pos, str, Len(str), 1, TabSize, orgxPos
      If tmp.Word(j).KeyWord = True Then
         RaiseEvent Draw(BBuffer, tmp.Word(j), xpos, Pos)
      End If
      xpos = (GetTabbedTextExtent(BBuffer.hdc, totstr, Len(totstr), 1, TabSize) And 65535) - HScroll1.Value * 8
      If xpos > BBuffer.ScaleWidth Then Exit For
   Next
End If
   FirstVis = VScroll1.Value
   LastVis = i - 1
   LastPossibleY = Int((BBuffer.ScaleHeight - (HScroll1.Height / Screen.TwipsPerPixelY)) / BBuffer.TextHeight("X")) + VScroll1.Value - 1
   RenderCaret
End Sub

Public Property Get CaretPixelX() As Long
   CaretPixelX = (CaretX - HScroll1.Value) * 8
End Property

Public Property Get CaretPixelY() As Long
   CaretPixelY = (CaretY - VScroll1.Value) * BBuffer.TextHeight("X")
End Property

Public Property Let SelText(Text As String)
On Error Resume Next
   UserControl.SetFocus
   SendKeys Text
End Property

Public Sub AddLines(Text As String)
   Dim rows() As String
   Dim id As Long
   Dim Data As String
   Dim i As Long
   Data = Text
   Data = Replace(Data, Chr(10), "")
   rows = Split(Data, Chr(13))
   For i = 0 To UBound(rows)
      Set line = New TextLine
      line.Text = rows(i)
      Content.Add line
      If (i / 30) And 1 Then DoEvents
   Next
   SetScroll
   Render
End Sub

Public Property Get ForeColor() As OLE_COLOR
ForeColor = Canvas.ForeColor
End Property

Public Property Let ForeColor(vNewValue As OLE_COLOR)
Canvas.ForeColor = vNewValue
mForeColor = Canvas.ForeColor
PropertyChanged "ForeColor"
End Property

'Public Property Get StringColor() As OLE_COLOR
'StringColor = Filler.ForeColor
'End Property
'
'Public Property Let StringColor(vNewValue As OLE_COLOR)
'Filler.ForeColor = vNewValue
'mStringColor = Filler.ForeColor
'PropertyChanged "StringColor"
'End Property

Public Property Get BackColor() As OLE_COLOR
BackColor = BBuffer.BackColor
End Property

Public Property Let BackColor(vNewValue As OLE_COLOR)
BBuffer.BackColor = vNewValue
PropertyChanged "BackColor"
End Property

Public Property Get Text() As String
On Error Resume Next
Dim l As Long
For l = 1 To Content.Count
Text = Text & Content(l) & vbNewLine
Next l
End Property

Public Sub Save(File As String)
Open File For Output As #1
Print #1, Me.Text
Close #1
Load File
End Sub

Function GetVBCase(Exp As String) As String
Select Case LCase(Exp)
Case "withevents"
GetVBCase = "WithEvents"
Case "doevents"
GetVBCase = "DoEvents"
Case "raiseevent"
GetVBCase = "RaiseEvent"
Case Else
GetVBCase = UCase(Left(Exp, 1)) & LCase(Right(Exp, Len(Exp) - 1))
End Select
End Function

Public Property Get LineCount() As Long
LineCount = Content.Count
End Property

Function FindString(lpString As String, Optional Start As Long = 1, Optional bHighLight As Boolean = False) As Long
FindString = InStr(Start, Me.Text, lpString)
If bHighLight Then DelChar "", FindString
End Function
