VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rtb 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   7223
      _Version        =   393217
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type tLinkInfo
LinkText As String
LinkDest As String
LinkStart As Long
LinkEnd As Long
End Type

Private Links(99) As tLinkInfo 'upto 100 links.
Private LnkCnt As Integer ' a counter.

Function GetPosOverCursor(Box As Object, x, y) As Long
Dim tmpPnt As POINTAPI, tmpPos As Long
tmpPnt.x = x / Screen.TwipsPerPixelX
tmpPnt.y = y / Screen.TwipsPerPixelY
tmpPos = SendMessage(Box.hwnd, &HD7, 0&, tmpPnt)
GetPosOverCursor = tmpPos
End Function

Private Sub Form_Load()
AddLink "Planet Source Code", "http://www.pscode.com"
rtb.SelText = " - Great Site!" & vbCrLf
AddLink "Email Me!", "mailto:mike15@blueyonder.co.uk"
rtb.SelText = " - Awww.....Go Wan!"
End Sub

Private Sub rtb_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If IsCursorOverLink(rtb, CLng(x), CLng(y)) = True Then
'rtb.MousePointer = rtfCustom
'rtb.MouseIcon = (I NEED A HAND ICON AND I CANT BE BOTHERED TO MAKE ONE) :p
rtb.MousePointer = rtfUpArrow
Me.Caption = GetLinkText(rtb, x, y) & " - " & GetLinkDest(rtb, x, y)
Else
rtb.MousePointer = rtfDefault
Me.Caption = "No Link"
End If
End Sub

Function IsCursorOverLink(Box As Object, x, y) As Boolean
Dim I As Integer, tmpT As Long
tmpT = GetPosOverCursor(Box, x, y)
For I = 0 To LnkCnt
If tmpT > Links(I).LinkStart And tmpT < Links(I).LinkEnd Then
IsCursorOverLink = True
GoTo Done
Else
If I = LnkCnt Then
IsCursorOverLink = False
GoTo Done
End If
End If
Next
Done:
End Function

Sub AddLink(Text As String, Dest As String)
Links(LnkCnt).LinkDest = Dest
Links(LnkCnt).LinkText = Text
Links(LnkCnt).LinkStart = rtb.SelStart
Links(LnkCnt).LinkEnd = rtb.SelStart + Len(Text)
LnkCnt = LnkCnt + 1
rtb.SelUnderline = True
rtb.SelColor = vbBlue
rtb.SelText = Text
rtb.SelColor = vbBlack
rtb.SelUnderline = False
End Sub

Function GetLinkText(Box As Object, x, y) As String
Dim I As Integer, tmpT As Long
tmpT = GetPosOverCursor(Box, x, y)
For I = 0 To LnkCnt
If tmpT > Links(I).LinkStart And tmpT < Links(I).LinkEnd Then
GetLinkText = Links(I).LinkText
GoTo Done
Else
If I = LnkCnt Then
GetLinkText = ""
GoTo Done
End If
End If
Next
Done:
End Function

Function GetLinkDest(Box As Object, x, y) As String
Dim I As Integer, tmpT As Long
tmpT = GetPosOverCursor(Box, x, y)
For I = 0 To LnkCnt
If tmpT > Links(I).LinkStart And tmpT < Links(I).LinkEnd Then
GetLinkDest = Links(I).LinkDest
GoTo Done
Else
If I = LnkCnt Then
GetLinkDest = ""
GoTo Done
End If
End If
Next
Done:
End Function

Private Sub rtb_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If IsCursorOverLink(rtb, CLng(x), CLng(y)) = True Then
MsgBox "ShellExecute aint part of this example...go look it up :)"
End If
End Sub
