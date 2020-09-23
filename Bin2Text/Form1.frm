VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bin2Text Converter"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Encode Text"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Decode Binary"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1680
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   885
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Text:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Binary:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim Length As Single
Dim BinaryString As String
Dim i As Integer
  Text2.Text = ""
  BinaryString = Text1.Text
  Do Until InStr(1, BinaryString, Chr(10)) = 0
    BinaryString = Left(BinaryString, InStr(1, BinaryString, Chr(10)) - 1) & Right(BinaryString, Len(BinaryString) - InStr(1, BinaryString, Chr(10)))
  Loop
  Do Until InStr(1, BinaryString, Chr(13)) = 0
    BinaryString = Left(BinaryString, InStr(1, BinaryString, Chr(13)) - 1) & Right(BinaryString, Len(BinaryString) - InStr(1, BinaryString, Chr(13)))
  Loop
  Do Until InStr(1, BinaryString, Chr(32)) = 0
    BinaryString = Left(BinaryString, InStr(1, BinaryString, Chr(32)) - 1) & Right(BinaryString, Len(BinaryString) - InStr(1, BinaryString, Chr(32)))
  Loop
  If Len(BinaryString) Mod 8 <> 0 Then
    MsgBox "Invalid Message!" & Chr(10) + Chr(10) & "The binary message is over by " & Len(BinaryString) Mod 8 & IIf(Len(BinaryString) Mod 8 = 1, " char", " chars") & " (short by " & 8 - (Len(BinaryString) Mod 8) & ")", vbOKOnly, "Invalid Message"
    Exit Sub
  End If
  Length = (Len(BinaryString) / 8)
  For i = 0 To Length - 1
    Text2.Text = Text2.Text & Chr(GetDecimalFromBinary(Mid(BinaryString, (i * 8) + 1, 8)))
  Next
End Sub

Private Sub Command2_Click()
Dim i As Integer
Dim Length As Single
Dim Binary As String
  Text1.Text = ""
  Length = Len(Text2.Text)
  For i = 1 To Length
    Binary = GetBinaryFromDecimal(Asc(Mid(Text2.Text, i, 1)))
    Text1.Text = Text1.Text & Binary
  Next
End Sub

Private Function GetBinaryFromDecimal(ByVal Number As Integer) As String
Dim i As Integer, j As Integer
Dim BinaryNum As String
Dim CheckNum As Integer
  CheckNum = 128
  For j = 0 To 7
    If CheckNum > Number Then
      i = 0
    Else
      i = 1
      Number = Number - CheckNum
    End If
    BinaryNum = BinaryNum & i
    CheckNum = CheckNum / 2
  Next
  GetBinaryFromDecimal = BinaryNum
End Function

Private Function GetDecimalFromBinary(ByVal Number As String) As Integer
Dim i As Integer, j As Integer
Dim DecNum As Integer
  j = 128
  For i = 1 To 8
    If Mid(Number, i, 1) = "1" Then
      DecNum = DecNum + j
    ElseIf Mid(Number, i, 1) = "0" Then
      DecNum = DecNum + 0
    End If
    j = j - (j / 2)
  Next i
  GetDecimalFromBinary = DecNum
End Function
