VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4740
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   4704
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   4704
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnStart 
      Caption         =   "Start"
      Height          =   500
      Left            =   2300
      TabIndex        =   10
      Top             =   3800
      Width           =   1000
   End
   Begin VB.CommandButton btnReset 
      Caption         =   "Reset"
      Height          =   500
      Left            =   1300
      TabIndex        =   9
      Top             =   3800
      Width           =   1000
   End
   Begin VB.CommandButton btn1 
      Enabled         =   0   'False
      Height          =   500
      Left            =   1000
      TabIndex        =   8
      Top             =   1000
      Width           =   500
   End
   Begin VB.CommandButton btn2 
      Enabled         =   0   'False
      Height          =   500
      Left            =   1000
      TabIndex        =   7
      Top             =   2000
      Width           =   500
   End
   Begin VB.CommandButton btn3 
      Enabled         =   0   'False
      Height          =   500
      Left            =   1000
      TabIndex        =   6
      Top             =   3000
      Width           =   500
   End
   Begin VB.CommandButton btn4 
      Enabled         =   0   'False
      Height          =   500
      Left            =   2000
      TabIndex        =   5
      Top             =   1000
      Width           =   500
   End
   Begin VB.CommandButton btn5 
      Enabled         =   0   'False
      Height          =   500
      Left            =   2000
      TabIndex        =   4
      Top             =   2000
      Width           =   500
   End
   Begin VB.CommandButton btn6 
      Enabled         =   0   'False
      Height          =   500
      Left            =   2000
      TabIndex        =   3
      Top             =   3000
      Width           =   500
   End
   Begin VB.CommandButton btn7 
      Enabled         =   0   'False
      Height          =   500
      Left            =   3000
      TabIndex        =   2
      Top             =   1000
      Width           =   500
   End
   Begin VB.CommandButton btn8 
      Enabled         =   0   'False
      Height          =   500
      Left            =   3000
      TabIndex        =   1
      Top             =   2000
      Width           =   500
   End
   Begin VB.CommandButton btn9 
      Enabled         =   0   'False
      Height          =   500
      Left            =   3000
      TabIndex        =   0
      Top             =   3000
      Width           =   500
   End
   Begin VB.Label lblTurn 
      Height          =   492
      Left            =   960
      TabIndex        =   11
      Top             =   360
      Width           =   2532
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CountValue As String

Private Sub btn1_Click()
    If CountValue = "1" Then
        btn1.Caption = "X"
        btn1.Enabled = False
        SwitchValue (CountValue)
        Status ("X")
    Else
        btn1.Caption = "O"
        btn1.Enabled = False
        SwitchValue (CountValue)
        Status ("O")
    End If
End Sub

Private Sub btn2_Click()
    If CountValue = "1" Then
        btn2.Caption = "X"
        btn2.Enabled = False
        SwitchValue (CountValue)
        Status ("X")
    Else
        btn2.Caption = "O"
        btn2.Enabled = False
        SwitchValue (CountValue)
        Status ("O")
    End If
End Sub

Private Sub btn3_Click()
    If CountValue = "1" Then
        btn3.Caption = "X"
        btn3.Enabled = False
        SwitchValue (CountValue)
        Status ("X")
    Else
        btn3.Caption = "O"
        btn3.Enabled = False
        SwitchValue (CountValue)
        Status ("O")
    End If
End Sub

Private Sub btn4_Click()
    If CountValue = "1" Then
        btn4.Caption = "X"
        btn4.Enabled = False
        SwitchValue (CountValue)
        Status ("X")
    Else
        btn4.Caption = "O"
        btn4.Enabled = False
        SwitchValue (CountValue)
        Status ("O")
    End If
End Sub

Private Sub btn5_Click()
    If CountValue = "1" Then
        btn5.Caption = "X"
        btn5.Enabled = False
        SwitchValue (CountValue)
        Status ("X")
    Else
        btn5.Caption = "O"
        btn5.Enabled = False
        SwitchValue (CountValue)
        Status ("O")
    End If
End Sub

Private Sub btn6_Click()
    If CountValue = "1" Then
        btn6.Caption = "X"
        btn6.Enabled = False
        SwitchValue (CountValue)
        Status ("X")
    Else
        btn6.Caption = "O"
        btn6.Enabled = False
        SwitchValue (CountValue)
        Status ("O")
    End If
End Sub

Private Sub btn7_Click()
    If CountValue = "1" Then
        btn7.Caption = "X"
        btn7.Enabled = False
        SwitchValue (CountValue)
        Status ("X")
    Else
        btn7.Caption = "O"
        btn7.Enabled = False
        SwitchValue (CountValue)
        Status ("O")
    End If
End Sub

Private Sub btn8_Click()
    If CountValue = "1" Then
        btn8.Caption = "X"
        btn8.Enabled = False
        SwitchValue (CountValue)
        Status ("X")
    Else
        btn8.Caption = "O"
        btn8.Enabled = False
        SwitchValue (CountValue)
        Status ("O")
    End If
End Sub

Private Sub btn9_Click()
    If CountValue = "1" Then
        btn9.Caption = "X"
        btn9.Enabled = False
        SwitchValue (CountValue)
        Status ("X")
    Else
        btn9.Caption = "O"
        btn9.Enabled = False
        SwitchValue (CountValue)
        Status ("O")
    End If
End Sub

Private Sub btnReset_Click()
btn1.Caption = ""
btn2.Caption = ""
btn3.Caption = ""
btn4.Caption = ""
btn5.Caption = ""
btn6.Caption = ""
btn7.Caption = ""
btn8.Caption = ""
btn9.Caption = ""
lblTurn.Caption = " "
EnableButton
CountValue = "1"
End Sub

Private Sub btnStart_Click()
EnableButton
'Static CountValue As String
CountValue = "1"
lblTurn.Caption = "X Turn"
End Sub

Private Function SwitchValue(count As Integer)
    If count = "1" Then
        CountValue = "0"
        lblTurn.Caption = "O Turn"
    Else
        CountValue = "1"
        lblTurn.Caption = "X Turn"
    End If
End Function
Private Sub EnableButton()
btn1.Enabled = True
btn2.Enabled = True
btn3.Enabled = True
btn4.Enabled = True
btn5.Enabled = True
btn6.Enabled = True
btn7.Enabled = True
btn8.Enabled = True
btn9.Enabled = True
End Sub
'Private Sub DisableButton()
'btn1.Enabled = False
'btn2.Enabled = False
'btn3.Enabled = False
'btn4.Enabled = False
'btn5.Enabled = False
'btn6.Enabled = False
'btn7.Enabled = False
'btn8.Enabled = False
'btn9.Enabled = False
'End Sub

Private Function Status(char As String)

If (btn1.Caption = char And btn2.Caption = char And btn3.Caption = char) Then
    MsgBox (char + " is the Winner")
    btnReset_Click
ElseIf (btn1.Caption = char And btn4.Caption = char And btn7.Caption = char) Then
    MsgBox (char + " is the Winner")
    btnReset_Click
ElseIf (btn1.Caption = char And btn5.Caption = char And btn9.Caption = char) Then
    MsgBox (char + " is the Winner")
    btnReset_Click
ElseIf (btn4.Caption = char And btn5.Caption = char And btn6.Caption = char) Then
    MsgBox (char + " is the Winner")
    btnReset_Click
ElseIf (btn7.Caption = char And btn8.Caption = char And btn9.Caption = char) Then
    MsgBox (char + " is the Winner")
    btnReset_Click
ElseIf (btn2.Caption = char And btn5.Caption = char And btn8.Caption = char) Then
    MsgBox (char + " is the Winner")
    btnReset_Click
ElseIf (btn3.Caption = char And btn6.Caption = char And btn9.Caption = char) Then
    MsgBox (char + " is the Winner")
    btnReset_Click
ElseIf (btn3.Caption = char And btn5.Caption = char And btn7.Caption = char) Then
    MsgBox (char + " is the Winner")
    btnReset_Click
'ElseIf (btn1.Caption = "O" And btn2.Caption = "O" And btn3.Caption = "O") Then
'    MsgBox ("O is the Winner")
'    btnReset_Click
'ElseIf (btn1.Caption = "O" And btn4.Caption = "O" And btn7.Caption = "O") Then
'    MsgBox ("O is the Winner")
'    btnReset_Click
'ElseIf (btn1.Caption = "O" And btn5.Caption = "O" And btn9.Caption = "O") Then
'    MsgBox ("O is the Winner")
'    btnReset_Click
'ElseIf (btn4.Caption = "O" And btn5.Caption = "O" And btn6.Caption = "O") Then
'    MsgBox ("O is the Winner")
'    btnReset_Click
'ElseIf (btn7.Caption = "O" And btn8.Caption = "O" And btn9.Caption = "O") Then
'    MsgBox ("O is the Winner")
'    btnReset_Click
'ElseIf (btn2.Caption = "O" And btn5.Caption = "O" And btn8.Caption = "O") Then
'    MsgBox ("O is the Winner")
'    btnReset_Click
'ElseIf (btn3.Caption = "O" And btn6.Caption = "O" And btn9.Caption = "O") Then
'    MsgBox ("O is the Winner")
'    btnReset_Click
'ElseIf (btn3.Caption = "O" And btn5.Caption = "O" And btn7.Caption = "O") Then
'    MsgBox ("O is the Winner")
'    btnReset_Click
ElseIf (btn1.Enabled = False And btn2.Enabled = False And btn3.Enabled = False And btn4.Enabled = False And btn5.Enabled = False And btn6.Enabled = False And btn7.Enabled = False And btn8.Enabled = False And btn9.Enabled = False) Then
    MsgBox ("It's a Draw")
    btnReset_Click
End If

End Function
