VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Increment"
   ClientHeight    =   3645
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   11130
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMax 
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txtVal 
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   2040
      Width           =   1455
   End
   Begin VB.OptionButton rdbDivision 
      Caption         =   "Division"
      Height          =   495
      Left            =   6480
      TabIndex        =   7
      Top             =   1320
      Width           =   1500
   End
   Begin VB.OptionButton rdbIncrement 
      Caption         =   "Increment"
      Height          =   495
      Left            =   6480
      TabIndex        =   6
      Top             =   360
      Value           =   -1  'True
      Width           =   1500
   End
   Begin VB.TextBox txtMin 
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox txtDisplay 
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      Locked          =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   3
      Top             =   2760
      Width           =   5055
   End
   Begin VB.CommandButton btnCalc 
      Caption         =   "Calculate"
      Height          =   615
      Left            =   8760
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lblValue 
      Caption         =   "Enter Value :"
      Height          =   615
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label lblMin 
      Caption         =   "Enter Minimum Value :"
      Height          =   615
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label lblMin 
      Caption         =   "Enter Maximum Value :"
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub btnCalc_Click()

Dim min As Integer
min = val(txtMin.Text)

Dim max As Integer
max = val(txtMax.Text)

Dim values As Integer
values = val(txtVal.Text)

Dim index As Integer

Dim sb As String
sb = ""

If rdbIncrement.Value = True Then
For index = min To max Step index + values
'sb.Append("" {0}",&index
sb = sb & " " & index

txtDisplay.Text = sb
Next

Else
For index = min To max Step index + 1
If index Mod values = 0 Then
sb = sb & " " & index
txtDisplay.Text = sb
End If

Next index

End If

End Sub
