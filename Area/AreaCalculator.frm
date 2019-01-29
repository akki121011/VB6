VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5592
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   13860
   LinkTopic       =   "Form1"
   ScaleHeight     =   5592
   ScaleWidth      =   13860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCalc 
      Caption         =   "Calculate"
      Height          =   500
      Left            =   6000
      TabIndex        =   13
      Top             =   1500
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox txtLen 
      Height          =   495
      Left            =   4000
      TabIndex        =   12
      Top             =   1000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtBrd 
      Height          =   495
      Left            =   4000
      TabIndex        =   11
      Top             =   2000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtRad 
      Height          =   495
      Left            =   4000
      TabIndex        =   10
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.OptionButton rdbSph 
      Caption         =   "Sphere Area"
      Height          =   495
      Left            =   500
      TabIndex        =   4
      Top             =   500
      Width           =   1215
   End
   Begin VB.OptionButton rdbTri 
      Caption         =   "Triangle Area"
      Height          =   495
      Left            =   500
      TabIndex        =   3
      Top             =   1500
      Width           =   1215
   End
   Begin VB.OptionButton rdbSqr 
      Caption         =   "Square Area"
      Height          =   495
      Left            =   500
      TabIndex        =   2
      Top             =   2500
      Width           =   1215
   End
   Begin VB.OptionButton rdbCir 
      Caption         =   "Circle Area"
      Height          =   495
      Left            =   500
      TabIndex        =   1
      Top             =   3500
      Width           =   1215
   End
   Begin VB.OptionButton rdbRect 
      Caption         =   "Rectangle Area"
      Height          =   495
      Left            =   500
      TabIndex        =   0
      Top             =   4500
      Width           =   1215
   End
   Begin VB.Label Label3 
      Height          =   495
      Left            =   2000
      TabIndex        =   9
      Top             =   1000
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   2000
      TabIndex        =   8
      Top             =   2000
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   2000
      TabIndex        =   7
      Top             =   3000
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label lblArea 
      Height          =   495
      Left            =   8500
      TabIndex        =   6
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblAr 
      Caption         =   "The Desired Area is :"
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   2000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCalc_Click()

btnCalc.Visible = False
Label1.Visible = False
txtRad.Visible = False
btnCalc.Visible = False
Label2.Visible = False
Label3.Visible = False
txtLen.Visible = False
txtBrd.Visible = False
lblAr.Visible = True
lblArea.Visible = True

'
'Dim pi As Double
'pi = 3.14

Dim ar As Double



If rdbCir.Value = True Then
    ar = CircleArea(Val(txtRad.Text))
    lblArea.Caption = ar
ElseIf rdbRect.Value = True Then
    ar = RectangleArea(Val(txtLen.Text), Val(txtBrd.Text))
    lblArea.Caption = ar
ElseIf rdbSph.Value = True Then
    ar = SphereArea(Val(txtRad.Text))
    lblArea.Caption = ar
ElseIf rdbSqr.Value = True Then
    ar = SquareArea(Val(txtLen.Text))
    lblArea.Caption = ar
ElseIf rdbTri.Value = True Then
    ar = TriangleArea(Val(txtLen.Text), Val(txtBrd.Text))
    lblArea.Caption = ar
End If

End Sub

Private Sub rdbCir_Click()

Label1.Visible = True
txtRad.Visible = True
btnCalc.Visible = True
lblAr.Visible = False
lblArea.Visible = False

Label2.Visible = False
Label3.Visible = False
txtLen.Visible = False
txtBrd.Visible = False

Label1.Caption = "Enter the radius :"



End Sub

Private Sub rdbRect_Click()
Label2.Visible = True
Label3.Visible = True
txtLen.Visible = True
txtBrd.Visible = True
lblAr.Visible = False
lblArea.Visible = False

Label1.Visible = False
txtRad.Visible = False
btnCalc.Visible = False

btnCalc.Visible = True
Label3.Caption = "Enter the Length :"
Label2.Caption = "Enter the Breadth :"

End Sub

Private Sub rdbSph_Click()
Label1.Visible = True
txtRad.Visible = True
btnCalc.Visible = True
lblAr.Visible = False
lblArea.Visible = False
Label2.Visible = False
Label3.Visible = False
txtLen.Visible = False
txtBrd.Visible = False

Label1.Caption = "Enter the radius :"

End Sub

Private Sub rdbSqr_Click()
Label2.Visible = False
Label3.Visible = True
txtLen.Visible = True
txtBrd.Visible = False
lblAr.Visible = False
lblArea.Visible = False
Label1.Visible = False
txtRad.Visible = False
btnCalc.Visible = False

btnCalc.Visible = True
Label3.Caption = "Enter the Side of Square:"

End Sub

Private Sub rdbTri_Click()

Label2.Visible = True
Label3.Visible = True
txtLen.Visible = True
txtBrd.Visible = True
lblAr.Visible = False
lblArea.Visible = False
Label1.Visible = False
txtRad.Visible = False
btnCalc.Visible = False

btnCalc.Visible = True
Label3.Caption = "Enter the Height :"
Label2.Caption = "Enter the Base :"
End Sub

Private Function CircleArea(dblRad As Double) As Double

CircleArea = pi * dblRad * dblRad

End Function
Private Function TriangleArea(base, length As Double)
TriangleArea = 0.5 * base * length

End Function

Private Function RectangleArea(length, breadth As Double)
RectangleArea = length * breadth

End Function


Private Function SphereArea(dblRad As Double) As Double

SphereArea = 4 * pi * dblRad * dblRad

End Function

Private Function SquareArea(side As Double)
SquareArea = side * side

End Function
