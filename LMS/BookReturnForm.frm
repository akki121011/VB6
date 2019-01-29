VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form BookReturnForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Book Return Form"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3531.915
   ScaleMode       =   0  'User
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker dpReturn 
      Height          =   317
      Left            =   1700
      TabIndex        =   3
      Top             =   1956
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   556
      _Version        =   393216
      Format          =   98566144
      CurrentDate     =   43321
   End
   Begin VB.CommandButton btnRtrn 
      Caption         =   "Return Book"
      Height          =   400
      Left            =   0
      TabIndex        =   4
      Top             =   3150
      Width           =   1500
   End
   Begin VB.CommandButton btnReset 
      Caption         =   "Reset"
      Height          =   400
      Left            =   1600
      TabIndex        =   11
      Top             =   3150
      Width           =   1500
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   400
      Left            =   3200
      TabIndex        =   10
      Top             =   3150
      Width           =   1500
   End
   Begin VB.TextBox txtBookId 
      Height          =   300
      Left            =   1700
      TabIndex        =   1
      Top             =   1150
      Width           =   2500
   End
   Begin VB.TextBox txtStudCode 
      Height          =   300
      Left            =   1700
      TabIndex        =   2
      Top             =   1550
      Width           =   3000
   End
   Begin VB.TextBox txtFine 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1700
      TabIndex        =   9
      Top             =   2350
      Width           =   3000
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   4300
      Picture         =   "BookReturnForm.frx":0000
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   300
   End
   Begin VB.Line Line2 
      X1              =   45
      X2              =   4695
      Y1              =   2794.326
      Y2              =   2794.326
   End
   Begin VB.Label Label7 
      Caption         =   "Book Id :"
      Height          =   300
      Left            =   100
      TabIndex        =   8
      Top             =   1150
      Width           =   1500
   End
   Begin VB.Label Label6 
      Caption         =   "Student Code :"
      Height          =   300
      Left            =   100
      TabIndex        =   7
      Top             =   1550
      Width           =   1500
   End
   Begin VB.Label Label5 
      Caption         =   "Date Returned :"
      Height          =   300
      Left            =   100
      TabIndex        =   6
      Top             =   1950
      Width           =   1500
   End
   Begin VB.Label Label4 
      Caption         =   "Fines Collected :"
      Height          =   300
      Left            =   100
      TabIndex        =   5
      Top             =   2350
      Width           =   1500
   End
   Begin VB.Image Image1 
      Height          =   500
      Left            =   150
      Picture         =   "BookReturnForm.frx":175C
      Stretch         =   -1  'True
      Top             =   150
      Width           =   500
   End
   Begin VB.Label Label1 
      Caption         =   $"BookReturnForm.frx":426C
      Height          =   900
      Left            =   900
      TabIndex        =   0
      Top             =   0
      Width           =   3800
   End
   Begin VB.Line Line1 
      DrawMode        =   1  'Blackness
      X1              =   50
      X2              =   4700
      Y1              =   898.345
      Y2              =   898.345
   End
End
Attribute VB_Name = "BookReturnForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCancel_Click()
Unload BookReturnForm
End Sub

Private Sub btnReset_Click()
txtBookId.Text = ""
txtStudCode.Text = ""
dpReturn.Value = DateTime.Now
txtFine.Text = ""
End Sub

Private Sub btnRtrn_Click()
Dim fineDays As Integer

If txtBookId.Text <> "" Or txtStudCode.Text <> "" Then
    'Receives no. of fineDays
    fineDays = RequiredFine(Val(txtBookId.Text), Val(txtStudCode.Text))
    
    'Hard coded calculation converted to automated calculation
    txtFine.Text = CStr(fineDays * fine) + "  INR"
    MsgBox "No. of extra Days : " + CStr(fineDays), vbExclamation, "Fine Days."

Else
    MsgBox "Please Enter the value in reuired fields.", vbExclamation, "Error."
End If
End Sub
Private Function RequiredFine(bookId, studId As Integer) As Integer
'usp_GetIssueDate
    Dim returnDate As Date
    Dim diffDate As Integer
    Dim conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    conn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=NGSS;Initial Catalog=TrainingDB;Data Source=."
    conn.Open
    
    'Done with Stored procedure
    'On Error GoTo Error1
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "usp_GetIssueDate"
    cmd.Parameters.Append cmd.CreateParameter("studID", adInteger, adParamInput, 20, studId)
    cmd.Parameters.Append cmd.CreateParameter("bookId", adInteger, adParamInput, 50, bookId)
    Set rs = cmd.Execute
    'Here we capture the return date from the database
    returnDate = rs.Fields(0)
    'Here we calculate the date difference
    diffDate = dpReturn - returnDate
    'If clauses can be added for further enhancement
    If diffDate > 0 Then
        RequiredFine = diffDate
    Else
        RequiredFine = 0
    End If
    
End Function

Private Sub Form_Load()
Move 0, 0
End Sub
