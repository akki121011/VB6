VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form BookIssueForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Book Issue Form"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4905
   Icon            =   "BookIssueForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker dpReturn 
      Height          =   300
      Left            =   1800
      TabIndex        =   15
      Top             =   3200
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   98566144
      CurrentDate     =   43321
   End
   Begin MSComCtl2.DTPicker dpIssue 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "M/d/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   5
      Top             =   2800
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman Greek"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   98566144
      CurrentDate     =   43321
      MaxDate         =   44196
      MinDate         =   43321
   End
   Begin VB.CommandButton btnIssue 
      Caption         =   "Issue Book"
      Height          =   400
      Left            =   100
      TabIndex        =   6
      Top             =   3800
      Width           =   1500
   End
   Begin VB.CommandButton btnReset 
      Caption         =   "Reset"
      Height          =   400
      Left            =   1700
      TabIndex        =   14
      Top             =   3800
      Width           =   1500
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   400
      Left            =   3300
      TabIndex        =   13
      Top             =   3800
      Width           =   1500
   End
   Begin VB.TextBox txtStudCode 
      Height          =   300
      Left            =   1800
      TabIndex        =   1
      Top             =   1200
      Width           =   2500
   End
   Begin VB.TextBox txtStudName 
      Height          =   300
      Left            =   1800
      TabIndex        =   2
      Top             =   1600
      Width           =   3000
   End
   Begin VB.TextBox txtBookId 
      Height          =   300
      Left            =   1800
      TabIndex        =   3
      Top             =   2000
      Width           =   2500
   End
   Begin VB.TextBox txtBookTitle 
      Height          =   300
      Left            =   1800
      TabIndex        =   4
      Top             =   2400
      Width           =   3000
   End
   Begin VB.Line Line2 
      X1              =   150
      X2              =   4800
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label7 
      Caption         =   "Student Code :"
      Height          =   300
      Left            =   200
      TabIndex        =   12
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Label Label6 
      Caption         =   "Student Name :"
      Height          =   300
      Left            =   200
      TabIndex        =   11
      Top             =   1600
      Width           =   1500
   End
   Begin VB.Label Label5 
      Caption         =   "Book Id :"
      Height          =   300
      Left            =   200
      TabIndex        =   10
      Top             =   2000
      Width           =   1500
   End
   Begin VB.Label Label4 
      Caption         =   "Book Title :"
      Height          =   300
      Left            =   200
      TabIndex        =   9
      Top             =   2400
      Width           =   1500
   End
   Begin VB.Label Label3 
      Caption         =   "Date Issued :"
      Height          =   300
      Left            =   200
      TabIndex        =   8
      Top             =   2800
      Width           =   1500
   End
   Begin VB.Label Label2 
      Caption         =   "Date to be returned :"
      Height          =   300
      Left            =   200
      TabIndex        =   7
      Top             =   3200
      Width           =   1500
   End
   Begin VB.Image Image1 
      Height          =   500
      Left            =   250
      Picture         =   "BookIssueForm.frx":133CC
      Stretch         =   -1  'True
      Top             =   200
      Width           =   500
   End
   Begin VB.Label Label1 
      Caption         =   $"BookIssueForm.frx":15EDC
      Height          =   900
      Left            =   1000
      TabIndex        =   0
      Top             =   50
      Width           =   3800
   End
   Begin VB.Line Line1 
      DrawMode        =   1  'Blackness
      X1              =   150
      X2              =   4800
      Y1              =   1000
      Y2              =   1000
   End
End
Attribute VB_Name = "BookIssueForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
Unload BookIssueForm
End Sub

Private Sub btnIssue_Click()

   On Error GoTo Error1
    Dim conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    conn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=NGSS;Initial Catalog=TrainingDB;Data Source=."
    conn.Open
    
    'Done with Stored procedure
    'On Error GoTo Error1
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "usp_IssueBook"
    
    cmd.Parameters.Append cmd.CreateParameter("StudentId", adInteger, adParamInput, , Val(txtStudCode.Text))
    cmd.Parameters.Append cmd.CreateParameter("StudentName", adVarChar, adParamInput, 50, txtStudName.Text)
    cmd.Parameters.Append cmd.CreateParameter("BookId", adInteger, adParamInput, , Val(txtBookId.Text))
    cmd.Parameters.Append cmd.CreateParameter("BookTitle", adVarChar, adParamInput, 100, txtBookTitle.Text)
    cmd.Parameters.Append cmd.CreateParameter("DateIssued", adDate, adParamInput, , dpIssue.Value)
    'Here we pass return date 14 after the issue date
    'This process is now automated
    cmd.Parameters.Append cmd.CreateParameter("DateReturn", adDate, adParamInput, , (dpIssue.Value + numDays))
    
    Set rs = cmd.Execute
    MsgBox "The record was added to database", vbInformation, "Record Added Successfully"
    dpReturn.Value = dpIssue + numDays
    conn.Close
    
    Set conn = Nothing
    Set cmd = Nothing
    Set rs = Nothing
    
   Exit Sub
Error1:
   MsgBox Err.Description, vbExclamation

End Sub

Private Sub btnReset_Click()
txtStudCode.Text = ""
txtStudName.Text = ""
txtBookId.Text = ""
txtBookTitle.Text = ""
End Sub

Private Sub Form_Load()
Move 0, 0
dpIssue.MinDate = DateTime.Now
Dim today As Date
today = DateTime.Now
dpReturn.Value = today + numDays
End Sub
