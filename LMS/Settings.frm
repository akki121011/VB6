VERSION 5.00
Begin VB.Form Settings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnSave 
      Caption         =   "Save and Close"
      Height          =   400
      Left            =   1560
      TabIndex        =   3
      Top             =   2800
      Width           =   1600
   End
   Begin VB.TextBox txtFine 
      Height          =   300
      Left            =   200
      TabIndex        =   2
      Top             =   2200
      Width           =   4600
   End
   Begin VB.TextBox txtDay 
      Height          =   300
      Left            =   200
      TabIndex        =   1
      Top             =   700
      Width           =   4600
   End
   Begin VB.Label Label2 
      Caption         =   "What are the maximum number of days a book can be kept before fines are generated?"
      Height          =   500
      Left            =   200
      TabIndex        =   4
      Top             =   200
      Width           =   4600
   End
   Begin VB.Label Label1 
      Caption         =   "What is the amount of fines enforced if a book is not returned on time?"
      Height          =   500
      Left            =   200
      TabIndex        =   0
      Top             =   1700
      Width           =   4600
   End
End
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSave_Click()
Dim conn As ADODB.Connection
Dim cmd As ADODB.Command
Dim rs As Recordset
Dim connString As String
connString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=NGSS;Initial Catalog=TrainingDB;Data Source=."
Set conn = New Connection
conn.ConnectionString = connString
conn.Open
'Done with Stored procedure
Set cmd = New ADODB.Command
cmd.ActiveConnection = conn
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "usp_updateSetting"
cmd.Parameters.Append cmd.CreateParameter("@id", adInteger, adParamInput, , 1)
cmd.Parameters.Append cmd.CreateParameter("@days", adInteger, adParamInput, , numDays)
cmd.Parameters.Append cmd.CreateParameter("@fine", adInteger, adParamInput, , fine)

Set rs = cmd.Execute

conn.Close
Set conn = Nothing
Set cmd = Nothing
Set rs = Nothing

Call Display

End Sub

Private Sub Form_Load()
Move 0, 0
Call Display
txtDay.Text = numDays
txtFine.Text = fine
End Sub

Private Function Display()
Dim conn As ADODB.Connection
Dim cmd As ADODB.Command
Dim rs As Recordset
Dim connString As String
connString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=NGSS;Initial Catalog=TrainingDB;Data Source=."
Set conn = New Connection
conn.ConnectionString = connString
conn.Open
'Done with Stored procedure
Set cmd = New ADODB.Command
cmd.ActiveConnection = conn
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "usp_DisplaySetting"
cmd.Parameters.Append cmd.CreateParameter("@id", adInteger, adParamInput, , 1)
Set rs = cmd.Execute
numDays = rs.Fields(1)
fine = rs.Fields(2)
rs.Close
conn.Close
Set conn = Nothing
Set cmd = Nothing
Set rs = Nothing

End Function

