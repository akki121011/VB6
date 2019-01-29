VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form BookDisplay 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Book Details"
   ClientHeight    =   4995
   ClientLeft      =   1050
   ClientTop       =   1380
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgDisplay 
      Height          =   4000
      Left            =   200
      TabIndex        =   0
      Top             =   200
      Width           =   7000
      _ExtentX        =   12356
      _ExtentY        =   7064
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "BookDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Move 0, 0
Call Display
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
 cmd.CommandText = "usp_BookDisplay"
 Set rs = cmd.Execute

''Setting Style to coloumn
'dgDisplay.ColWidth(0) = 1000
'dgDisplay.ColWidth(1) = 1500
'dgDisplay.ColWidth(2) = 1800
'dgDisplay.ColWidth(3) = 1940

Set dgDisplay.DataSource = rs

rs.Close
conn.Close

Set conn = Nothing
Set cmd = Nothing
Set rs = Nothing

End Function
