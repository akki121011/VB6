VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form StudentDisplay 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "StudentDisplay"
   ClientHeight    =   4425
   ClientLeft      =   1050
   ClientTop       =   1380
   ClientWidth     =   7755
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgDisplay 
      Height          =   3000
      Left            =   200
      TabIndex        =   0
      Top             =   200
      Width           =   7000
      _ExtentX        =   12356
      _ExtentY        =   5292
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "StudentDisplay"
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
 cmd.CommandText = "usp_MemberDisplay"
 Set rs = cmd.Execute

Set dgDisplay.DataSource = rs

rs.Close
conn.Close

Set conn = Nothing
Set cmd = Nothing
Set rs = Nothing

End Function

