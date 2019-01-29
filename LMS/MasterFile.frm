VERSION 5.00
Begin VB.MDIForm MasterFile 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Sunnydale Library Management System-[Books Master File]"
   ClientHeight    =   8130
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   8760
   Icon            =   "MasterFile.frx":0000
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu frmAddBook 
         Caption         =   "Add Book"
      End
      Begin VB.Menu frmAddStudent 
         Caption         =   "Add Student"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Transaction 
      Caption         =   "Transaction"
      Begin VB.Menu frmBookIssue 
         Caption         =   "Issue Book"
      End
      Begin VB.Menu frmReturnBook 
         Caption         =   "Return Book"
      End
   End
   Begin VB.Menu Records 
      Caption         =   "Records"
      Begin VB.Menu mbrDisplay 
         Caption         =   "Member Display"
      End
      Begin VB.Menu bkDisplay 
         Caption         =   "Book Display"
      End
   End
   Begin VB.Menu frmSettings 
      Caption         =   "Settings"
   End
End
Attribute VB_Name = "MasterFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bkDisplay_Click()
BookDisplay.Show
End Sub

Private Sub Exit_Click()
Unload Me
End Sub

Private Sub frmAddBook_Click()
AddRecordBook.Show
'Unload AddRecordStudent
End Sub

Private Sub frmAddStudent_Click()
'Unload AddRecordBook
AddRecordStudent.Show
End Sub

Private Sub frmBookIssue_Click()
BookIssueForm.Show
End Sub

Private Sub frmReturnBook_Click()
BookReturnForm.Show
End Sub

Private Sub frmSettings_Click()
Settings.Show
End Sub

Private Sub mbrDisplay_Click()
StudentDisplay.Show
End Sub

Private Sub MDIForm_Load()
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
cmd.CommandText = "usp_DisplaySetting"
cmd.Parameters.Append cmd.CreateParameter("@id", adInteger, adParamInput, , 1)
Set rs = cmd.Execute

'initialising global variables
numDays = rs.Fields(1)
fine = rs.Fields(2)
rs.Close
conn.Close
Set conn = Nothing
Set cmd = Nothing
Set rs = Nothing

End Function

