VERSION 5.00
Begin VB.Form AddRecordStudent 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Record"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnAdd 
      Caption         =   "Add"
      Height          =   400
      Left            =   1000
      TabIndex        =   8
      Top             =   4245
      Width           =   1000
   End
   Begin VB.CommandButton btnReset 
      Caption         =   "Reset"
      Height          =   400
      Left            =   2115
      TabIndex        =   17
      Top             =   4245
      Width           =   1000
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   400
      Left            =   3240
      TabIndex        =   16
      Top             =   4245
      Width           =   1000
   End
   Begin VB.TextBox txtRoll 
      Height          =   300
      Left            =   1100
      TabIndex        =   7
      Top             =   3600
      Width           =   1200
   End
   Begin VB.ComboBox cbxClass 
      Height          =   315
      Left            =   1100
      TabIndex        =   5
      Text            =   "-----------------------Select--------------------------"
      Top             =   2800
      Width           =   3000
   End
   Begin VB.ComboBox cbxSection 
      Height          =   315
      Left            =   1100
      TabIndex        =   6
      Text            =   "-----------------------Select--------------------------"
      Top             =   3200
      Width           =   3000
   End
   Begin VB.TextBox txtLastName 
      Height          =   300
      Left            =   1100
      TabIndex        =   4
      Top             =   2400
      Width           =   3000
   End
   Begin VB.TextBox txtMidInitial 
      Height          =   300
      Left            =   1100
      TabIndex        =   3
      Top             =   2000
      Width           =   400
   End
   Begin VB.TextBox txtFirstName 
      Height          =   300
      Left            =   1100
      TabIndex        =   2
      Top             =   1600
      Width           =   3000
   End
   Begin VB.TextBox txtStudId 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1100
      TabIndex        =   1
      Top             =   1200
      Width           =   2000
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   100
      Index           =   4
      Left            =   2300
      TabIndex        =   22
      Top             =   3600
      Width           =   100
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   100
      Index           =   3
      Left            =   4100
      TabIndex        =   21
      Top             =   2400
      Width           =   100
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   100
      Index           =   2
      Left            =   4100
      TabIndex        =   20
      Top             =   2800
      Width           =   100
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   100
      Index           =   1
      Left            =   4100
      TabIndex        =   19
      Top             =   3200
      Width           =   100
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   100
      Index           =   0
      Left            =   4100
      TabIndex        =   18
      Top             =   1600
      Width           =   100
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   4250
      Y1              =   4095
      Y2              =   4095
   End
   Begin VB.Line Line2 
      DrawMode        =   1  'Blackness
      X1              =   0
      X2              =   4650
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label8 
      Caption         =   "Student ID :"
      Height          =   300
      Left            =   100
      TabIndex        =   15
      Top             =   1200
      Width           =   1000
   End
   Begin VB.Label Label7 
      Caption         =   "First Name :"
      Height          =   300
      Left            =   100
      TabIndex        =   14
      Top             =   1600
      Width           =   1000
   End
   Begin VB.Label Label6 
      Caption         =   "Middle Initial :"
      Height          =   300
      Left            =   100
      TabIndex        =   13
      Top             =   2000
      Width           =   1000
   End
   Begin VB.Label Label5 
      Caption         =   "Last Name :"
      Height          =   300
      Left            =   100
      TabIndex        =   12
      Top             =   2400
      Width           =   1000
   End
   Begin VB.Label Label4 
      Caption         =   "Class :"
      Height          =   300
      Left            =   100
      TabIndex        =   11
      Top             =   2800
      Width           =   1000
   End
   Begin VB.Label Label3 
      Caption         =   "Section :"
      Height          =   300
      Left            =   100
      TabIndex        =   10
      Top             =   3200
      Width           =   1000
   End
   Begin VB.Label Label2 
      Caption         =   "Roll :"
      Height          =   300
      Left            =   100
      TabIndex        =   9
      Top             =   3600
      Width           =   1000
   End
   Begin VB.Image Image1 
      Height          =   500
      Left            =   120
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   500
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":1491
      Height          =   855
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   3500
   End
   Begin VB.Line Line1 
      DrawMode        =   1  'Blackness
      X1              =   120
      X2              =   4250
      Y1              =   1080
      Y2              =   1080
   End
End
Attribute VB_Name = "AddRecordStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------Select--------------------------
Private Sub btnAdd_Click()
Dim obj As New Student
obj.FirstNames = txtFirstName.Text
obj.MiddleNames = txtMidInitial.Text
obj.LastNames = txtLastName.Text
obj.Classes = cbxClass.Text
obj.Sections = cbxSection.Text
obj.Rolls = txtRoll.Text

If Validate(obj) Then
   On Error GoTo Error1
    Dim conn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
'    On Error GoTo Error1
    conn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Password=NGSS;Initial Catalog=TrainingDB;Data Source=."
    conn.Open
    
    'Done with Stored procedure
    'On Error GoTo Error1
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "usp_AddMember"

    cmd.Parameters.Append cmd.CreateParameter("FirstName", adVarChar, adParamInput, 20, obj.FirstNames)
    cmd.Parameters.Append cmd.CreateParameter("MiddleName", adVarChar, adParamInput, 5, obj.MiddleNames)
    cmd.Parameters.Append cmd.CreateParameter("LastName", adVarChar, adParamInput, 20, obj.LastNames)
    cmd.Parameters.Append cmd.CreateParameter("Class", adVarChar, adParamInput, 10, obj.Classes)
    cmd.Parameters.Append cmd.CreateParameter("Section", adVarChar, adParamInput, 10, obj.Sections)
    cmd.Parameters.Append cmd.CreateParameter("Roll", adVarChar, adParamInput, 10, obj.Rolls)
    
    Set rs = cmd.Execute
    MsgBox "The record was added to database", vbInformation, "Record Added Successfully"
    conn.Close
    
    Set conn = Nothing
    Set cmd = Nothing
    Set rs = Nothing
   Exit Sub
Error1:
   MsgBox Err.Description, vbExclamation
Else
    MsgBox "Enter a Valid Data", vbExclamation, "Invalid Data"
End If

'Error1:
'   MsgBox Err.Description, vbExclamation
End Sub

Private Sub btnCancel_Click()
Unload AddRecordStudent
End Sub

Private Sub btnReset_Click()
txtStudId.Text = ""
txtFirstName.Text = ""
txtMidInitial.Text = ""
txtLastName.Text = ""
cbxClass.Text = "-----------------------Select--------------------------"
cbxSection.Text = "-----------------------Select--------------------------"
txtRoll.Text = ""
End Sub

Private Sub Form_Load()
Move 0, 0
cbxClass.AddItem "I"
cbxClass.AddItem "II"
cbxClass.AddItem "III"
cbxClass.AddItem "IV"
cbxClass.AddItem "V"
cbxClass.AddItem "VI"
cbxClass.AddItem "VII"
cbxClass.AddItem "VIII"
cbxClass.AddItem "IX"
cbxClass.AddItem "X"
cbxClass.AddItem "XI"
cbxClass.AddItem "XII"
cbxSection.AddItem "A"
cbxSection.AddItem "B"
cbxSection.AddItem "C"
End Sub
Private Function Validate(obj As Student) As Boolean
Dim status As Boolean
status = True
If obj.FirstNames = "" Or obj.LastNames = "" Or cbxClass.Text = "-----------------------Select--------------------------" Or cbxSection.Text = "-----------------------Select--------------------------" Or obj.Rolls = "" Then
    MsgBox "Fields cant't be empty", vbExclamation
    status = False
End If
If status Then
    Validate = True
End If

End Function
