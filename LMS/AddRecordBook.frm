VERSION 5.00
Begin VB.Form AddRecordBook 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Record"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPrice 
      Height          =   300
      Left            =   1100
      TabIndex        =   5
      Top             =   3200
      Width           =   3000
   End
   Begin VB.TextBox txtBookId 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1100
      TabIndex        =   0
      Top             =   1200
      Width           =   3000
   End
   Begin VB.TextBox txtTitle 
      Height          =   300
      Left            =   1100
      TabIndex        =   1
      Top             =   1600
      Width           =   3000
   End
   Begin VB.TextBox txtAuthor 
      Height          =   300
      Left            =   1100
      TabIndex        =   2
      Top             =   2000
      Width           =   2500
   End
   Begin VB.TextBox txtPublisher 
      Height          =   300
      Left            =   1100
      TabIndex        =   3
      Top             =   2400
      Width           =   3000
   End
   Begin VB.ComboBox cbxCategory 
      Height          =   315
      Left            =   1100
      TabIndex        =   4
      Text            =   "N/A"
      Top             =   2800
      Width           =   3000
   End
   Begin VB.TextBox txtIsbn 
      Height          =   300
      Left            =   1100
      TabIndex        =   6
      Top             =   3600
      Width           =   3000
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   400
      Left            =   3240
      TabIndex        =   9
      Top             =   5200
      Width           =   1000
   End
   Begin VB.CommandButton btnReset 
      Caption         =   "Reset"
      Height          =   400
      Left            =   2115
      TabIndex        =   7
      Top             =   5200
      Width           =   1000
   End
   Begin VB.CommandButton btnAdd 
      Caption         =   "Add"
      Height          =   400
      Left            =   1000
      TabIndex        =   8
      Top             =   5200
      Width           =   1000
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   105
      Index           =   0
      Left            =   3600
      TabIndex        =   21
      Top             =   2040
      Width           =   105
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "*"
      ForeColor       =   &H000000FF&
      Height          =   100
      Index           =   1
      Left            =   4100
      TabIndex        =   20
      Top             =   1600
      Width           =   100
   End
   Begin VB.Label Label10 
      Caption         =   $"AddRecordBook.frx":0000
      Enabled         =   0   'False
      Height          =   1000
      Left            =   1100
      TabIndex        =   19
      Top             =   4000
      Width           =   3000
   End
   Begin VB.Line Line1 
      DrawMode        =   1  'Blackness
      X1              =   120
      X2              =   4250
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label1 
      Caption         =   $"AddRecordBook.frx":009F
      Height          =   855
      Left            =   840
      TabIndex        =   18
      Top             =   120
      Width           =   3500
   End
   Begin VB.Image Image1 
      Height          =   500
      Left            =   120
      Picture         =   "AddRecordBook.frx":0160
      Stretch         =   -1  'True
      Top             =   240
      Width           =   500
   End
   Begin VB.Label Label2 
      Caption         =   "ISBN :"
      Height          =   300
      Left            =   100
      TabIndex        =   17
      Top             =   3600
      Width           =   1000
   End
   Begin VB.Label Label3 
      Caption         =   "Price :"
      Height          =   300
      Left            =   100
      TabIndex        =   16
      Top             =   3200
      Width           =   1000
   End
   Begin VB.Label Label4 
      Caption         =   "Category :"
      Height          =   300
      Left            =   100
      TabIndex        =   15
      Top             =   2800
      Width           =   1000
   End
   Begin VB.Label Label5 
      Caption         =   "Publisher :"
      Height          =   300
      Left            =   100
      TabIndex        =   14
      Top             =   2400
      Width           =   1000
   End
   Begin VB.Label lblAuthor 
      Caption         =   "Author :"
      Height          =   300
      Left            =   100
      TabIndex        =   13
      Top             =   2000
      Width           =   1000
   End
   Begin VB.Label Label7 
      Caption         =   "Title :"
      Height          =   300
      Left            =   100
      TabIndex        =   12
      Top             =   1600
      Width           =   1000
   End
   Begin VB.Label Label8 
      Caption         =   "Book ID :"
      Height          =   300
      Left            =   100
      TabIndex        =   11
      Top             =   1200
      Width           =   1000
   End
   Begin VB.Label Label9 
      Caption         =   "Borrowed:"
      Height          =   300
      Left            =   100
      TabIndex        =   10
      Top             =   4000
      Width           =   1000
   End
   Begin VB.Line Line2 
      DrawMode        =   1  'Blackness
      X1              =   0
      X2              =   4650
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   4250
      Y1              =   5050
      Y2              =   5050
   End
End
Attribute VB_Name = "AddRecordBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAdd_Click()
Dim obj As New Book
obj.Authors = txtAuthor.Text
obj.BookIds = Val(txtBookId.Text)
obj.Category = cbxCategory.Text
obj.ISBNS = txtIsbn.Text
obj.Prices = Val(txtPrice.Text)
obj.Publishers = txtPublisher.Text
obj.Titles = txtTitle.Text

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
    cmd.CommandText = "usp_AddBook"
    
    cmd.Parameters.Append cmd.CreateParameter("Title", adVarChar, adParamInput, 20, obj.Titles)
    cmd.Parameters.Append cmd.CreateParameter("Author", adVarChar, adParamInput, 50, obj.Authors)
    cmd.Parameters.Append cmd.CreateParameter("PublisherName", adVarChar, adParamInput, 50, obj.Publishers)
    cmd.Parameters.Append cmd.CreateParameter("Category", adVarChar, adParamInput, 20, obj.Category)
    cmd.Parameters.Append cmd.CreateParameter("Price", adInteger, adParamInput, , obj.Prices)
    cmd.Parameters.Append cmd.CreateParameter("ISBN", adVarChar, adParamInput, 20, obj.ISBNS)
    cmd.Parameters.Append cmd.CreateParameter("Borrowed", adBoolean, adParamInput, , obj.Borrows)
    
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
End Sub

Private Sub btnCancel_Click()
Unload AddRecordBook
End Sub

Private Sub btnReset_Click()
txtBookId.Text = ""
txtAuthor.Text = ""
txtIsbn.Text = ""
txtPrice.Text = ""
txtPublisher.Text = ""
txtTitle.Text = ""
cbxCategory.Text = "N/A"
End Sub

Private Function Validate(obj As Book) As Boolean
Dim status As Boolean
status = True
If obj.Authors = "" Or obj.ISBNS = "" Or obj.BookIds Or obj.Category = "N/A" Or obj.Prices = Null Or obj.Publishers = "" Or obj.Titles = "" Then
    MsgBox "Fields cant't be empty", vbExclamation
    status = False
End If
If status Then
    Validate = True
End If
End Function

Private Sub Form_Load()
Move 0, 0
cbxCategory.AddItem "Fiction"
cbxCategory.AddItem "Short Novel"
cbxCategory.AddItem "Guide"
cbxCategory.AddItem "Thrill"
End Sub
