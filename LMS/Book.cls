VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Book"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim bookId As Integer
Dim title As String
Dim author As String
Dim publisher As String
Dim categ As String
Dim price As Integer
Dim isbn As String
Dim borrow As Boolean

Public Property Get BookIds() As Integer
    BookIds = bookId
End Property

Public Property Let BookIds(ByVal var As Integer)
    bookId = var
End Property
Public Property Get Prices() As Integer
    Prices = price
End Property

Public Property Let Prices(ByVal var As Integer)
    price = var
End Property
Public Property Get Titles() As String
    Titles = title
End Property

Public Property Let Titles(ByVal var As String)
    title = var
End Property
Public Property Get Authors() As String
    Authors = author
End Property

Public Property Let Authors(ByVal var As String)
    author = var
End Property
Public Property Get Publishers() As String
    Publishers = publisher
End Property

Public Property Let Publishers(ByVal var As String)
    publisher = var
End Property
Public Property Get Category() As String
    Category = categ
End Property

Public Property Let Category(ByVal var As String)
    categ = var
End Property
Public Property Get ISBNS() As String
    ISBNS = isbn
End Property

Public Property Let ISBNS(ByVal var As String)
    isbn = var
End Property
Public Property Get Borrows() As Boolean
    Borrows = borrow
End Property

Public Property Let Borrows(ByVal var As Boolean)
    borrow = var
End Property

