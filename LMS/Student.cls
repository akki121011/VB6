VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Student"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim studId As Integer
Dim firstName As String
Dim middleName As String
Dim lastName As String
Dim ClassValue As String
Dim section As String
Dim roll As String

Public Property Get StudIds() As Integer
    StudIds = studId
End Property

Public Property Let StudIds(ByVal var As Integer)
    studId = var
End Property
Public Property Get FirstNames() As String
    FirstNames = firstName
End Property

Public Property Let FirstNames(ByVal var As String)
    firstName = var
End Property
Public Property Get MiddleNames() As String
    MiddleNames = middleName
End Property

Public Property Let MiddleNames(ByVal var As String)
    middleName = var
End Property
Public Property Get LastNames() As String
    LastNames = lastName
End Property

Public Property Let LastNames(ByVal var As String)
    lastName = var
End Property
Public Property Get Classes() As String
    Classes = ClassValue
End Property

Public Property Let Classes(ByVal var As String)
    ClassValue = var
End Property
Public Property Get Sections() As String
    Sections = section
End Property

Public Property Let Sections(ByVal var As String)
    section = var
End Property
Public Property Get Rolls() As String
    Rolls = roll
End Property

Public Property Let Rolls(ByVal var As String)
    roll = var
End Property
