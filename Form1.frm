VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15660
   LinkTopic       =   "Form1"
   ScaleHeight     =   8940
   ScaleWidth      =   15660
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Forgot Password?"
      Height          =   375
      Left            =   13320
      TabIndex        =   4
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Register"
      Height          =   495
      Left            =   9240
      TabIndex        =   3
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   495
      Left            =   7320
      TabIndex        =   2
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   7320
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2640
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7320
      TabIndex        =   0
      Top             =   1800
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CONN As ADODB.Connection
Dim rs As ADODB.Recordset
Private Sub Command1_Click()
    rs.Open "Select * from Login where Username = '" & Text1.Text & "' and Password = '" & Text2.Text & "'", CONN
    If rs.EOF <> True Then
        MsgBox "Login Successful"
        Load Form5
        Form5.Show
        Unload Me
        rs.Close
        CONN.Close
    Else
        MsgBox "Check Username or password"
        rs.Close
    End If
End Sub

Private Sub Command2_Click()
    Load Form2
    Form2.Show vbModal, Form1
End Sub
    

Private Sub Command3_Click()
    Load Form3
    Form3.Show
    Unload Me
End Sub

Private Sub Form_Load()
    'Unload Me
    'Form1.Show
    Set CONN = New ADODB.Connection
    Set rs = New ADODB.Recordset
    CONN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=I:\Project\Final\Trial.mdb;Persist Security Info=False"
    CONN.Open
    rs.CursorType = adOpenDynamic
End Sub

