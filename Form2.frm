VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00FF8080&
   Caption         =   "Form2"
   ClientHeight    =   7665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15435
   LinkTopic       =   "Form2"
   ScaleHeight     =   7665
   ScaleWidth      =   15435
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   10200
      TabIndex        =   20
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      Height          =   495
      Left            =   8640
      TabIndex        =   19
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update"
      Height          =   495
      Left            =   6960
      TabIndex        =   18
      Top             =   7200
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=I:\Project\Final\Trial.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=I:\Project\Final\Trial.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Login"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   2400
      TabIndex        =   1
      Top             =   1440
      Width           =   11295
      Begin VB.TextBox Text9 
         Height          =   405
         Left            =   1320
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text8 
         Height          =   615
         Left            =   9120
         TabIndex        =   17
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox Text7 
         Height          =   615
         Left            =   9120
         TabIndex        =   16
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         Height          =   615
         Left            =   9120
         TabIndex        =   15
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Height          =   615
         Left            =   9120
         TabIndex        =   14
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Height          =   615
         IMEMode         =   3  'DISABLE
         Left            =   4920
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   615
         Left            =   4920
         TabIndex        =   12
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   615
         Left            =   4920
         TabIndex        =   11
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         DataSource      =   "Adodc1"
         Height          =   615
         Left            =   4920
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         Height          =   615
         Left            =   7200
         TabIndex        =   9
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         Height          =   615
         Left            =   7200
         TabIndex        =   8
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         Height          =   615
         Left            =   7200
         TabIndex        =   7
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Pincode"
         Height          =   615
         Left            =   7200
         TabIndex        =   6
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         Height          =   615
         Left            =   3000
         TabIndex        =   5
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
         Height          =   615
         Left            =   3000
         TabIndex        =   4
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone Number"
         Height          =   615
         Left            =   3000
         TabIndex        =   3
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   615
         Left            =   3000
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Register"
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Broker Registration"
      BeginProperty Font 
         Name            =   "Lucida Sans"
         Size            =   24
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5145
      TabIndex        =   21
      Top             =   240
      Width           =   5190
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CONN As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub Command1_Click()
    CONN.Open
    rs.LockType = adLockOptimistic
    rs.Open "Login", CONN
    rs.AddNew
    rs.Fields(0) = Text1.Text
    rs.Fields(1) = Text2.Text
    rs.Fields(2) = Text3.Text
    rs.Fields(3) = Text4.Text
    rs.Fields(4) = Text5.Text
    rs.Fields(5) = Text6.Text
    rs.Fields(6) = Text7.Text
    rs.Fields(7) = Text8.Text
    rs.Update
    rs.Close
    CONN.Close
    MsgBox "Registered Succesfully"
    Unload Me
    Form4.Show
End Sub

Private Sub Command2_Click()
CONN.Open
rs.LockType = adLockOptimistic
rs.Open "Login", CONN
rs.Update
rs.Close
CONN.Close
End Sub

Private Sub Command3_Click()
CONN.Open
rs.LockType = adLockOptimistic
rs.Open "Login", CONN
rs.Delete
rs.Close
CONN.Close
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Set CONN = New ADODB.Connection
    Set rs = New ADODB.Recordset
    CONN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=I:\Project\Final\Trial.mdb;Persist Security Info=False"
End Sub
