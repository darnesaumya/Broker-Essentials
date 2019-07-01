VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form6 
   BackColor       =   &H000080FF&
   Caption         =   "Form6"
   ClientHeight    =   8730
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16140
   LinkTopic       =   "Form6"
   ScaleHeight     =   8730
   ScaleWidth      =   16140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      Height          =   495
      Left            =   10080
      TabIndex        =   21
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update"
      Height          =   495
      Left            =   8040
      TabIndex        =   20
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Clear"
      Height          =   495
      Left            =   2040
      TabIndex        =   17
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00000000&
      Caption         =   "GENERATE REPORT"
      Height          =   495
      Left            =   3720
      TabIndex        =   16
      Top             =   8160
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Register"
      Height          =   495
      Left            =   6000
      TabIndex        =   15
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   11880
      TabIndex        =   14
      Top             =   8160
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   240
      Top             =   0
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1085
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
      RecordSource    =   "Customer"
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
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Height          =   6855
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   13215
      Begin VB.TextBox Text1 
         DataField       =   "CName"
         Height          =   615
         Left            =   4920
         TabIndex        =   7
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         DataField       =   "CPhone"
         Height          =   615
         Left            =   4920
         TabIndex        =   6
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         DataField       =   "CAddress"
         Height          =   615
         Left            =   4920
         TabIndex        =   5
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         DataField       =   "CEmail"
         Height          =   615
         Left            =   9120
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         DataField       =   "Age"
         Height          =   615
         Left            =   9120
         TabIndex        =   3
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text7 
         DataField       =   "City"
         Height          =   615
         Left            =   9120
         TabIndex        =   2
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox Text9 
         DataField       =   "ID"
         Height          =   405
         Left            =   1320
         TabIndex        =   1
         Top             =   120
         Width           =   855
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form6.frx":0000
         Height          =   2775
         Left            =   240
         TabIndex        =   18
         Top             =   3600
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   4895
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   16393
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   615
         Left            =   3000
         TabIndex        =   13
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone Number"
         Height          =   615
         Left            =   3000
         TabIndex        =   12
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   615
         Left            =   3000
         TabIndex        =   11
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         Height          =   615
         Left            =   7200
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         Height          =   615
         Left            =   7200
         TabIndex        =   9
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         Height          =   615
         Left            =   7200
         TabIndex        =   8
         Top             =   2640
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn, con As ADODB.Connection
Dim rs, rs1 As ADODB.Recordset
Dim vid As Integer

Private Sub Command1_Click()
    conn.Open
    rs.LockType = adLockOptimistic
    rs.Open "Customer", conn
    rs.AddNew
    rs.Fields(1) = Text1.Text
    rs.Fields(2) = Text2.Text
    rs.Fields(3) = Text3.Text
    rs.Fields(4) = Text5.Text
    rs.Fields(5) = Text6.Text
    rs.Fields(6) = Text7.Text
    rs.Update
    rs.Close
    conn.Close
    MsgBox "Registered Succesfully"
    DataGrid1.Refresh
End Sub

Private Sub Command2_Click()
conn.Open
rs.LockType = adLockOptimistic
rs.Open "Customer", conn
rs.Update
MsgBox "Updated"
DataGrid1.Refresh
rs.Close
conn.Close
End Sub

Private Sub Command3_Click()
conn.Open
rs.LockType = adLockOptimistic
rs.Open "Customer", conn
rs.Delete
MsgBox "Deleted"
rs.Close
conn.Close
DataGrid1.Refresh
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
Dim CON1 As ADODB.Connection
Dim RSS As ADODB.Recordset
Set CON1 = New ADODB.Connection
Set RSS = New ADODB.Recordset

CON1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=I:\Project\Final\Trial.mdb;Persist Security Info=False"
CON1.Open
RSS.Open "Customer", CON1, adOpenDynamic, adLockOptimistic
Set DataReport1.DataSource = RSS
DataReport1.Show


End Sub

Private Sub DataGrid1_Click()
vid = CInt(DataGrid1.Text)
End Sub

Private Sub Form_Load()
    Adodc1.Visible = False
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=I:\Project\Final\Trial.mdb;Persist Security Info=False"
    Set con = New ADODB.Connection
    Set rs1 = New ADODB.Recordset
    rs1.CursorLocation = adUseClient
    con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=I:\Project\Final\Trial.mdb;Persist Security Info=False"
    con.Open
    rs1.Open "Select * from Customer", con, adOpenKeyset, adLockPessimistic, adcmdtxt
    Set DataGrid1.DataSource = rs1
    DataGrid1.Refresh
    Set rs1 = Nothing
End Sub


