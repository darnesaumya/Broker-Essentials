VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   6735
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11085
   LinkTopic       =   "Form4"
   ScaleHeight     =   6735
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Microsoft"
      Height          =   3255
      Left            =   12840
      TabIndex        =   4
      Top             =   720
      Width           =   4935
      Begin MSChart20Lib.MSChart MSChart3 
         Height          =   2175
         Left            =   600
         OleObjectBlob   =   "Form4.frx":0000
         TabIndex        =   5
         Top             =   600
         Width           =   3975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Apple"
      Height          =   3255
      Left            =   6600
      TabIndex        =   2
      Top             =   720
      Width           =   4935
      Begin MSChart20Lib.MSChart MSChart2 
         Height          =   2175
         Left            =   480
         OleObjectBlob   =   "Form4.frx":2354
         TabIndex        =   3
         Top             =   600
         Width           =   3975
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   0
      Top             =   8040
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
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
      RecordSource    =   "Stock"
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
      Caption         =   "Google"
      Height          =   3255
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   4935
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   2175
         Left            =   480
         OleObjectBlob   =   "Form4.frx":46A8
         TabIndex        =   1
         Top             =   600
         Width           =   3975
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Private Sub Form_Load()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=I:\Project\Final\Trial.mdb;Persist Security Info=False"
conn.Open
rs.CursorLocation = adUseClient
rs.Open "Select Google from Login", conn, adOpenKeyset
With MSChart1
    
End With
rs.Close
conn.Close
End Sub

