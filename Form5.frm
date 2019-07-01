VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form5 
   Caption         =   " "
   ClientHeight    =   8490
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14745
   LinkTopic       =   "Form5"
   ScaleHeight     =   8490
   ScaleWidth      =   14745
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Apple"
      Height          =   3375
      Left            =   360
      TabIndex        =   1
      Top             =   4440
      Width           =   9015
      Begin MSChart20Lib.MSChart MSChart2 
         Height          =   2655
         Left            =   240
         OleObjectBlob   =   "Form5.frx":0000
         TabIndex        =   3
         Top             =   480
         Width           =   7935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Google"
      Height          =   3375
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   9015
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   2655
         Left            =   240
         OleObjectBlob   =   "Form5.frx":2356
         TabIndex        =   2
         Top             =   360
         Width           =   7935
      End
   End
   Begin VB.Menu sg 
      Caption         =   "Stock Graphs"
   End
   Begin VB.Menu cust 
      Caption         =   "Customer"
      Begin VB.Menu add 
         Caption         =   "Add New"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu pur 
      Caption         =   "Purchase"
   End
   Begin VB.Menu logout 
      Caption         =   "Logout"
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim conn As ADODB.Connection
Private Sub add_Click()
Form6.Show
End Sub

Private Sub Form_Load()
Frame1.Visible = False
Frame2.Visible = False
Load Form6
Load Form7
Set rs = New ADODB.Recordset
Set conn = New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=I:\Project\Final\Trial.mdb;Persist Security Info=False"
conn.Open
rs.Open "Select price from Google", conn, adOpenKeyset
With MSChart1
.ShowLegend = True
Set .DataSource = rs
End With
rs.Close
rs.Open "Select price from Apple", conn, adOpenKeyset
With MSChart2
.ShowLegend = True
Set .DataSource = rs
End With
End Sub

Private Sub logout_Click()
End
End Sub

Private Sub pur_Click()
Form7.Show
End Sub


Private Sub sg_Click()
Frame1.Visible = True
Frame2.Visible = True
End Sub
