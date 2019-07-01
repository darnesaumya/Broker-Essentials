VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form7"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   10200
      TabIndex        =   10
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   10200
      TabIndex        =   8
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   6360
      TabIndex        =   3
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   13560
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   6360
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Height          =   495
      Left            =   7800
      TabIndex        =   0
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Rate per stock"
      Height          =   495
      Left            =   8640
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Brokerage"
      Height          =   495
      Left            =   8640
      TabIndex        =   9
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Brokerage and Total"
      BeginProperty Font 
         Name            =   "Lucida Sans Typewriter"
         Size            =   24
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   7
      Top             =   480
      Width           =   6135
   End
   Begin VB.Label Label3 
      Caption         =   "Total"
      Height          =   495
      Left            =   4800
      TabIndex        =   6
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Brokerage percent"
      Height          =   495
      Left            =   12000
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "No. of Stocks"
      Height          =   495
      Left            =   4800
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Private Sub Command1_Click()
Dim n As Integer
Dim t As Integer
Dim r As Integer
Dim per As Integer
Dim total As Integer
Dim brokerage As Integer
n = Text1.Text
r = Text5.Text
per = Text2.Text
t = n * r
total = ((t * per) / 100) + t
brokerage = (t * per) / 100
Text3.Text = total
Text4.Text = brokerage

 conn.Open
    rs.LockType = adLockOptimistic
    rs.Open "Calculate", conn
    rs.AddNew
    rs.Fields(0) = Text3.Text
    rs.Fields(1) = Text4.Text
    rs.update
    rs.Close
    conn.Close
End Sub

Private Sub Form_Load()
Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=I:\Project\Final\Trial.mdb;Persist Security Info=False"
End Sub
