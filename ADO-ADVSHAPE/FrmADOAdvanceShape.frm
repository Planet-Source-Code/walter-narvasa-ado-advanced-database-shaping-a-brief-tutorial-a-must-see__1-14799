VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ADO Advance Shape Sample by Walter A. Narvasa"
   ClientHeight    =   6225
   ClientLeft      =   1095
   ClientTop       =   615
   ClientWidth     =   8880
   Icon            =   "FrmADOAdvanceShape.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame2 
      Caption         =   "SQL Statement Executed:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2280
      TabIndex        =   12
      Top             =   5040
      Width           =   6495
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   6255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Important Notes:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   2055
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   2775
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   2775
      Left            =   120
      ScaleHeight     =   2715
      ScaleWidth      =   1995
      TabIndex        =   1
      Top             =   120
      Width           =   2055
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   300
         Left            =   120
         MouseIcon       =   "FrmADOAdvanceShape.frx":0E42
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   2280
         Width           =   1800
      End
      Begin VB.CommandButton cmdSingleLevel 
         Caption         =   "&Single Level Relation"
         Height          =   300
         Left            =   120
         MouseIcon       =   "FrmADOAdvanceShape.frx":114C
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   120
         Width           =   1800
      End
      Begin VB.CommandButton cmdMultipleLevel 
         Caption         =   "&Multi Level Relation"
         Height          =   300
         Left            =   120
         MouseIcon       =   "FrmADOAdvanceShape.frx":1456
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   480
         Width           =   1800
      End
      Begin VB.CommandButton cmdParameterized 
         Caption         =   "&Parameterized"
         Height          =   300
         Left            =   120
         MouseIcon       =   "FrmADOAdvanceShape.frx":1760
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   840
         Width           =   1800
      End
      Begin VB.CommandButton cmdMultipleRelation 
         Caption         =   "M&ultiple Relation"
         Height          =   300
         Left            =   120
         MouseIcon       =   "FrmADOAdvanceShape.frx":1A6A
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   1200
         Width           =   1800
      End
      Begin VB.CommandButton cmdWithAggregate 
         Caption         =   "&Relation with Agregate"
         Height          =   300
         Left            =   120
         MouseIcon       =   "FrmADOAdvanceShape.frx":1D74
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   1560
         Width           =   1800
      End
      Begin VB.CommandButton cmdGroup 
         Caption         =   "&Group"
         Height          =   300
         Left            =   120
         MouseIcon       =   "FrmADOAdvanceShape.frx":207E
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   1920
         Width           =   1800
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfgMain 
      Height          =   4455
      Left            =   2280
      TabIndex        =   0
      ToolTipText     =   "If you want to see the effect just scroll up/down/left/right."
      Top             =   480
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7858
      _Version        =   393216
      BackColor       =   65280
      WordWrap        =   -1  'True
      AllowUserResizing=   3
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ADO Recordset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2280
      TabIndex        =   9
      Top             =   120
      Width           =   6495
   End
   Begin VB.Menu options 
      Caption         =   "&Options"
      Begin VB.Menu slr 
         Caption         =   "&Single Level Relation"
      End
      Begin VB.Menu mlr 
         Caption         =   "&Multi Level Relation"
      End
      Begin VB.Menu p 
         Caption         =   "&Parameterized"
      End
      Begin VB.Menu mr 
         Caption         =   "M&ultiple Relation"
      End
      Begin VB.Menu rwa 
         Caption         =   "&Relation with Agregate"
      End
      Begin VB.Menu g 
         Caption         =   "&Group"
      End
      Begin VB.Menu bar 
         Caption         =   "-"
      End
      Begin VB.Menu shutdown 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu about 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Author: Walter A. Narvasa
' Copyright (c) 2000
' Email: jawoltze@edsamail.com.ph / walter_narvasa@hotmail.com
' Website: http://jawoltze.gq.nu/

Private mcnn As New ADODB.Connection
Private mrst As New ADODB.Recordset

Private Sub Form_Load()
  Set mcnn = New Connection
  mcnn.CursorLocation = adUseClient
  strDatabase = "C:\Program Files\Microsoft Visual Studio\VB98\NWIND.MDB"
  mcnn.Open "PROVIDER=MSDataShape;Data PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & strDatabase & ";Jet OLEDB:Database Password='';"
End Sub

Private Sub about_Click()
    msg = MsgBox("This is a sample ADO Advanced Shape" & vbCrLf & "By Walter A. Narvasa" & vbCrLf & "For more details email me at" & vbCrLf & "walter_narvasa@hotmail.com or" & vbCrLf & "jawoltze@edsamail.com.ph", vbOKOnly, "About")
End Sub

Private Sub cmdGroup_Click()
    Set mrst = New ADODB.Recordset
    mrst.Open "SHAPE {SELECT Customers.CustomerID AS CustID, Customers.CompanyName, Orders.* FROM Customers INNER JOIN Orders ON Customers.CustomerID = Orders.CustomerID} AS rstOrders COMPUTE rstOrders BY CustID, CompanyName", mcnn, adOpenForwardOnly, adLockReadOnly
    mrst.Requery
    Set hfgMain.DataSource = mrst
    Label1.Caption = "Group ADO Recordset"
    Label2.Caption = ""
    Label3.Caption = "SHAPE {SELECT Customers.CustomerID AS CustID, Customers.CompanyName, Orders.* FROM Customers INNER JOIN Orders ON Customers.CustomerID = Orders.CustomerID} AS rstOrders COMPUTE rstOrders BY CustID, CompanyName"
End Sub

Private Sub cmdMultipleLevel_Click()
    Set mrst = New ADODB.Recordset
    mrst.Open "SHAPE {SELECT * FROM Customers} APPEND ((SHAPE {SELECT * FROM Orders} APPEND ({SELECT * FROM [Order Details]} AS rstOrderDetails RELATE OrderID TO OrderID)) RELATE CustomerID TO CustomerID)", mcnn, adOpenForwardOnly, adLockReadOnly
    mrst.Requery
    Set hfgMain.DataSource = mrst
    Label1.Caption = "Multi Level Relation ADO Recordset"
    Label2.Caption = "The next step up in complexity is to nest two SHAPE and APPEND commands to create a Recordset based on the Customers, Orders, and Order Details tables."
    Label3.Caption = "SHAPE {SELECT * FROM Customers} APPEND ((SHAPE {SELECT * FROM Orders} APPEND ({SELECT * FROM [Order Details]} AS rstOrderDetails RELATE OrderID TO OrderID)) RELATE CustomerID TO CustomerID)"
End Sub

Private Sub cmdMultipleRelation_Click()
    Set mrst = New ADODB.Recordset
    mrst.Open "SHAPE {SELECT * FROM Customers} APPEND({SELECT * FROM Orders WHERE ShippedDate > #1/1/97#} RELATE CustomerID TO CustomerID) as rstNewOrders, ({SELECT * FROM Orders WHERE ShippedDate <= #1/1/97#} RELATE CustomerID TO CustomerID) as rstOldOrders", mcnn, adOpenForwardOnly, adLockReadOnly
    mrst.Requery
    Set hfgMain.DataSource = mrst
    Label1.Caption = "Multiple Relation ADO Recordset"
    Label2.Caption = "By using more than one clause in the APPEND part of the SHAPE statement, you can create a Recordset with more than one chapter field, and thus more than one child Recordset."
    Label3.Caption = "SHAPE {SELECT * FROM Customers} APPEND({SELECT * FROM Orders WHERE ShippedDate > #1/1/97#} RELATE CustomerID TO CustomerID) as rstNewOrders, ({SELECT * FROM Orders WHERE ShippedDate <= #1/1/97#} RELATE CustomerID TO CustomerID) as rstOldOrders"
End Sub

Private Sub cmdParameterized_Click()
    Set mrst = New ADODB.Recordset
    mrst.Open "SHAPE {" & _
     "SELECT * " & _
     "FROM Customers} APPEND ({SELECT * " & _
     "FROM Orders WHERE CustomerID = ?} RELATE CustomerID " & _
     "TO PARAMETER 0)", mcnn, adOpenForwardOnly, adLockReadOnly
    mrst.Requery
    Set hfgMain.DataSource = mrst
    Label1.Caption = "Parameterized ADO Recordset"
    Label2.Caption = "Theres no difference between the Recordset retrieved by a parameterized hierarchy and that retrieved by the equivalent relation hierarchy. Here is the parameterized equivalent of the first, single-level example."
    Label3.Caption = "SHAPE {" & "SELECT * " & "FROM Customers} APPEND ({SELECT * " & "FROM Orders WHERE CustomerID = ?} RELATE CustomerID " & "TO PARAMETER 0)"
End Sub

Private Sub cmdSingleLevel_Click()
    Set mrst = New ADODB.Recordset
    mrst.Open "SHAPE {" & _
     "SELECT * " & _
     "FROM Customers} APPEND ({SELECT * " & _
     "FROM Orders} RELATE CustomerID " & _
     "TO CustomerID)", mcnn, adOpenForwardOnly, adLockReadOnly
    mrst.Requery
    Set hfgMain.DataSource = mrst
    Label1.Caption = "Single Level Relation ADO Recordset"
    Label2.Caption = "A single-level relation hierarchy relates two Recordsets, in this case recordsets based on the Customers and Orders tables."
    Label3.Caption = "SHAPE {" & "SELECT * " & "FROM Customers} APPEND ({SELECT * " & "FROM Orders} RELATE CustomerID " & "TO CustomerID)"
End Sub

Private Sub cmdWithAggregate_Click()
    Set mrst = New ADODB.Recordset
    mrst.Open "SHAPE {" & _
     "SELECT * " & _
     "FROM Customers} APPEND ({SELECT * " & _
     "FROM Orders} RELATE CustomerID " & _
     "TO CustomerID), MIN(Chapter1.ShippedDate) AS FirstShip", mcnn, adOpenForwardOnly, adLockReadOnly
    mrst.Requery
    Set hfgMain.DataSource = mrst
    Label1.Caption = "&Relation with Agregate ADO Recordset"
    Label2.Caption = "This creates a Recordset with Customer and Order information, plus an additional aggregate column that contains the minimum value from any record in the ShippedDate column for each customer."
    Label3.Caption = "SHAPE {" & "SELECT * " & "FROM Customers} APPEND ({SELECT * " & "FROM Orders} RELATE CustomerID " & "TO CustomerID), MIN(Chapter1.ShippedDate) AS FirstShip"
End Sub

Private Sub Command1_Click()
    shutdown_Click
End Sub

Private Sub g_Click()
    cmdGroup_Click
End Sub

Private Sub mlr_Click()
    cmdMultipleLevel_Click
End Sub

Private Sub mr_Click()
    cmdMultipleRelation_Click
End Sub

Private Sub p_Click()
    cmdParameterized_Click
End Sub

Private Sub rwa_Click()
    cmdWithAggregate_Click
End Sub

Private Sub shutdown_Click()
    mrst.Close
    Unload Me
End Sub

Private Sub slr_Click()
    cmdSingleLevel_Click
End Sub
