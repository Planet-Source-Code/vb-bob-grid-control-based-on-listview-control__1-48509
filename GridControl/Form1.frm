VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   ScaleHeight     =   7965
   ScaleWidth      =   10725
   StartUpPosition =   2  'CenterScreen
   Begin Project1.GridControl GridControl1 
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   5953
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Show Num Rows Selected"
      Height          =   495
      Left            =   7920
      TabIndex        =   43
      Top             =   7320
      Width           =   1815
   End
   Begin VB.Frame Frame6 
      Caption         =   "Sort by column"
      Height          =   1335
      Left            =   7200
      TabIndex        =   25
      Top             =   5280
      Width           =   3375
      Begin VB.CommandButton cmdSort 
         Caption         =   "Sort"
         Height          =   495
         Left            =   1800
         TabIndex        =   28
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox cboSort 
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   240
         List            =   "Form1.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtColSort 
         Height          =   285
         Left            =   600
         TabIndex        =   27
         Text            =   "3"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Col"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   285
         Width           =   225
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Set Tool Tips"
      Height          =   975
      Left            =   3720
      TabIndex        =   36
      Top             =   6720
      Width           =   4095
      Begin VB.CommandButton cmdTooltips 
         Caption         =   "Set Tooltips"
         Height          =   495
         Left            =   2640
         TabIndex        =   40
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtToolCol1 
         Height          =   285
         Left            =   1680
         TabIndex        =   38
         Text            =   "2"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtToolCol2 
         Height          =   285
         Left            =   1680
         TabIndex        =   41
         Text            =   "3"
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Column for Header"
         Height          =   195
         Left            =   240
         TabIndex        =   37
         Top             =   285
         Width           =   1320
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Column for Body"
         Height          =   195
         Left            =   240
         TabIndex        =   39
         Top             =   645
         Width           =   1155
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Load Grid"
      Height          =   1575
      Left            =   240
      TabIndex        =   1
      Top             =   3600
      Width           =   7695
      Begin VB.CommandButton cmdLoadGrid 
         Caption         =   "Load Grid"
         Height          =   495
         Left            =   6240
         TabIndex        =   8
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtMsgCol 
         Height          =   285
         Left            =   5040
         TabIndex        =   7
         Text            =   "1"
         Top             =   795
         Width           =   735
      End
      Begin VB.TextBox txtSQL 
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Text            =   "select * from cmn_state "
         Top             =   435
         Width           =   6735
      End
      Begin VB.TextBox txtMsg 
         Height          =   285
         Left            =   1680
         TabIndex        =   5
         Text            =   "Nothing"
         Top             =   795
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "SQL:"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "No Data Message"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   1290
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "No Data Msg Column"
         Height          =   195
         Left            =   3240
         TabIndex        =   6
         Top             =   840
         Width           =   1515
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Set Grid Data"
      Height          =   1335
      Left            =   240
      TabIndex        =   9
      Top             =   5280
      Width           =   3375
      Begin VB.TextBox txtData 
         Height          =   285
         Left            =   720
         TabIndex        =   17
         Text            =   "text"
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdSetData 
         Caption         =   "Set Data"
         Height          =   495
         Left            =   2040
         TabIndex        =   15
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtRow3 
         Height          =   285
         Left            =   720
         TabIndex        =   11
         Text            =   "2"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtCol3 
         Height          =   285
         Left            =   720
         TabIndex        =   14
         Text            =   "3"
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "use 0 for current row"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1560
         TabIndex        =   12
         Top             =   285
         Width           =   1455
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   1005
         Width           =   345
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Row"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   285
         Width           =   330
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Col"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   645
         Width           =   225
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Set Column Width"
      Height          =   975
      Left            =   240
      TabIndex        =   30
      Top             =   6720
      Width           =   3375
      Begin VB.TextBox txtWidth 
         Height          =   285
         Left            =   720
         TabIndex        =   32
         Text            =   "1000"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdSetColWidth 
         Caption         =   "Set Width"
         Height          =   495
         Left            =   1800
         TabIndex        =   34
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtCol2 
         Height          =   285
         Left            =   720
         TabIndex        =   35
         Text            =   "3"
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Width"
         Height          =   195
         Left            =   240
         TabIndex        =   31
         Top             =   285
         Width           =   420
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Col"
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   645
         Width           =   225
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Get Grid Data"
      Height          =   1335
      Left            =   3720
      TabIndex        =   18
      Top             =   5280
      Width           =   3375
      Begin VB.TextBox txtCol 
         Height          =   285
         Left            =   720
         TabIndex        =   23
         Text            =   "3"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtRow 
         Height          =   285
         Left            =   720
         TabIndex        =   20
         Text            =   "2"
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdShowData 
         Caption         =   "Show Data"
         Height          =   495
         Left            =   1920
         TabIndex        =   24
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "use 0 for current row"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1560
         TabIndex        =   21
         Top             =   285
         Width           =   1455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Col"
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   645
         Width           =   225
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Row"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   285
         Width           =   330
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show Rowcount"
      Height          =   495
      Left            =   7920
      TabIndex        =   42
      Top             =   6720
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdLoadGrid_Click()
    
    GridControl1.LoadGrid txtSQL.Text, DataEnvironmentPDR.PDR.ConnectionString, txtMsg.Text, Val(txtMsgCol.Text)
    
End Sub


Private Sub cmdSetColWidth_Click()

    GridControl1.SetColumnWidth txtCol2.Text, txtWidth.Text
    
End Sub

Private Sub cmdShowData_Click()

    MsgBox GridControl1.GetGridData(txtRow.Text, txtCol.Text)
    
End Sub

Private Sub cmdSort_Click()

    If cboSort.Text = "ASC" Then
        GridControl1.SortByColumn txtColSort.Text, True
    Else
        GridControl1.SortByColumn txtColSort.Text, False
    End If
    
End Sub

Private Sub cmdTooltips_Click()
    
    GridControl1.SetToolTips txtToolCol1.Text, txtToolCol2.Text

End Sub

Private Sub cmdSetData_Click()

    GridControl1.SetGridData txtRow3.Text, txtCol3.Text, txtData.Text
    
End Sub

Private Sub Command2_Click()

    MsgBox GridControl1.RowCount
    
End Sub


Private Sub Command3_Click()

    MsgBox GridControl1.SelectedCount
    
End Sub

