VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form pp 
   Caption         =   "Form2"
   ClientHeight    =   7095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14730
   LinkTopic       =   "Form2"
   ScaleHeight     =   7095
   ScaleWidth      =   14730
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton c3 
      Caption         =   "recherche avancée"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   5
      Top             =   6000
      Width           =   2295
   End
   Begin VB.CommandButton c2 
      Caption         =   "marche"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   6000
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   -240
      Top             =   6840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton c1 
      Caption         =   "tout le materiel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   6000
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "imprimer "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   6480
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DG 
      Height          =   5055
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   8916
      _Version        =   393216
      Enabled         =   -1  'True
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
            LCID            =   1036
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
            LCID            =   1036
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
   Begin VB.Label l1 
      Caption         =   "Stock :"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   240
      Width           =   5895
   End
End
Attribute VB_Name = "pp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ct As ADODB.Connection
Dim rc As ADODB.Recordset
Dim rc1 As ADODB.Recordset
Dim rc2 As ADODB.Recordset
Dim rc3 As ADODB.Recordset

Private Sub c3_Click()
rav.Show
pp.Hide
End Sub

Private Sub Form_Load()
Set ct = New ADODB.Connection
ct.Provider = "microsoft.jet.oledb.4.0"
ct.ConnectionString = "dataStock.mdb"
ct.Open
'*******************************************************stock******************************************************'
Set rc = New ADODB.Recordset
rc.CursorLocation = adUseClient
rc.Open "select * from materials where emplacement='stock'order by etat desc,type ", ct, adOpenDynamic, adLockOptimistic
 Set DG.DataSource = rc
'*******************************************************stock******************************************************'
Set rc1 = New ADODB.Recordset
rc1.CursorLocation = adUseClient
rc1.Open "select * from materials order by etat desc,type,emplacement", ct, adOpenDynamic, adLockOptimistic
'*******************************************************marche******************************************************'
Set rc2 = New ADODB.Recordset
rc2.CursorLocation = adUseClient
rc2.Open "select * from materials where etat='marche'order by emplacement desc,type ", ct, adOpenDynamic, adLockOptimistic
'*******************************************************hors-service************************************************'
Set rc3 = New ADODB.Recordset
rc3.CursorLocation = adUseClient
rc3.Open "select * from materials where etat='hors-service'order by emplacement desc,type ", ct, adOpenDynamic, adLockOptimistic

End Sub
Private Sub Command1_Click()
CommonDialog1.ShowPrinter
End Sub
Private Sub c1_Click()
If c1.Caption = "stock" Then
c1.Caption = "tout le materiel"
l1.Caption = "stock :"
Set DG.DataSource = rc
Else
c1.Caption = "stock"
l1.Caption = "tout le materiel :"
Set DG.DataSource = rc1
End If
End Sub
Private Sub c2_Click()
If c2.Caption = "marche" Then
c2.Caption = "hors-service"
l1.Caption = "les materiels marche :"
Set DG.DataSource = rc2
Else
c2.Caption = "marche"
l1.Caption = " les materiels hors-service  :"
Set DG.DataSource = rc3
End If
End Sub
