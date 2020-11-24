VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form rav 
   Caption         =   "Form1"
   ClientHeight    =   8730
   ClientLeft      =   2025
   ClientTop       =   1260
   ClientWidth     =   12210
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   12210
   Begin VB.ComboBox fc2 
      Height          =   315
      Left            =   4680
      TabIndex        =   6
      Top             =   6120
      Width           =   2415
   End
   Begin VB.ComboBox fc1 
      Height          =   315
      Left            =   4680
      TabIndex        =   5
      Top             =   5520
      Width           =   2415
   End
   Begin VB.OptionButton to1 
      Caption         =   "Fournissour"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   5520
      Width           =   1575
   End
   Begin VB.OptionButton to2 
      Caption         =   "Materials"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   6120
      Width           =   1575
   End
   Begin VB.OptionButton to3 
      Caption         =   "Emplacement"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   5160
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid DG 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
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
   Begin VB.Label l2 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   6120
      Width           =   2055
   End
   Begin VB.Label l1 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   5520
      Width           =   2175
   End
End
Attribute VB_Name = "rav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ct As ADODB.Connection
Dim rc As ADODB.Recordset
Dim rc1 As ADODB.Recordset
Dim rc2 As ADODB.Recordset
Dim rc3 As ADODB.Recordset

Private Sub fc1_click()
fc2.Clear
Set rc2 = New ADODB.Recordset
rc2.CursorLocation = adUseClient
rc2.Open "select distinct nomF from fournisseur where adress='" & fc1.Text & "' order by nomF", ct, adOpenDynamic, adLockOptimistic
Do Until rc2.EOF
fc2.AddItem rc2![nomF]
rc2.MoveNext
Loop

End Sub
Private Sub fc2_click()
Set rc = New ADODB.Recordset
rc.CursorLocation = adUseClient
rc.Open "select * from fournisseur where nomF like'" & fc2.Text & "'", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = rc
End Sub

Private Sub Command1_Click()
Text1.Text = c1.Text
Set rc2 = New ADODB.Recordset
rc2.CursorLocation = adUseClient
rc2.Open "select distinct nomF from fournisseur where adress like '" & c1.Text & "' order by nomF", ct, adOpenDynamic, adLockOptimistic
Do Until rc2.EOF
c2.AddItem rc2![nomF]
rc2.MoveNext
Loop
End Sub

Private Sub Form_Load()
Set ct = New ADODB.Connection
ct.Provider = "microsoft.jet.oledb.4.0"
ct.ConnectionString = "dataStock.mdb"
ct.Open
End Sub
'**************************************************fournisseur******************************************************'
Private Sub to1_Click()
l1.Caption = "Adress :"
l2.Caption = "Nom de fournisseur :"
fc1.Clear
fc2.Clear
Set rc = New ADODB.Recordset
rc.CursorLocation = adUseClient
rc.Open "select * from fournisseur ", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = rc
Set rc1 = New ADODB.Recordset
rc1.CursorLocation = adUseClient
rc1.Open "select distinct adress from fournisseur order by adress", ct, adOpenDynamic, adLockOptimistic
Do Until rc1.EOF
fc1.AddItem rc1![adress]
rc1.MoveNext
Loop
Do Until rc.EOF
fc2.AddItem rc![nomF]
rc.MoveNext
Loop
End Sub
'**************************************************materials******************************************************'
Private Sub to2_Click()
fc1.Clear
fc2.Clear
Set rc = New ADODB.Recordset
rc.CursorLocation = adUseClient
rc.Open "select * from materials ", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = rc
End Sub
'**************************************************Emplacement******************************************************'
Private Sub to3_Click()
Set rc = New ADODB.Recordset
rc.CursorLocation = adUseClient
rc.Open "select * from Emplacement where emplacement <> 'stock' ", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = rc
End Sub
