VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form aj 
   ClientHeight    =   5325
   ClientLeft      =   4080
   ClientTop       =   1710
   ClientWidth     =   7440
   FillColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleMode       =   0  'User
   ScaleWidth      =   7584.128
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00000000&
      Caption         =   "Fermer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   21
      Top             =   4560
      Width           =   1815
   End
   Begin VB.TextBox tm4 
      Height          =   315
      Left            =   4080
      TabIndex        =   20
      Top             =   2040
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox tm3 
      Height          =   315
      Left            =   4080
      TabIndex        =   19
      Top             =   1440
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000000&
      Caption         =   "Ajouter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   18
      Top             =   4560
      Width           =   1815
   End
   Begin VB.ComboBox tm7 
      Height          =   315
      Left            =   4080
      TabIndex        =   17
      Top             =   3840
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox tm6 
      Height          =   315
      Left            =   4080
      TabIndex        =   16
      Top             =   3240
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox tm5 
      Height          =   315
      Left            =   4080
      TabIndex        =   15
      Top             =   2640
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox tm8 
      Height          =   315
      Left            =   4080
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox tm2 
      Height          =   315
      Left            =   4080
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox tm1 
      Height          =   315
      Left            =   4080
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1815
      Begin VB.OptionButton o3 
         Caption         =   "agent"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   975
      End
      Begin VB.OptionButton o2 
         Caption         =   "fournisseur"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton o1 
         Caption         =   "materials"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid DG 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   6240
      Visible         =   0   'False
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   9763
      _Version        =   393216
      Enabled         =   0   'False
      ForeColor       =   -2147483630
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
   Begin VB.Label lm5 
      Alignment       =   1  'Right Justify
      Caption         =   "Catégorie :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   14
      Top             =   2640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lm6 
      Alignment       =   1  'Right Justify
      Caption         =   "Etat :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      Top             =   3240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lm7 
      Alignment       =   1  'Right Justify
      Caption         =   "Fournisseur  :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   3840
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lm4 
      Alignment       =   1  'Right Justify
      Caption         =   "Marque :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   2040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lm3 
      Alignment       =   1  'Right Justify
      Caption         =   "Matricule utilisateur :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   1440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lm2 
      Alignment       =   1  'Right Justify
      Caption         =   "Numéro série :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lm1 
      Alignment       =   1  'Right Justify
      Caption         =   "Code ONEP :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "aj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ct As ADODB.Connection
Dim GDG As ADODB.Recordset
Dim GDG0 As ADODB.Recordset
Dim com As Integer
Sub op1()
tm5.Clear
tm6.Clear
tm3.Clear
tm7.Clear
'*****************************************************************************************************************'
'//////////////////////////////////////////// C O M B O B O X ////////////////////////////////////////////////////'
'*****************************************************************************************************************'
'**********************************************COMBO 1 *************************************************************'
Set GDG = New ADODB.Recordset
GDG.CursorLocation = adUseClient
GDG.Open "select  DISTINCT categorie from materials order by categorie", ct, adOpenDynamic, adLockOptimistic
 Do Until GDG.EOF
 tm5.AddItem GDG!Categorie
 GDG.MoveNext
 Loop
 '**********************************************COMBO 2 *************************************************************'
Set GDG = New ADODB.Recordset
GDG.CursorLocation = adUseClient
GDG.Open "select  DISTINCT etat from materials order by etat", ct, adOpenDynamic, adLockOptimistic
 Do Until GDG.EOF
 tm6.AddItem GDG!etat
 GDG.MoveNext
 Loop
 '**********************************************COMBO 3 *************************************************************'
Set GDG0 = New ADODB.Recordset
GDG0.CursorLocation = adUseClient
GDG0.Open "select  DISTINCT fournisseur from fournisseur order by fournisseur", ct, adOpenDynamic, adLockOptimistic
Do Until GDG0.EOF
tm7.AddItem GDG0!fournisseur
GDG0.MoveNext
Loop
 '**********************************************COMBO 3 *************************************************************'
Set GDG0 = New ADODB.Recordset
GDG0.CursorLocation = adUseClient
GDG0.Open "select  DISTINCT matricule from agent order by matricule", ct, adOpenDynamic, adLockOptimistic
Do Until GDG0.EOF
tm3.AddItem GDG0!matricule
GDG0.MoveNext
Loop
'*****************************************************************************************************************'
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////'
'*****************************************************************************************************************'
tm2.Text = ""
tm3.Text = ""
tm4.Text = ""
tm5.Text = ""
tm6.Text = ""
tm8.Text = ""
tm7.Text = ""
com = 1
Set GDG = New ADODB.Recordset
GDG.CursorLocation = adUseClient
GDG.Open "select * from materials order by codeONEP", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = GDG
lm1.Caption = "Code ONEP :"
lm2.Caption = "Numéro série :"
lm3.Caption = "Matricule utilisateur :"
lm4.Caption = "Marque :"
lm1.Visible = True
lm2.Visible = True
lm3.Visible = True
lm4.Visible = True
lm5.Visible = True
lm6.Visible = True
lm7.Visible = True
tm1.Visible = True
tm2.Visible = True
tm3.Visible = True
tm4.Visible = True
tm8.Visible = False
tm5.Visible = True
tm6.Visible = True
tm7.Visible = True

End Sub
Private Sub Command1_Click()
If com = 1 Then
    If tm1.Text = "" Then
        MsgBox "entre le code ONEP": tm1.SetFocus
        Exit Sub
    Else
        With GDG
        If .RecordCount > 0 Then .MoveFirst
        .Find " codeONEP = '" & tm1.Text & "'"
        If Not .EOF Then MsgBox "Les donne est déjà attribué": tm1.SetFocus: Exit Sub
        .AddNew
        ![codeONEP] = tm1.Text
        ![Nserie] = tm2.Text
        ![matricule] = tm3.Text
        ![Marque] = tm4.Text
        ![Categorie] = tm5.Text
        ![etat] = tm6.Text
        ![fournisseur] = tm7.Text
        .Update
        End With
    End If
ElseIf com = 2 Then
    If tm1.Text = "" Or tm2.Text = "" Then
        MsgBox "remplir tous les champs SVP": tm1.SetFocus
        Exit Sub
    Else
        With GDG
        If .RecordCount > 0 Then .MoveFirst
        .Find " fournisseur = '" & tm1.Text & "'"
        .Find " telephone = '" & tm2.Text & "'"
        If Not .EOF Then MsgBox "Les donne est déjà attribué": tm1.SetFocus: Exit Sub
        .AddNew
        ![fournisseur] = tm1.Text
        ![telephone] = tm2.Text
        ![email] = tm8.Text
        ![adress] = tm4.Text
        .Update
        End With
    End If
ElseIf com = 3 Then
    If tm1.Text = "" Or tm2.Text = "" Then
        MsgBox "remplir tous les champs SVP": tm1.SetFocus
        Exit Sub
    Else
        With GDG
        If .RecordCount > 0 Then .MoveFirst
        .Find " Matricule = '" & tm1.Text & "'"
        .Find " nomAgent = '" & tm2.Text & "'"
        If Not .EOF Then MsgBox "Les donne est déjà attribué": tm1.SetFocus: Exit Sub
        .AddNew
        ![matricule] = tm1.Text
        ![nomAgent] = tm2.Text
        .Update
        End With
    End If
End If
tm1.Text = ""
tm2.Text = ""
tm3.Text = ""
tm4.Text = ""
tm5.Text = ""
tm6.Text = ""
tm8.Text = ""
tm7.Text = ""
End Sub

Private Sub Command2_Click()
pp.Show
Unload aj
End Sub

Private Sub Form_Load()
com = 0
Set ct = New ADODB.Connection
ct.Provider = "microsoft.jet.oledb.4.0"
ct.ConnectionString = "datastock.mdb"
ct.Open
op1
End Sub

Private Sub o1_Click()
Command1.Top = 4500
Command2.Top = 4500
aj.Height = 5910
tm5.Clear
tm6.Clear
tm3.Clear
tm7.Clear
'*****************************************************************************************************************'
'//////////////////////////////////////////// C O M B O B O X ////////////////////////////////////////////////////'
'*****************************************************************************************************************'
'**********************************************COMBO 1 *************************************************************'
Set GDG = New ADODB.Recordset
GDG.CursorLocation = adUseClient
GDG.Open "select  DISTINCT categorie from materials order by categorie", ct, adOpenDynamic, adLockOptimistic
 Do Until GDG.EOF
 tm5.AddItem GDG!Categorie
 GDG.MoveNext
 Loop
 '**********************************************COMBO 2 *************************************************************'
Set GDG = New ADODB.Recordset
GDG.CursorLocation = adUseClient
GDG.Open "select  DISTINCT etat from materials order by etat", ct, adOpenDynamic, adLockOptimistic
 Do Until GDG.EOF
 tm6.AddItem GDG!etat
 GDG.MoveNext
 Loop
 '**********************************************COMBO 3 *************************************************************'
Set GDG0 = New ADODB.Recordset
GDG0.CursorLocation = adUseClient
GDG0.Open "select  DISTINCT fournisseur from fournisseur order by fournisseur", ct, adOpenDynamic, adLockOptimistic
Do Until GDG0.EOF
tm7.AddItem GDG0!fournisseur
GDG0.MoveNext
Loop
 '**********************************************COMBO 3 *************************************************************'
Set GDG0 = New ADODB.Recordset
GDG0.CursorLocation = adUseClient
GDG0.Open "select  DISTINCT matricule from agent order by matricule", ct, adOpenDynamic, adLockOptimistic
Do Until GDG0.EOF
tm3.AddItem GDG0!matricule
GDG0.MoveNext
Loop
'*****************************************************************************************************************'
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////'
'*****************************************************************************************************************'
tm2.Text = ""
tm3.Text = ""
tm4.Text = ""
tm5.Text = ""
tm6.Text = ""
tm8.Text = ""
tm7.Text = ""
com = 1
Set GDG = New ADODB.Recordset
GDG.CursorLocation = adUseClient
GDG.Open "select * from materials order by codeONEP", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = GDG
lm1.Caption = "Code ONEP :"
lm2.Caption = "Numéro série :"
lm3.Caption = "Matricule utilisateur :"
lm4.Caption = "Marque :"
lm1.Visible = True
lm2.Visible = True
lm3.Visible = True
lm4.Visible = True
lm5.Visible = True
lm6.Visible = True
lm7.Visible = True
tm1.Visible = True
tm2.Visible = True
tm3.Visible = True
tm4.Visible = True
tm8.Visible = False
tm5.Visible = True
tm6.Visible = True
tm7.Visible = True


End Sub

Private Sub o3_Click()
Command1.Top = 2760
Command2.Top = 2760
aj.Height = 3900
tm1.Text = ""
tm2.Text = ""
tm3.Text = ""
tm4.Text = ""
tm5.Text = ""
tm6.Text = ""
tm8.Text = ""
tm7.Text = ""
com = 3
lm1.Caption = "Matricule :"
lm2.Caption = "Nom Utilisateur :"
lm1.Visible = True
lm2.Visible = True
lm3.Visible = False
lm4.Visible = False
lm5.Visible = False
lm6.Visible = False
lm7.Visible = False
tm1.Visible = True
tm2.Visible = True
tm3.Visible = False
tm4.Visible = False
tm5.Visible = False
tm6.Visible = False
tm7.Visible = False
tm8.Visible = False
Set GDG = New ADODB.Recordset
GDG.CursorLocation = adUseClient
GDG.Open "select * from agent ", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = GDG
End Sub

Private Sub o2_Click()
Command1.Top = 2760
Command2.Top = 2760
aj.Height = 3900
tm1.Text = ""
tm2.Text = ""
tm3.Text = ""
tm4.Text = ""
tm5.Text = ""
tm6.Text = ""
tm8.Text = ""
tm7.Text = ""
com = 2
lm1.Caption = "Fornisseur :"
lm2.Caption = "Telephone :"
lm3.Caption = "Email :"
lm4.Caption = "Adress :"
lm1.Visible = True
lm2.Visible = True
lm3.Visible = True
lm4.Visible = True
lm5.Visible = False
lm6.Visible = False
lm7.Visible = False
tm1.Visible = True
tm2.Visible = True
tm3.Visible = False
tm4.Visible = True
tm5.Visible = False
tm6.Visible = False
tm7.Visible = False
tm8.Visible = True
Set GDG = New ADODB.Recordset
GDG.CursorLocation = adUseClient
GDG.Open "select * from fournisseur ", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = GDG
End Sub
