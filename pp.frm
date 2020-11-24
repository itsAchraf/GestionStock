VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form pp 
   ClientHeight    =   4590
   ClientLeft      =   5640
   ClientTop       =   3630
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   9735
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Fermer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   24
      Top             =   3720
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Recherche Multiple"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   23
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton m2 
      BackColor       =   &H00000000&
      Caption         =   "supprimer"
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
      Left            =   120
      TabIndex        =   22
      Top             =   2160
      Width           =   1815
   End
   Begin VB.ComboBox tm1 
      Height          =   315
      Left            =   4080
      TabIndex        =   21
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox tm2 
      Height          =   315
      Left            =   4080
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox tm8 
      Height          =   315
      Left            =   4080
      TabIndex        =   12
      Top             =   1320
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox tm5 
      Height          =   315
      Left            =   4080
      TabIndex        =   11
      Top             =   2520
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox tm6 
      Height          =   315
      Left            =   4080
      TabIndex        =   10
      Top             =   3120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox tm7 
      Height          =   315
      Left            =   4080
      TabIndex        =   9
      Top             =   3720
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox tm3 
      Height          =   315
      Left            =   4080
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox tm4 
      Height          =   315
      Left            =   4080
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1815
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
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
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
         TabIndex        =   4
         Top             =   840
         Width           =   1335
      End
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
         TabIndex        =   3
         Top             =   1320
         Width           =   975
      End
   End
   Begin VB.CommandButton m1 
      BackColor       =   &H00000000&
      Caption         =   "modifier"
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
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton aa 
      Caption         =   "ajoute"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   0
      Top             =   2760
      Width           =   2175
   End
   Begin MSDataGridLib.DataGrid DG 
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   7680
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1508
      _Version        =   393216
      Enabled         =   -1  'True
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
      TabIndex        =   20
      Top             =   120
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
      TabIndex        =   19
      Top             =   720
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
      TabIndex        =   18
      Top             =   1320
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
      TabIndex        =   17
      Top             =   1920
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
      TabIndex        =   16
      Top             =   3720
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
      TabIndex        =   15
      Top             =   3120
      Visible         =   0   'False
      Width           =   1815
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
      Top             =   2520
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "pp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ct As ADODB.Connection
Dim GDG As ADODB.Recordset
Dim GDG0 As ADODB.Recordset
Dim com As Integer
Dim n1 As Integer
Dim n2 As Integer
Dim pas As Integer
Sub op1()
tm1.Clear
tm5.Clear
tm6.Clear
tm7.Clear
tm2.Text = ""
tm3.Text = ""
tm4.Text = ""
tm5.Text = ""
tm6.Text = ""
tm8.Text = ""
tm7.Text = ""
com = 1
'*****************************************************************************************************************'
'//////////////////////////////////////////// C O M B O B O X ////////////////////////////////////////////////////'
'*****************************************************************************************************************'
'**********************************************COMBO1 *************************************************************'
Set GDG = New ADODB.Recordset
GDG.CursorLocation = adUseClient
GDG.Open "select  DISTINCT codeONEP from materials order by codeONEP", ct, adOpenDynamic, adLockOptimistic
 Do Until GDG.EOF
 tm1.AddItem GDG!codeONEP
 GDG.MoveNext
 Loop

'**********************************************COMBO 5 *************************************************************'
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


Private Sub aa_Click()
aj.Show
Unload pp
End Sub

Private Sub Command1_Click()
rech.Show
Unload pp
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub m1_Click()
If com = 1 Then
    If tm1.Text = "" Then
        MsgBox "entre le code ONEP": tm1.SetFocus
        Exit Sub
    Else
        With GDG
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
        ![fournisseur] = tm1.Text
        ![telephone] = tm2.Text
        ![email] = tm8.Text
        ![adress] = tm4.Text
        End With
    End If
ElseIf com = 3 Then
    If tm1.Text = "" Or tm2.Text = "" Then
        MsgBox "remplir tous les champs SVP": tm1.SetFocus
        Exit Sub
    Else
        With GDG
        ![matricule] = tm1.Text
        ![nomAgent] = tm2.Text
        .Update
        End With
    End If
End If
End Sub

Private Sub m2_Click()
If pas = 0 Then
Exit Sub
ElseIf pas = 1 Then
If com = 2 Then
    If n1 = "0" Then
    r = MsgBox("voulez vous vraiment Supprimer l'enregistrement ", vbYesNo, "Gestion materials")
    If r = vbNo Then
        Exit Sub
    Else
        GDG.Delete
        tm1.RemoveItem (tm1.ListIndex)
        tm2.Text = ""
        tm3.Text = ""
        tm4.Text = ""
        tm5.Text = ""
        tm6.Text = ""
        tm8.Text = ""
        tm7.Text = ""
    End If
    Else
       r = MsgBox("Il y a un ou plusieur materials de ce Fournisseur", vbTitleBarText, "attention")
    End If
ElseIf com = 3 Then
    If n2 = "0" Then
    r = MsgBox("voulez vous vraiment Supprimer l'enregistrement ", vbYesNo, "Gestion materials")
    If r = vbNo Then
        Exit Sub
    Else
        GDG.Delete
         tm1.RemoveItem (tm1.ListIndex)
        tm2.Text = ""
        tm3.Text = ""
        tm4.Text = ""
        tm5.Text = ""
        tm6.Text = ""
        tm8.Text = ""
        tm7.Text = ""
    End If
    Else
       r = MsgBox("Il y a un ou plusieur materials de ce Utilisateur", vbTitleBarText, "attention")
    End If
ElseIf com = 1 Then
    r = MsgBox("voulez vous vraiment Supprimer l'enregistrement ", vbYesNo, "Gestion materials")
    If r = vbNo Then
        Exit Sub
    Else
        With GDG
        .Delete
         tm1.RemoveItem (tm1.ListIndex)
        tm2.Text = ""
        tm3.Text = ""
        tm4.Text = ""
        tm5.Text = ""
        tm6.Text = ""
        tm8.Text = ""
        tm7.Text = ""
        End With
    End If
End If
pas = 0
End If
End Sub

Private Sub Command3_Click()
        Set GDG0 = New ADODB.Recordset
        GDG0.CursorLocation = adUseClient
        GDG0.Open "select  count(codeONEP)as n from materials where matricule='" & tm1.Text & "'", ct, adOpenDynamic, adLockOptimistic
        Set DG.DataSource = GDG0
     
End Sub

Private Sub Form_Load()
o1.Value = True
com = 0
Set ct = New ADODB.Connection
ct.Provider = "microsoft.jet.oledb.4.0"
ct.ConnectionString = "datastock.mdb"
ct.Open
op1
End Sub

Private Sub lb_Click()
  GDG.Delete
End Sub


Private Sub o1_Click()
m1.Visible = True
tm1.Clear
tm5.Clear
tm6.Clear
tm7.Clear
tm2.Text = ""
tm3.Text = ""
tm4.Text = ""
tm5.Text = ""
tm6.Text = ""
tm8.Text = ""
tm7.Text = ""
com = 1
'*****************************************************************************************************************'
'//////////////////////////////////////////// C O M B O B O X ////////////////////////////////////////////////////'
'*****************************************************************************************************************'
'**********************************************COMBO1 *************************************************************'
Set GDG = New ADODB.Recordset
GDG.CursorLocation = adUseClient
GDG.Open "select  DISTINCT codeONEP from materials order by codeONEP", ct, adOpenDynamic, adLockOptimistic
 Do Until GDG.EOF
 tm1.AddItem GDG!codeONEP
 GDG.MoveNext
 Loop

'**********************************************COMBO 5 *************************************************************'
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
m1.Visible = False
tm1.Clear
tm2.Text = ""
tm3.Text = ""
tm4.Text = ""
tm5.Text = ""
tm6.Text = ""
tm8.Text = ""
tm7.Text = ""
com = 3
'************************************************** COMBO 1 *******************************************************'
Set GDG = New ADODB.Recordset
GDG.CursorLocation = adUseClient
GDG.Open "select  DISTINCT matricule from agent order by matricule", ct, adOpenDynamic, adLockOptimistic
 Do Until GDG.EOF
 tm1.AddItem GDG!matricule
 GDG.MoveNext
 Loop
'******************************************************************************************************************'
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
m1.Visible = False
tm1.Clear
tm2.Text = ""
tm3.Text = ""
tm4.Text = ""
tm5.Text = ""
tm6.Text = ""
tm8.Text = ""
tm7.Text = ""
com = 2
'************************************************** COMBO 1 *******************************************************'
Set GDG = New ADODB.Recordset
GDG.CursorLocation = adUseClient
GDG.Open "select  DISTINCT fournisseur from fournisseur order by fournisseur", ct, adOpenDynamic, adLockOptimistic
 Do Until GDG.EOF
 tm1.AddItem GDG!fournisseur
 GDG.MoveNext
 Loop
'******************************************************************************************************************'
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
Private Sub tm1_click()
pas = 1
If com = 1 Then
        With GDG
        If .RecordCount > 0 Then .MoveFirst
        .Find " codeONEP = '" & tm1.Text & "'"
        If .EOF Then MsgBox "Les donne est déjà attribué": tm1.SetFocus: Exit Sub
        tm1.Text = ![codeONEP]
        tm2.Text = ![Nserie]
        tm3.Text = ![matricule]
        tm4.Text = ![Marque]
        tm5.Text = ![Categorie]
        tm6.Text = ![etat]
       tm7.Text = ![fournisseur]
        End With
ElseIf com = 2 Then
        With GDG
        If .RecordCount > 0 Then .MoveFirst
        .Find " fournisseur = '" & tm1.Text & "'"
        If .EOF Then MsgBox "Les donne est déjà attribué": tm1.SetFocus: Exit Sub
        tm1.Text = ![fournisseur]
        tm2.Text = ![telephone]
        tm8.Text = ![email]
        tm4.Text = ![adress]
        End With
        Set GDG0 = New ADODB.Recordset
        GDG0.CursorLocation = adUseClient
        GDG0.Open "select  count(codeONEP)as n from materials where fournisseur ='" & tm1.Text & "'", ct, adOpenDynamic, adLockOptimistic
        n1 = GDG0!n
ElseIf com = 3 Then
        With GDG
        If .RecordCount > 0 Then .MoveFirst
        .Find " Matricule = '" & tm1.Text & "'"
        If .EOF Then MsgBox "Les donne est déjà attribué": tm1.SetFocus: Exit Sub
        tm1.Text = ![matricule]
        tm2.Text = ![nomAgent]
        End With
        Set GDG0 = New ADODB.Recordset
        GDG0.CursorLocation = adUseClient
        GDG0.Open "select  count(codeONEP)as n from materials where matricule='" & tm1.Text & "'", ct, adOpenDynamic, adLockOptimistic
        n2 = GDG0!n

End If
End Sub
