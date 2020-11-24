VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form rech 
   ClientHeight    =   8085
   ClientLeft      =   4140
   ClientTop       =   1710
   ClientWidth     =   12885
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   12885
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Retour 
      Caption         =   "Retour"
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
      Left            =   10680
      TabIndex        =   13
      Top             =   6480
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Recherche"
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
      Left            =   10680
      TabIndex        =   12
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   2055
      Left            =   1680
      TabIndex        =   11
      Top             =   9840
      Width           =   1335
   End
   Begin VB.ComboBox tm2 
      Height          =   315
      Left            =   2280
      TabIndex        =   4
      Top             =   6600
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox tm5 
      Height          =   315
      Left            =   7320
      TabIndex        =   3
      Top             =   6600
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox tm4 
      Height          =   315
      Left            =   7320
      TabIndex        =   2
      Top             =   6000
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox tm3 
      Height          =   315
      Left            =   2280
      TabIndex        =   1
      Top             =   7200
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ComboBox tm1 
      Height          =   315
      Left            =   2280
      TabIndex        =   0
      Top             =   6000
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSDataGridLib.DataGrid DG 
      Height          =   5655
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   9975
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
   Begin VB.Label lm3 
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
      Left            =   360
      TabIndex        =   10
      Top             =   7200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lm4 
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
      Left            =   5400
      TabIndex        =   9
      Top             =   6000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lm5 
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
      Left            =   5400
      TabIndex        =   8
      Top             =   6600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lm1 
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
      Left            =   360
      TabIndex        =   7
      Top             =   6000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lm2 
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
      Left            =   360
      TabIndex        =   6
      Top             =   6600
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "rech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ct As ADODB.Connection
Dim GDG0 As ADODB.Recordset
Dim GDG As ADODB.Recordset
Dim DATG As ADODB.Recordset
Dim v1 As Integer
Dim v2 As Integer
Dim v3 As Integer
Dim v4 As Integer
Dim v5 As Integer
Private Sub Command1_Click()
RchDt1
End Sub

Private Sub Form_Load()
Set ct = New ADODB.Connection
ct.Provider = "microsoft.jet.oledb.4.0"
ct.ConnectionString = "datastock.mdb"
ct.Open
intro
End Sub

Sub intro()
tm1.Clear
tm2.Clear
tm3.Clear
tm4.Clear
tm5.Clear
tm1.Text = ""
tm2.Text = ""
tm3.Text = ""
tm4.Text = ""
tm5.Text = ""
tm1.Visible = True
tm2.Visible = True
tm3.Visible = True
tm4.Visible = True
tm5.Visible = True
lm1.Visible = True
lm2.Visible = True
lm3.Visible = True
lm4.Visible = True
lm5.Visible = True
lm1.Caption = "Matricule utilisateur :"
lm2.Caption = "Marque :"
lm3.Caption = "Catégorie :"
lm4.Caption = "Etat :"
lm5.Caption = "Fournisseur :"
'******************************************************************************************************************'
'//////////////////////////////////////////// C O M B O B O X /////////////////////////////////////////////////////'
'******************************************************************************************************************'
'**********************************************COMBO1 *************************************************************'
Set GDG0 = New ADODB.Recordset
GDG0.CursorLocation = adUseClient
GDG0.Open "select  DISTINCT matricule from agent order by matricule", ct, adOpenDynamic, adLockOptimistic
Do Until GDG0.EOF
tm1.AddItem GDG0!matricule
GDG0.MoveNext
Loop
'**********************************************COMBO2 *************************************************************'
Set GDG = New ADODB.Recordset
GDG.CursorLocation = adUseClient
GDG.Open "select  DISTINCT Marque from materials order by Marque", ct, adOpenDynamic, adLockOptimistic
 Do Until GDG.EOF
 tm2.AddItem GDG!Marque
 GDG.MoveNext
 Loop
'**********************************************COMBO3 *************************************************************'
Set GDG = New ADODB.Recordset
GDG.CursorLocation = adUseClient
GDG.Open "select  DISTINCT Categorie from materials order by Categorie", ct, adOpenDynamic, adLockOptimistic
 Do Until GDG.EOF
 tm3.AddItem GDG!Categorie
 GDG.MoveNext
 Loop
'**********************************************COMBO4 *************************************************************'
Set GDG = New ADODB.Recordset
GDG.CursorLocation = adUseClient
GDG.Open "select  DISTINCT etat from materials order by etat", ct, adOpenDynamic, adLockOptimistic
 Do Until GDG.EOF
 tm4.AddItem GDG!etat
 GDG.MoveNext
 Loop
'**********************************************COMBO5 *************************************************************'
Set GDG0 = New ADODB.Recordset
GDG0.CursorLocation = adUseClient
GDG0.Open "select  DISTINCT fournisseur from fournisseur order by fournisseur", ct, adOpenDynamic, adLockOptimistic
Do Until GDG0.EOF
tm5.AddItem GDG0!fournisseur
GDG0.MoveNext
Loop
'*****************************************************************************************************************'
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////'
'*****************************************************************************************************************'
Set GDG = New ADODB.Recordset
GDG.CursorLocation = adUseClient
GDG.Open "select * from materials ", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = GDG
End Sub
Sub pasValide()
If tm1.Text = "" Then
v1 = 0
Else
v1 = 1
End If
If tm2.Text = "" Then
v2 = 0
Else
v2 = 1
End If
If tm3.Text = "" Then
v3 = 0
Else
v3 = 1
End If
If tm4.Text = "" Then
v4 = 0
Else
v4 = 1
End If
If tm5.Text = "" Then
v5 = 0
Else
v5 = 1
End If
End Sub
Sub RchDt1()
pasValide
If v1 = 1 And v2 = 1 And v3 = 1 And v4 = 1 And v5 = 1 Then
Set DATG = New ADODB.Recordset
DATG.CursorLocation = adUseClient
DATG.Open "select * from materials where Matricule='" & tm1.Text & "' and marque='" & tm2.Text & "' and Categorie='" & tm3.Text & "' and etat='" & tm4.Text & "' and Fournisseur='" & tm5.Text & "'", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = DATG
ElseIf v1 = 1 And v2 = 1 And v3 = 0 And v4 = 1 And v5 = 1 Then
Set DATG = New ADODB.Recordset
DATG.CursorLocation = adUseClient
DATG.Open "select * from materials where Matricule='" & tm1.Text & "' and marque='" & tm2.Text & "'and etat='" & tm4.Text & "'and Fournisseur='" & tm5.Text & "'", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = DATG
ElseIf v1 = 1 And v2 = 0 And v3 = 1 And v4 = 1 And v5 = 1 Then
Set DATG = New ADODB.Recordset
DATG.CursorLocation = adUseClient
DATG.Open "select * from materials where Matricule='" & tm1.Text & "' and Categorie='" & tm3.Text & "'and etat='" & tm4.Text & "'and Fournisseur='" & tm5.Text & "'", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = DATG
ElseIf v1 = 0 And v2 = 1 And v3 = 1 And v4 = 1 And v5 = 1 Then
Set DATG = New ADODB.Recordset
DATG.CursorLocation = adUseClient
DATG.Open "select * from materials where marque='" & tm2.Text & "' and Categorie='" & tm3.Text & "'and etat='" & tm4.Text & "'and Fournisseur='" & tm5.Text & "'", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = DATG
ElseIf v1 = 1 And v2 = 0 And v3 = 0 And v4 = 1 And v5 = 1 Then
Set DATG = New ADODB.Recordset
DATG.CursorLocation = adUseClient
DATG.Open "select * from materials where Matricule='" & tm1.Text & "' and etat='" & tm4.Text & "'and Fournisseur='" & tm5.Text & "'", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = DATG
ElseIf v1 = 0 And v2 = 0 And v3 = 1 And v4 = 1 And v5 = 1 Then
Set DATG = New ADODB.Recordset
DATG.CursorLocation = adUseClient
DATG.Open "select * from materials where Categorie='" & tm3.Text & "'and etat='" & tm4.Text & "'and Fournisseur='" & tm5.Text & "'", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = DATG
ElseIf v1 = 0 And v2 = 1 And v3 = 0 And v4 = 1 And v5 = 1 Then
Set DATG = New ADODB.Recordset
DATG.CursorLocation = adUseClient
DATG.Open "select * from materials where  marque='" & tm2.Text & "'and etat='" & tm4.Text & "'and Fournisseur='" & tm5.Text & "'", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = DATG
ElseIf v1 = 0 And v2 = 0 And v3 = 0 And v4 = 1 And v5 = 1 Then
Set DATG = New ADODB.Recordset
DATG.CursorLocation = adUseClient
DATG.Open "select * from materials and etat='" & tm4.Text & "'and Fournisseur='" & tm5.Text & "'", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = DATG
'************************************************************************************************************************************'
ElseIf v1 = 1 And v2 = 1 And v3 = 1 And v4 = 0 And v5 = 1 Then
Set DATG = New ADODB.Recordset
DATG.CursorLocation = adUseClient
DATG.Open "select * from materials where Matricule='" & tm1.Text & "' and marque='" & tm2.Text & "' and Categorie='" & tm3.Text & "'and Fournisseur='" & tm5.Text & "'", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = DATG
ElseIf v1 = 1 And v2 = 1 And v3 = 0 And v4 = 0 And v5 = 1 Then
Set DATG = New ADODB.Recordset
DATG.CursorLocation = adUseClient
DATG.Open "select * from materials where Matricule='" & tm1.Text & "' and marque='" & tm2.Text & "'and Fournisseur='" & tm5.Text & "'", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = DATG
ElseIf v1 = 1 And v2 = 0 And v3 = 1 And v4 = 0 And v5 = 1 Then
Set DATG = New ADODB.Recordset
DATG.CursorLocation = adUseClient
DATG.Open "select * from materials where Matricule='" & tm1.Text & "' and Categorie='" & tm3.Text & "'and Fournisseur='" & tm5.Text & "'", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = DATG
ElseIf v1 = 0 And v2 = 1 And v3 = 1 And v4 = 0 And v5 = 1 Then
Set DATG = New ADODB.Recordset
DATG.CursorLocation = adUseClient
DATG.Open "select * from materials where marque='" & tm2.Text & "' and Categorie='" & tm3.Text & "'and Fournisseur='" & tm5.Text & "'", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = DATG
ElseIf v1 = 1 And v2 = 0 And v3 = 0 And v4 = 0 And v5 = 1 Then
Set DATG = New ADODB.Recordset
DATG.CursorLocation = adUseClient
DATG.Open "select * from materials where Matricule='" & tm1.Text & "' and Fournisseur='" & tm5.Text & "'", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = DATG
ElseIf v1 = 0 And v2 = 0 And v3 = 1 And v4 = 0 And v5 = 1 Then
Set DATG = New ADODB.Recordset
DATG.CursorLocation = adUseClient
DATG.Open "select * from materials where Categorie='" & tm3.Text & "'and Fournisseur='" & tm5.Text & "'", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = DATG
ElseIf v1 = 0 And v2 = 1 And v3 = 0 And v4 = 0 And v5 = 1 Then
Set DATG = New ADODB.Recordset
DATG.CursorLocation = adUseClient
DATG.Open "select * from materials where  marque='" & tm2.Text & "'and Fournisseur='" & tm5.Text & "'", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = DATG
ElseIf v1 = 0 And v2 = 0 And v3 = 0 And v4 = 0 And v5 = 1 Then
Set DATG = New ADODB.Recordset
DATG.CursorLocation = adUseClient
DATG.Open "select * from materials where Fournisseur='" & tm5.Text & "'", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = DATG
'************************************************************************************************************************************'
'************************************************************************************************************************************'
ElseIf v1 = 1 And v2 = 1 And v3 = 1 And v4 = 1 And v5 = 0 Then
Set DATG = New ADODB.Recordset
DATG.CursorLocation = adUseClient
DATG.Open "select * from materials where Matricule='" & tm1.Text & "' and marque='" & tm2.Text & "' and Categorie='" & tm3.Text & "' and etat='" & tm4.Text & "'", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = DATG
ElseIf v1 = 1 And v2 = 1 And v3 = 0 And v4 = 1 And v5 = 0 Then
Set DATG = New ADODB.Recordset
DATG.CursorLocation = adUseClient
DATG.Open "select * from materials where Matricule='" & tm1.Text & "' and marque='" & tm2.Text & "'and etat='" & tm4.Text & "'", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = DATG
ElseIf v1 = 1 And v2 = 0 And v3 = 1 And v4 = 1 And v5 = 0 Then
Set DATG = New ADODB.Recordset
DATG.CursorLocation = adUseClient
DATG.Open "select * from materials where Matricule='" & tm1.Text & "' and Categorie='" & tm3.Text & "'and etat='" & tm4.Text & "'", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = DATG
ElseIf v1 = 0 And v2 = 1 And v3 = 1 And v4 = 1 And v5 = 0 Then
Set DATG = New ADODB.Recordset
DATG.CursorLocation = adUseClient
DATG.Open "select * from materials where marque='" & tm2.Text & "' and Categorie='" & tm3.Text & "'and etat='" & tm4.Text & "'", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = DATG
ElseIf v1 = 1 And v2 = 0 And v3 = 0 And v4 = 1 And v5 = 0 Then
Set DATG = New ADODB.Recordset
DATG.CursorLocation = adUseClient
DATG.Open "select * from materials where Matricule='" & tm1.Text & "' and etat='" & tm4.Text & "'", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = DATG
ElseIf v1 = 0 And v2 = 0 And v3 = 1 And v4 = 1 And v5 = 0 Then
Set DATG = New ADODB.Recordset
DATG.CursorLocation = adUseClient
DATG.Open "select * from materials where Categorie='" & tm3.Text & "'and etat='" & tm4.Text & "'", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = DATG
ElseIf v1 = 0 And v2 = 1 And v3 = 0 And v4 = 1 And v5 = 0 Then
Set DATG = New ADODB.Recordset
DATG.CursorLocation = adUseClient
DATG.Open "select * from materials where  marque='" & tm2.Text & "'and etat='" & tm4.Text & "'", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = DATG
ElseIf v1 = 0 And v2 = 0 And v3 = 0 And v4 = 1 And v5 = 0 Then
Set DATG = New ADODB.Recordset
DATG.CursorLocation = adUseClient
DATG.Open "select * from materials where etat='" & tm4.Text & "'", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = DATG
'************************************************************************************************************************************'
ElseIf v1 = 1 And v2 = 1 And v3 = 1 And v4 = 0 And v5 = 0 Then
Set DATG = New ADODB.Recordset
DATG.CursorLocation = adUseClient
DATG.Open "select * from materials where Matricule='" & tm1.Text & "' and marque='" & tm2.Text & "' and Categorie='" & tm3.Text & "'", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = DATG
ElseIf v1 = 1 And v2 = 1 And v3 = 0 And v4 = 0 And v5 = 0 Then
Set DATG = New ADODB.Recordset
DATG.CursorLocation = adUseClient
DATG.Open "select * from materials where Matricule='" & tm1.Text & "' and marque='" & tm2.Text & "'", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = DATG
ElseIf v1 = 1 And v2 = 0 And v3 = 1 And v4 = 0 And v5 = 0 Then
Set DATG = New ADODB.Recordset
DATG.CursorLocation = adUseClient
DATG.Open "select * from materials where Matricule='" & tm1.Text & "' and Categorie='" & tm3.Text & "'", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = DATG
ElseIf v1 = 0 And v2 = 1 And v3 = 1 And v4 = 0 And v5 = 0 Then
Set DATG = New ADODB.Recordset
DATG.CursorLocation = adUseClient
DATG.Open "select * from materials where marque='" & tm2.Text & "' and Categorie='" & tm3.Text & "'", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = DATG
ElseIf v1 = 1 And v2 = 0 And v3 = 0 And v4 = 0 And v5 = 0 Then
Set DATG = New ADODB.Recordset
DATG.CursorLocation = adUseClient
DATG.Open "select * from materials where Matricule='" & tm1.Text & "' ", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = DATG
ElseIf v1 = 0 And v2 = 0 And v3 = 1 And v4 = 0 And v5 = 0 Then
Set DATG = New ADODB.Recordset
DATG.CursorLocation = adUseClient
DATG.Open "select * from materials where Categorie='" & tm3.Text & "'", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = DATG
ElseIf v1 = 0 And v2 = 1 And v3 = 0 And v4 = 0 And v5 = 0 Then
Set DATG = New ADODB.Recordset
DATG.CursorLocation = adUseClient
DATG.Open "select * from materials where  marque='" & tm2.Text & "'", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = DATG
ElseIf v1 = 0 And v2 = 0 And v3 = 0 And v4 = 0 And v5 = 0 Then
Set DATG = New ADODB.Recordset
DATG.CursorLocation = adUseClient
DATG.Open "select * from materials ", ct, adOpenDynamic, adLockOptimistic
Set DG.DataSource = DATG
'************************************************************************************************************************************'
'************************************************************************************************************************************'
End If
End Sub

Private Sub o2_Click()
tm1.Clear
tm2.Clear
tm3.Clear
tm4.Clear
tm5.Clear
tm1.Text = ""
tm2.Text = ""
tm3.Text = ""
tm4.Text = ""
tm5.Text = ""
tm1.Visible = True
tm2.Visible = True
tm3.Visible = True
tm4.Visible = True
tm5.Visible = True
lm1.Visible = True
lm2.Visible = True
lm3.Visible = True
lm4.Visible = True
lm5.Visible = True
lm1.Caption = "Matricule utilisateur :"
lm2.Caption = "Marque :"
lm3.Caption = "Catégorie :"
lm4.Caption = "Etat :"
lm5.Caption = "Fournisseur :"

End Sub

Private Sub Retour_Click()
Unload rech
pp.Show
End Sub

Private Sub tm1_click()
RchDt1
End Sub


Private Sub tm2_click()
RchDt1
End Sub


Private Sub tm3_click()
RchDt1
End Sub

Private Sub tm4_click()
RchDt1
End Sub

Private Sub tm5_click()
RchDt1
End Sub
