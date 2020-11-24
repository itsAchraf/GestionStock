VERSION 5.00
Begin VB.Form intro 
   ClientHeight    =   2355
   ClientLeft      =   7365
   ClientTop       =   4665
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   ScaleHeight     =   2355
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox user 
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   240
      Width           =   3015
   End
   Begin VB.TextBox psw 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   840
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
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
      Left            =   1200
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Entrer"
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
      Left            =   2880
      TabIndex        =   0
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label msg 
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label go 
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Utilisateur"
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
      Left            =   480
      TabIndex        =   5
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "mot de passe"
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
      Left            =   480
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "intro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
If user.Text = "admin" And psw.Text = "user0" Then
pp.Show
Unload intro
ElseIf user.Text = "" Or psw.Text = "" Then
r = MsgBox("Entrez le nom d'utilisation et le mot de passe SVP.", vbTitleBarText, "Message d'erreur")
user.Text = ""
psw.Text = ""
user.SetFocus
ElseIf user.Text <> "admin" Or psw.Text <> "user0" Then
r = MsgBox("le nom d'utilisation ou le mot de passe est incorrect.", vbTitleBarText, "Message d'erreur")
user.Text = ""
psw.Text = ""
user.SetFocus
End If
End Sub

Private Sub go_Click()
pp.Show
Unload intro
End Sub

Private Sub msg_Click()
r = MsgBox("username:  admin" & vbNewLine & "password:  user0 ")
End Sub
