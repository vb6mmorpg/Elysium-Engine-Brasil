VERSION 5.00
Begin VB.Form frmChars 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Elysium Engine Brasil"
   ClientHeight    =   3780
   ClientLeft      =   120
   ClientTop       =   -45
   ClientWidth     =   3435
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   3435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstChars 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2505
      ItemData        =   "frmChars.frx":0000
      Left            =   360
      List            =   "frmChars.frx":0002
      TabIndex        =   0
      Top             =   960
      Width           =   2745
   End
   Begin VB.Label picUseChar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Usar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   360
      TabIndex        =   4
      Top             =   720
      Width           =   825
   End
   Begin VB.Label picNewChar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Criar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2280
      TabIndex        =   3
      Top             =   720
      Width           =   765
   End
   Begin VB.Label picCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.Label picDelChar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Deletar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   720
      Width           =   900
   End
End
Attribute VB_Name = "frmChars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright (c) 2008 - Elysium Source. Alguns direitos reservados.
' Tradução e revisão por MMODEV Brasil @ http://www.mmodev.com.br
' Este código está licensiado sob a licença EGL.

Option Explicit

Private Sub Form_Load()
Dim I As Long
Dim Ending As String
    For I = 1 To 3
        If I = 1 Then Ending = ".gif"
        If I = 2 Then Ending = ".jpg"
        If I = 3 Then Ending = ".png"
 
        If FileExist("GUI\mediumlist" & Ending) Then frmChars.Picture = LoadPicture(App.Path & "\GUI\mediumlist" & Ending)
    Next I
End Sub

Private Sub picCancel_Click()
    Call TcpDestroy
    Me.Hide
    frmLogin.Show , frmMainMenu
End Sub

Private Sub picNewChar_Click()
    If lstChars.List(lstChars.ListIndex) <> "Lugar Livre" Then
        MsgBox "Já existe um personagem ocupando esse lugar!"
        Exit Sub
    End If
    Call MenuState(MENU_STATE_NEWCHAR)
End Sub

Private Sub picUseChar_Click()
    If lstChars.List(lstChars.ListIndex) = "Lugar Livre" Then
        MsgBox "Não há nenhum personagem nesse lugar!"
        Exit Sub
    End If
    Call MenuState(MENU_STATE_USECHAR)
End Sub

Private Sub picDelChar_Click()
Dim Value As Long

    If lstChars.List(lstChars.ListIndex) = "Lugar Livre" Then
        MsgBox "Não há nenhum personagem nesse lugar!"
        Exit Sub
    End If

    Value = MsgBox("Você tem certeza que deseja deletar esse personagem?", vbYesNo, GAME_NAME)
    If Value = vbYes Then
        Call MenuState(MENU_STATE_DELCHAR)
    End If
End Sub

