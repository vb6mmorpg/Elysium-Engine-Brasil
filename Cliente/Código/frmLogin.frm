VERSION 5.00
Begin VB.Form frmLogin 
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
   ScaleHeight     =   252
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   229
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrInfo 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   5520
      Top             =   0
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Save Password"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   195
   End
   Begin VB.TextBox txtPassword 
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
      Height          =   225
      IMEMode         =   3  'DISABLE
      Left            =   480
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1800
      Width           =   2475
   End
   Begin VB.TextBox txtName 
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
      Height          =   225
      Left            =   480
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1320
      Width           =   2475
   End
   Begin VB.Label lblSenha 
      BackStyle       =   0  'Transparent
      Caption         =   "Senha:"
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
      Index           =   2
      Left            =   480
      TabIndex        =   9
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblLogin 
      BackStyle       =   0  'Transparent
      Caption         =   "Login:"
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
      Index           =   1
      Left            =   480
      TabIndex        =   8
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Lembrar Senha"
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
      Height          =   165
      Index           =   0
      Left            =   720
      TabIndex        =   7
      Top             =   2160
      Width           =   930
   End
   Begin VB.Label lblOnOff 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "O Servidor está Offline"
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
      Left            =   1440
      TabIndex        =   6
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label lblPlayers 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "com XXX jogadores"
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
      TabIndex        =   5
      Top             =   3360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label picConnect 
      Alignment       =   2  'Center
      BackColor       =   &H00789298&
      BackStyle       =   0  'Transparent
      Caption         =   "Entrar"
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
      Height          =   165
      Left            =   1530
      TabIndex        =   4
      Top             =   2640
      Width           =   405
   End
   Begin VB.Label picCancel 
      Alignment       =   2  'Center
      BackColor       =   &H00789298&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   345
      Left            =   3000
      TabIndex        =   3
      Top             =   120
      Width           =   330
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright (c) 2009 - Elysium Source. Alguns direitos reservados.
' Tradução e revisão por MMODEV Brasil @ http://www.mmodev.com.br
' Este código está licensiado sob a licença EGL.

Option Explicit

Private Sub Form_Load()
Dim I As Long
Dim Ending As String
Dim Packet As String

    For I = 1 To 3
        If I = 1 Then Ending = ".gif"
        If I = 2 Then Ending = ".jpg"
        If I = 3 Then Ending = ".png"

        If FileExist("GUI\medium" & Ending) Then frmLogin.Picture = LoadPicture(App.Path & "\GUI\medium" & Ending)
    Next I

    txtName.Text = Trim(GetVar(App.Path & "\config.ini", "CONFIG", "Conta"))
    txtPassword.Text = Trim(GetVar(App.Path & "\config.ini", "CONFIG", "Senha"))
    If Trim(txtPassword.Text) <> vbNullString Then
        Check1.Value = Checked
    Else
        Check1.Value = Unchecked
    End If
        
    lblPlayers.Visible = True
    lblPlayers.Caption = "Pegando Informações..."
    
    If ConnectToServer = True Then
        tmrInfo.Enabled = True
        Packet = "getinfo" & END_CHAR
        Call SendData(Packet)
    Else
        lblOnOff.Caption = "O Servidor está Offline"
        lblPlayers.Visible = False
    End If
End Sub

Private Sub picCancel_Click()
    frmLogin.Visible = False
    frmMainMenu.Visible = True
End Sub

Private Sub picConnect_Click()
    If Trim(txtName.Text) <> vbNullString And Trim(txtPassword.Text) <> vbNullString Then
        If Len(Trim(txtName.Text)) < 3 Or Len(Trim(txtPassword.Text)) < 3 Then
            MsgBox "Seu login e sua senha devem possuir, no mínimo, 3 caractéres."
            Exit Sub
        End If
        Call MenuState(MENU_STATE_LOGIN)
        Call PutVar(App.Path & "\config.ini", "CONFIG", "Conta", txtName.Text)
        If Check1.Value = Checked Then
            Call PutVar(App.Path & "\config.ini", "CONFIG", "Senha", txtPassword.Text)
        Else
            Call PutVar(App.Path & "\config.ini", "CONFIG", "Senha", "")
        End If
    End If
End Sub

Private Sub tmrInfo_Timer()
    lblOnOff.Caption = "O Servidor está Offline"
    lblPlayers.Visible = False
    tmrInfo.Enabled = False
End Sub
