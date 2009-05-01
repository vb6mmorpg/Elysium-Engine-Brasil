VERSION 5.00
Begin VB.Form frmIpconfig 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Elysium Engine Brasil"
   ClientHeight    =   3780
   ClientLeft      =   90
   ClientTop       =   -60
   ClientWidth     =   3435
   ControlBox      =   0   'False
   Icon            =   "frmIpconfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   3435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtPort 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   2040
      Width           =   2775
   End
   Begin VB.TextBox TxtIP 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label PicCancel 
      BackStyle       =   0  'Transparent
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
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.Label PicConfirm 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmar"
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
      Left            =   1425
      TabIndex        =   4
      Top             =   2880
      Width           =   645
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "IP do Servidor:"
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
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   960
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Porta do Servidor:"
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
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   1140
   End
End
Attribute VB_Name = "frmIpconfig"
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
 
        If FileExist("GUI\medium" & Ending) Then frmIpconfig.Picture = LoadPicture(App.Path & "\GUI\medium" & Ending)
    Next I
    
    Dim FileName As String
    FileName = App.Path & "\config.ini"
    TxtIP = GetVar(FileName, "IPCONFIG", "IP")
    TxtPort = GetVar(FileName, "IPCONFIG", "Porta")
    TxtIP.Text = GetVar(FileName, "IPCONFIG", "IP")
    TxtPort.Text = GetVar(FileName, "IPCONFIG", "Porta")
End Sub

Private Sub picCancel_Click()
    frmMainMenu.Visible = True
    frmIpconfig.Visible = False
End Sub

Private Sub picConfirm_Click()
    Dim IP, Port As String
    Dim fErr As Integer
    Dim Texto As String
    Dim Packet As String
        
    IP = TxtIP
    Port = Val(TxtPort)

    fErr = 0
    If fErr = 0 And Len(Trim(IP)) = 0 Then
        fErr = 1
        Call MsgBox("IP inválido!", vbCritical, GAME_NAME)
        Exit Sub
    End If
    If fErr = 0 And Port <= 0 Then
        fErr = 1
        Call MsgBox("Porta inválida!", vbCritical, GAME_NAME)
        Exit Sub
    End If
    If fErr = 0 Then
        ' Gravar IP e Porta
        Call PutVar(App.Path & "\config.ini", "IPCONFIG", "IP", TxtIP.Text)
        Call PutVar(App.Path & "\config.ini", "IPCONFIG", "Porta", TxtPort.Text)
        'Call MenuState(MENU_STATE_IPCONFIG)
    End If
    frmMirage.Socket.Close
    frmMirage.Socket.RemoteHost = TxtIP.Text
    frmMirage.Socket.RemotePort = TxtPort.Text
    frmMainMenu.Visible = True
    frmIpconfig.Visible = False

    If frmLogin.tmrInfo.Enabled = True Then frmLogin.tmrInfo.Enabled = False

    frmLogin.lblPlayers.Visible = True
    frmLogin.lblPlayers.Caption = "Recebendo informação..."
    
    If ConnectToServer = True Then
        frmLogin.tmrInfo.Enabled = True
        Packet = "getinfo" & END_CHAR
        Call SendData(Packet)
    Else
        frmLogin.lblOnOff.Caption = "Offline"
        frmLogin.lblPlayers.Visible = False
    End If
End Sub
