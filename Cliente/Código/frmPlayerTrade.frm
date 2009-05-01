VERSION 5.00
Begin VB.Form frmPlayerTrade 
   BorderStyle     =   0  'None
   Caption         =   "Negociação"
   ClientHeight    =   3795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3435
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   3435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox Items2 
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
      Height          =   855
      Left            =   960
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.ListBox Items1 
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
      Height          =   1350
      Left            =   1920
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.ListBox PlayerInv1 
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
      Height          =   1350
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Itens para a troca"
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
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Top             =   1560
      Width           =   1325
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Seus itens"
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
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Command3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selecionar"
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
      TabIndex        =   7
      Top             =   3360
      Width           =   645
   End
   Begin VB.Label Command4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remover"
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
      Left            =   2625
      TabIndex        =   6
      Top             =   3360
      Width           =   585
   End
   Begin VB.Label Command5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Height          =   285
      Left            =   3075
      TabIndex        =   5
      Top             =   120
      Width           =   300
   End
   Begin VB.Label Command2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Troca"
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
      TabIndex        =   4
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trocar"
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
      Left            =   1560
      TabIndex        =   3
      Top             =   3360
      Width           =   405
   End
End
Attribute VB_Name = "frmPlayerTrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright (c) 2009 - Elysium Source. Alguns direitos reservados.
' Tradução e revisão por MMODEV Brasil @ http://www.mmodev.com.br
' Este código está licensiado sob a licença EGL.

Private Sub Command1_Click()
Dim Packet As String
Dim I As Long

    Packet = "swapitems" & SEP_CHAR
    For I = 1 To MAX_PLAYER_TRADES
        Packet = Packet & Trading(I).InvNum & SEP_CHAR
    Next I
    Packet = Packet & END_CHAR
    
    Call SendData(Packet)
    
    If Command1.ForeColor = &HFF00& Then
        Command1.ForeColor = &H0&
    Else
        Command1.ForeColor = &HFF00&
    End If
End Sub

Private Sub Command3_Click()
Dim I As Long, n As Long
I = PlayerInv1.ListIndex + 1

If GetPlayerInvItemNum(MyIndex, I) > 0 And GetPlayerInvItemNum(MyIndex, I) <= MAX_ITEMS Then
    For n = 1 To MAX_PLAYER_TRADES
        If Trading(n).InvNum = I Then
            MsgBox "Você só pode trocar esse item uma vez!"
            Exit Sub
        End If
        If Trading(n).InvNum <= 0 Then
            If Item(GetPlayerInvItemNum(MyIndex, I)).Type = ITEM_TYPE_CURRENCY Then
                MsgBox "Não pode trocar ouro!"
                Exit Sub
            Else
                If GetPlayerWeaponSlot(MyIndex) = I Or GetPlayerArmorSlot(MyIndex) = I Or GetPlayerHelmetSlot(MyIndex) = I Or GetPlayerShieldSlot(MyIndex) = I Then
                    MsgBox "Não pode trocar itens equipados!"
                    Exit Sub
                Else
                    PlayerInv1.List(I - 1) = PlayerInv1.Text & " **"
                    Items1.List(n - 1) = n & ": " & Trim(Item(GetPlayerInvItemNum(MyIndex, I)).Name)
                    Trading(n).InvNum = I
                    Trading(n).InvName = Trim(Item(GetPlayerInvItemNum(MyIndex, I)).Name)
                    Call SendData("updatetradeinv" & SEP_CHAR & n & SEP_CHAR & Trading(n).InvNum & SEP_CHAR & Trading(n).InvName & END_CHAR)
                    Exit Sub
                End If
            End If
        End If
    Next n
End If
End Sub

Private Sub Command4_Click()
Dim I As Long, n As Long
I = Items1.ListIndex + 1

    If Trading(I).InvNum <= 0 Then
        MsgBox "Sem item para remover!"
        Exit Sub
    End If

    PlayerInv1.List(Trading(I).InvNum - 1) = Mid(Trim(PlayerInv1.List(Trading(I).InvNum - 1)), 1, Len(PlayerInv1.List(Trading(I).InvNum - 1)) - 3)
    Items1.List(I - 1) = n & ": <Nada>"
    Trading(I).InvNum = 0
    Trading(I).InvName = vbNullString
    Call SendData("updatetradeinv" & SEP_CHAR & I & SEP_CHAR & 0 & SEP_CHAR & vbNullString & END_CHAR)
    Command1.ForeColor = &H80000012
End Sub

Private Sub Command5_Click()
    Call SendData("qtrade" & END_CHAR)
End Sub

Private Sub Form_Load()
Dim I As Long
Dim Ending As String
    For I = 1 To 3
        If I = 1 Then Ending = ".gif"
        If I = 2 Then Ending = ".jpg"
        If I = 3 Then Ending = ".png"
 
        If FileExist("GUI\content" & Ending) Then frmPlayerTrade.Picture = LoadPicture(App.Path & "\GUI\content" & Ending)
    Next I
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Command1.ForeColor = &H0&
    Command2.ForeColor = &H0&
End Sub

