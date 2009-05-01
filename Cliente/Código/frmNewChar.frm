VERSION 5.00
Begin VB.Form frmNewChar 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Elysium Engine Brasil"
   ClientHeight    =   3780
   ClientLeft      =   90
   ClientTop       =   -60
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
   Begin VB.OptionButton optFemale 
      BackColor       =   &H00FF8080&
      Caption         =   "Mulher"
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
      Height          =   195
      Left            =   1320
      TabIndex        =   19
      Top             =   1440
      Width           =   855
   End
   Begin VB.OptionButton optMale 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      Caption         =   "Homem"
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
      Height          =   195
      Left            =   360
      TabIndex        =   18
      Top             =   1440
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   2520
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   8
      Top             =   960
      Width           =   555
      Begin VB.PictureBox Picpic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   15
         ScaleHeight     =   31
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   31
         TabIndex        =   9
         Top             =   15
         Width           =   495
      End
   End
   Begin VB.PictureBox Picsprites 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3240
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   3600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer2 
      Interval        =   250
      Left            =   240
      Top             =   240
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   720
      Top             =   240
   End
   Begin VB.ComboBox cmbClass 
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
      Height          =   285
      Left            =   360
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2040
      Width           =   2835
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
      Left            =   360
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1200
      Width           =   1800
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Atributos Iniciais:"
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
      TabIndex        =   26
      Top             =   2400
      Width           =   1065
   End
   Begin VB.Label picAddChar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Criar o Personagem"
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
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   3225
   End
   Begin VB.Label lblSPEED 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   1920
      TabIndex        =   25
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "AGI"
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
      Left            =   1560
      TabIndex        =   24
      Top             =   2880
      Width           =   540
   End
   Begin VB.Label lblDEF 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   1320
      TabIndex        =   23
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "DEF"
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
      Left            =   960
      TabIndex        =   22
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lblSTR 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   720
      TabIndex        =   21
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "FOR"
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
      Left            =   360
      TabIndex        =   20
      Top             =   2880
      Width           =   360
   End
   Begin VB.Label lblMAGI 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   2520
      TabIndex        =   17
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "INT"
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
      Left            =   2160
      TabIndex        =   16
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label lblSP 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   1800
      TabIndex        =   15
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "SP"
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
      Left            =   1560
      TabIndex        =   14
      Top             =   2640
      Width           =   300
   End
   Begin VB.Label lblMP 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   1200
      TabIndex        =   13
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "MP"
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
      Left            =   960
      TabIndex        =   12
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label lblHP 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Height          =   345
      Left            =   600
      TabIndex        =   11
      Top             =   2640
      Width           =   405
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "HP"
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
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   2640
      Width           =   240
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Classe:"
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
      Height          =   210
      Left            =   360
      TabIndex        =   7
      Top             =   1800
      Width           =   885
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Personagem:"
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
      Height          =   210
      Left            =   360
      TabIndex        =   6
      Top             =   960
      Width           =   1755
   End
   Begin VB.Label picCancel 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00808080&
      Height          =   345
      Left            =   3075
      TabIndex        =   4
      Top             =   120
      Width           =   285
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Sex"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6960
      TabIndex        =   2
      Top             =   2190
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmNewChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright (c) 2009 - Elysium Source. Alguns direitos reservados.
' Tradução e revisão por MMODEV Brasil @ http://www.mmodev.com.br
' Este código está licensiado sob a licença EGL.

Option Explicit
Public animi As Long

Private Sub cmbClass_Click()
    lblHP.Caption = STR(Class(cmbClass.ListIndex + 1).HP)
    lblMP.Caption = STR(Class(cmbClass.ListIndex + 1).MP)
    lblSP.Caption = STR(Class(cmbClass.ListIndex + 1).SP)
    
    lblSTR.Caption = STR(Class(cmbClass.ListIndex + 1).STR)
    lblDEF.Caption = STR(Class(cmbClass.ListIndex + 1).DEF)
    lblSPEED.Caption = STR(Class(cmbClass.ListIndex + 1).Speed)
    lblMAGI.Caption = STR(Class(cmbClass.ListIndex + 1).MAGI)
End Sub


Private Sub picAddChar_Click()
Dim Msg As String
Dim I As Long

    If Trim(txtName.Text) <> vbNullString Then
        Msg = Trim(txtName.Text)
        
        If Len(Trim(txtName.Text)) < 3 Then
            MsgBox "Nomes de personagens precisam ser de no mínimo 3 caracteres."
            Exit Sub
        End If
        
        ' Prevent high ascii chars
        For I = 1 To Len(Msg)
            If Asc(Mid(Msg, I, 1)) < 32 Or Asc(Mid(Msg, I, 1)) > 126 Then
                Call MsgBox("Você não pode usar acentos no nome do seu personagem.", vbOKOnly, GAME_NAME)
                txtName.Text = vbNullString
                Exit Sub
            End If
        Next I
        
        Call MenuState(MENU_STATE_ADDCHAR)
    End If
End Sub

Private Sub picCancel_Click()
    Me.Hide
    frmChars.Show , frmMainMenu
End Sub

Private Sub Timer1_Timer()
    If cmbClass.ListIndex < 0 Then Exit Sub
    
    Picpic.Width = SIZE_X
    Picpic.Height = SIZE_Y
    Picture4.Width = SIZE_X + 4
    Picture4.Height = SIZE_Y + 4
    
    If optMale.Value = True Then
        Call BitBlt(Picpic.hDC, 0, 0, SIZE_X, SIZE_Y, picSprites.hDC, animi * SIZE_X, Int(Class(cmbClass.ListIndex + 1).MaleSprite) * SIZE_Y - (SIZE_Y - PIC_Y), SRCCOPY)
    Else
        Call BitBlt(Picpic.hDC, 0, 0, SIZE_X, SIZE_Y, picSprites.hDC, animi * SIZE_X, Int(Class(cmbClass.ListIndex + 1).FemaleSprite) * SIZE_Y - (SIZE_Y - PIC_Y), SRCCOPY)
    End If
End Sub

Private Sub Form_Load()
Dim I As Long
Dim Ending As String
    For I = 1 To 3
        If I = 1 Then Ending = ".gif"
        If I = 2 Then Ending = ".jpg"
        If I = 3 Then Ending = ".png"
 
        If FileExist("GUI\medium" & Ending) Then frmNewChar.Picture = LoadPicture(App.Path & "\GUI\medium" & Ending)
    Next I
    picSprites.Picture = LoadPicture(App.Path & "\GFX\sprites.bmp")
End Sub

Private Sub Timer2_Timer()
    animi = animi + 1
If animi > 4 Then
    animi = 3
End If
End Sub
