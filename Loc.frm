VERSION 5.00
Begin VB.Form frmLoc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Localizar / Substituir"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   Icon            =   "Loc.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   125
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   453
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1545
      MaxLength       =   50
      TabIndex        =   7
      Top             =   570
      Width           =   5115
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1545
      MaxLength       =   50
      TabIndex        =   6
      Top             =   150
      Width           =   5115
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Somente palavras inteiras"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   90
      TabIndex        =   5
      Top             =   1470
      Width           =   2265
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Diferenciar maiúsculas e minúsculas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   90
      TabIndex        =   4
      Top             =   1170
      Width           =   3060
   End
   Begin VB.CommandButton cmdSubstituiTodas 
      Caption         =   "Substituir &todas"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4815
      TabIndex        =   3
      Top             =   1395
      Width           =   1890
   End
   Begin VB.CommandButton cmdSubstitui 
      Caption         =   "S&ubstituir"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3465
      TabIndex        =   2
      Top             =   1410
      Width           =   1215
   End
   Begin VB.CommandButton cmdLocalizaProximo 
      Caption         =   "Localizar &próximo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4815
      TabIndex        =   1
      Top             =   1005
      Width           =   1890
   End
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "L&ocalizar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3450
      TabIndex        =   0
      Top             =   1005
      Width           =   1230
   End
   Begin VB.Label Label2 
      Caption         =   "&Substituir por:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   9
      Top             =   660
      Width           =   1485
   End
   Begin VB.Label Label1 
      Caption         =   "&Localizar o que:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   150
      TabIndex        =   8
      Top             =   195
      Width           =   1410
   End
End
Attribute VB_Name = "frmLoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Posição As Long

Private Sub cmdLocalizaProximo_Click()
Dim Param As Integer

Param = Check1.Value * 4 + Check2.Value * 2
Posição = frmMain.txtRTF.Find(Text1.Text, Posição + 1, , Param)
If Posição > 0 Then
    frmMain.SetFocus
Else
    MsgBox "Não há mais ocorrências.", vbSystemModal + vbInformation, Caption
    Me.cmdSubstitui.Enabled = False
    Me.cmdSubstituiTodas.Enabled = False
End If

End Sub

Private Sub cmdLocalizar_Click()
Dim Param As Integer

    Posição = 0
    Param = Check1.Value * 4 + Check2.Value * 2
    Posição = frmMain.txtRTF.Find(Text1.Text, Posição + 1, , Param)
    If Posição >= 0 Then
        Me.cmdSubstitui.Enabled = True
        Me.cmdSubstituiTodas.Enabled = True
        frmMain.SetFocus
    Else
        MsgBox "Texto não localizado.", vbSystemModal + vbInformation, Caption
        Me.cmdSubstitui.Enabled = False
        Me.cmdSubstituiTodas.Enabled = False
    End If
End Sub

Private Sub cmdSubstitui_Click()
Dim Param As Integer
    If Not IsNull(frmMain.txtRTF.SelText) Then
        If UCase(frmMain.txtRTF.SelText) = UCase(Me.Text1.Text) Then
            frmMain.txtRTF.SelText = Text2.Text
            Posição = Posição + Len(Text2.Text)
        End If
    End If
    Param = Check1.Value * 4 + Check2.Value * 2
    Posição = frmMain.txtRTF.Find(Text1.Text, Posição + 1, , Param)
    If Posição > 0 Then
        frmMain.SetFocus
    Else
        MsgBox "Texto não localizado.", vbSystemModal + vbInformation, Caption
        Me.cmdSubstitui.Enabled = False
        Me.cmdSubstituiTodas.Enabled = False
    End If
End Sub

Private Sub cmdSubstituiTodas_Click()
Dim Param As Integer

    Param = Check1.Value * 4 + Check2.Value * 2
    If Not IsNull(frmMain.txtRTF.SelText) Then
        If UCase(frmMain.txtRTF.SelText) = UCase(Me.Text1.Text) Then
            frmMain.txtRTF.SelText = Text2.Text
            Posição = Posição + Len(Text2.Text)
        End If
    End If
    Posição = frmMain.txtRTF.Find(Text1.Text, Posição + 1, , Param)
    While Posição > 0
        frmMain.txtRTF.SelText = Text2.Text
        Posição = Posição + Len(Text2.Text)
        Posição = frmMain.txtRTF.Find(Text1.Text, Posição + 1, , Param)
    Wend
    Me.cmdSubstitui.Enabled = False
    Me.cmdSubstituiTodas.Enabled = False
    MsgBox "Substituição finalizada.", vbSystemModal + vbInformation, Caption
End Sub
