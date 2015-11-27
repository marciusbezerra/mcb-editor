VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmOffSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Off Set"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2415
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOffSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   2415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   315
      Left            =   645
      TabIndex        =   1
      Top             =   615
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      BuddyControl    =   "txtOff"
      BuddyDispid     =   196609
      OrigLeft        =   990
      OrigTop         =   180
      OrigRight       =   1230
      OrigBottom      =   345
      Max             =   100
      Min             =   -100
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtOff 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Text            =   "0"
      Top             =   615
      Width           =   525
   End
   Begin VB.Label Label1 
      Caption         =   "Selecione valores entre -1000 e 1000:"
      Height          =   405
      Left            =   105
      TabIndex        =   2
      Top             =   75
      Width           =   1650
   End
End
Attribute VB_Name = "frmOffSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub txtOff_Change()
    On Error GoTo Erro
    If Trim(Me.txtOff.Text) = "" Then Me.txtOff.Text = "0"
    If IsNumeric(Me.txtOff.Text) Then
        If CDbl(Me.txtOff.Text) <= 1000 And CDbl(Me.txtOff.Text) >= -1000 Then
            frmMain.txtRTF.SelCharOffset = Me.txtOff
        Else
            MsgBox "Apenas valores entre -1000 e 1000.", vbSystemModal + vbCritical, Caption
        End If
    Else
        MsgBox "Apenas valores numéricos.", vbSystemModal + vbCritical, Caption
    End If
    Exit Sub
Erro:
    MsgBox Error$, vbCritical, "ERRO"
End Sub
