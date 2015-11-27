VERSION 5.00
Begin VB.Form frmSobre 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sobre o sistema"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSobre.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Copyright © por Marcius C. Bezerra - 1999"
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   2805
      Width           =   5565
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "RUA BOULEVAR JOÃO BARBOSA, 1013 - CENTRO - SOBRA CE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   285
      TabIndex        =   3
      Top             =   1425
      Width           =   4980
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Todos os direitos reservados"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   2520
      Width           =   5565
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Versão 1.0"
      Height          =   255
      Left            =   45
      TabIndex        =   1
      Top             =   2220
      Width           =   5565
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "M. C. B. EDITOR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   60
      TabIndex        =   0
      Top             =   1890
      Width           =   5565
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   75
      Picture         =   "frmSobre.frx":0442
      Stretch         =   -1  'True
      Top             =   105
      Width           =   5565
   End
End
Attribute VB_Name = "frmSobre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

