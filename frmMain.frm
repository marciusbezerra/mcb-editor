VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "M.C.B. Editor - Desconhecido"
   ClientHeight    =   6390
   ClientLeft      =   -105
   ClientTop       =   795
   ClientWidth     =   9690
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   9690
   Begin MSComDlg.CommonDialog Dialogo 
      Left            =   3135
      Top             =   3795
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   5
      Top             =   6090
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "M. C. B. Editor"
            TextSave        =   "M. C. B. Editor"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "12/01/2000"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "10:52"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Bevel           =   2
            Enabled         =   0   'False
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "NUM"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSld 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   9690
      TabIndex        =   2
      Top             =   360
      Width           =   9690
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   -15
         Width           =   9315
         _ExtentX        =   16431
         _ExtentY        =   450
         _Version        =   393216
         Max             =   200
      End
      Begin MSComctlLib.Slider Slider2 
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   255
         Width           =   9315
         _ExtentX        =   16431
         _ExtentY        =   450
         _Version        =   393216
         Max             =   200
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1530
      Top             =   5130
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0556
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":066A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":077E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0892
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":09A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0ABA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0BCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0DF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":101E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1132
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1246
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":135A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   21
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NOVO"
            Object.ToolTipText     =   "Novo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ABRIR"
            Object.ToolTipText     =   "Abrir"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SALVA"
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "IMPRIMIR"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "COP"
            Object.ToolTipText     =   "Copiar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "REC"
            Object.ToolTipText     =   "Recortar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "COL"
            Object.ToolTipText     =   "Colar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FIND"
            Object.ToolTipText     =   "Localizar / Substituir"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NEG"
            Object.ToolTipText     =   "Negrito"
            ImageIndex      =   9
            Style           =   1
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ITAL"
            Object.ToolTipText     =   "Itálico"
            ImageIndex      =   10
            Style           =   1
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SUB"
            Object.ToolTipText     =   "Sublinhado"
            ImageIndex      =   11
            Style           =   1
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "RISC"
            Object.ToolTipText     =   "Riscado"
            ImageIndex      =   12
            Style           =   1
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "MARCA"
            Object.ToolTipText     =   "Marcadores"
            ImageIndex      =   15
            Style           =   1
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "HELP"
            Object.ToolTipText     =   "Créditos"
            ImageIndex      =   13
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtRTF 
      Height          =   4620
      Left            =   45
      TabIndex        =   0
      Top             =   930
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   8149
      _Version        =   393217
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      RightMargin     =   4989,764
      TextRTF         =   $"frmMain.frx":17AE
   End
   Begin VB.Menu Arquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu ArquivoNovo 
         Caption         =   "&Novo"
      End
      Begin VB.Menu ArquivoAbrir 
         Caption         =   "&Abrir"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu ArquivoSalvar 
         Caption         =   "&Salvar"
         Shortcut        =   ^B
      End
      Begin VB.Menu ArquivoSalvarComo 
         Caption         =   "Salvar &como ..."
      End
      Begin VB.Menu mnu2 
         Caption         =   "-"
      End
      Begin VB.Menu ArquivoImprimir 
         Caption         =   "&Imprimir"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnu3 
         Caption         =   "-"
      End
      Begin VB.Menu ArqRecentes 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu ArqRecentes 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu ArqRecentes 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu ArqRecentes 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnu7 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu ArquivoSair 
         Caption         =   "Sai&r"
      End
   End
   Begin VB.Menu Editar 
      Caption         =   "&Editar"
      Begin VB.Menu EditarCopiar 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu EditarRecortar 
         Caption         =   "&Recortar"
         Shortcut        =   ^R
      End
      Begin VB.Menu EditarColar 
         Caption         =   "C&olar"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnu4 
         Caption         =   "-"
      End
      Begin VB.Menu EditarSelecionarTudo 
         Caption         =   "&Selecionar tudo"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnu5 
         Caption         =   "-"
      End
      Begin VB.Menu EditarLocalizarSubstitur 
         Caption         =   "&Localizar / Substitur"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu Formatar 
      Caption         =   "&Formatar"
      Begin VB.Menu FormatarFonte 
         Caption         =   "&Fonte"
      End
      Begin VB.Menu mnu6 
         Caption         =   "-"
      End
      Begin VB.Menu FormatarNegrito 
         Caption         =   "&Negrito"
         Shortcut        =   ^N
      End
      Begin VB.Menu FormatarItálico 
         Caption         =   "&Itálico"
         Shortcut        =   ^I
      End
      Begin VB.Menu FormatarSublinhado 
         Caption         =   "&Sublinhado"
         Shortcut        =   ^S
      End
      Begin VB.Menu FormatarRiscado 
         Caption         =   "&Riscado"
      End
      Begin VB.Menu FormatarOffSet 
         Caption         =   "&Off Set"
      End
      Begin VB.Menu mnuMarcadores 
         Caption         =   "&Marcadores"
      End
   End
   Begin VB.Menu Ajuda 
      Caption         =   "Aj&uda"
      Begin VB.Menu AjudaCréditos 
         Caption         =   "&Créditos"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AjudaCréditos_Click()
    frmSobre.Show vbModal
End Sub

Private Sub ArqRecentes_Click(Index As Integer)
    Dim Arquivo As String
    Dim MSG As String
    Dim MSGResp As Integer
    Dim Resposta As Boolean
    If FStado.Modificado = True Then
        Arquivo = Right(frmMain.Caption, Len(frmMain.Caption) - 16)
        If UCase(ArqRecentes(Index).Caption) = UCase(Arquivo) Then Exit Sub
        MSG = "O arquivo [" & Arquivo & "] foi modificado."
        MSG = MSG & vbCrLf
        MSG = MSG & "Salvar as modificações ?"
        MSGResp = MsgBox(MSG, vbQuestion + vbYesNoCancel, frmMain.Caption)
        Select Case MSGResp
            Case vbYes
                Resposta = SalvarArquivo
                If Resposta = False Then Exit Sub
            Case vbCancel
                Exit Sub
        End Select
    End If
    If UCase(ArqRecentes(Index).Caption) = UCase(Arquivo) Then Exit Sub
    
    If Right(ArqRecentes(Index).Caption, 3) = "RTF" Then
        TipoSalvar = rtfRTF
    Else
        TipoSalvar = rtfText
    End If
    
    AbrirArquivo ArqRecentes(Index).Caption
    ArquivosRecentes
End Sub

Private Sub ArquivoAbrir_Click()
    AbrirArquivoProc
End Sub

Private Sub ArquivoImprimir_Click()
    Dim UltPos As Long
    Dim UltSel As Long
    
    On Error GoTo Erro
    
    UltPos = Me.txtRTF.SelStart
    UltSel = Me.txtRTF.SelLength
    
    Me.Dialogo.Flags = cdlPDPrintSetup Or cdlPDReturnDC
    Me.Dialogo.CancelError = True
    Me.Dialogo.ShowPrinter
    
    EditarSelecionarTudo_Click
    
    Mens "Aguarde, Imprimindo ..."
    
    txtRTF.SelPrint Me.Dialogo.hDC
    
    Mens "Pronto"
    
    txtRTF.SelStart = UltPos
    txtRTF.SelLength = UltSel
    
    Exit Sub
Erro:
    If Err = cdlCancel Then
        Exit Sub
    Else
        MsgBox Error$, vbCritical, "ERRO"
        Mens "Pronto"
        Exit Sub
    End If
    
End Sub

Private Sub ArquivoNovo_Click()
    NovoArquivo
End Sub

Private Sub ArquivoSair_Click()
    Unload Me
End Sub

Private Sub ArquivoSalvar_Click()
    SalvarArquivo
End Sub

Private Sub ArquivoSalvarComo_Click()
    Dim ArquivoSalvo As String
    Dim NomePadrao As String
    
    NomePadrao = Right$(Caption, Len(Caption) - 16)
    If Me.Caption = "M.C.B. Editor - Desconhecido" Then
        ArquivoSalvo = NomeDoArquivo("Doc.RTF")
        If Trim(ArquivoSalvo) <> "" Then SalvarArquivoComo ArquivoSalvo
        AjeitaMenu ArquivoSalvo
    Else
        ArquivoSalvo = NomeDoArquivo(NomePadrao)
        If Trim(ArquivoSalvo) <> "" Then SalvarArquivoComo ArquivoSalvo
        AjeitaMenu ArquivoSalvo
        AbrirArquivo ArquivoSalvo
    End If
End Sub

Private Sub EditarColar_Click()
    On Error GoTo Erro
    If Clipboard.GetFormat(vbCFRTF) Then
        Me.txtRTF.SelRTF = Clipboard.GetText(vbCFRTF)
    ElseIf Clipboard.GetFormat(vbCFText) Then
        Me.txtRTF.SelText = Clipboard.GetText
    End If
    Exit Sub
Erro:
    MsgBox Error$, vbCritical, "ERRO"
End Sub

Private Sub EditarCopiar_Click()
    Clipboard.Clear
    Clipboard.SetText Me.txtRTF.SelRTF, vbCFRTF
End Sub

Private Sub EditarLocalizarSubstitur_Click()
    SempreVisível frmLoc, CurrentX, CurrentY, 470, 155
End Sub

Private Sub EditarRecortar_Click()
    EditarCopiar_Click
    Me.txtRTF.SelText = ""
End Sub

Private Sub EditarSelecionarTudo_Click()
    On Error GoTo Erro
    Me.txtRTF.SelStart = 0
    Me.txtRTF.SelLength = Len(Me.txtRTF.Text)
    Exit Sub
Erro:
    MsgBox Error$, vbCritical, "ERRO"
End Sub

Private Sub Form_Load()
    Dim I As Integer
    Show
    ChDir App.Path
    FStado.Modificado = False
    Me.ArquivoSalvar.Enabled = False
    Me.Toolbar1.Buttons(4).Enabled = False
    ArquivosRecentes
    AtribuiParametrosDoForm Me
    Me.EditarCopiar.Enabled = False
    Me.EditarRecortar.Enabled = False
    Me.Toolbar1.Buttons(8).Enabled = False
    Me.Toolbar1.Buttons(9).Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim MSG As String
    Dim Arquivo As String
    Dim Resposta As Integer

    If FStado.Modificado Then
        Arquivo = Right(Me.Caption, Len(Me.Caption) - 16)
        MSG = "O arquivo [" & Arquivo & "] foi modificado."
        MSG = MSG & vbCrLf
        MSG = MSG & "Salvar as modificações ?"
        Resposta = MsgBox(MSG, vbQuestion + vbYesNoCancel, Caption)
        Select Case Resposta
            Case vbYes
                If Right(Caption, 12) = "Desconhecido" Then
                    Arquivo = "Doc.rtf"
                    Arquivo = NomeDoArquivo(Arquivo)
                Else
                    Arquivo = Right(Me.Caption, Len(Me.Caption) - 16)
                End If
                If Arquivo <> "" Then
                    txtRTF.SaveFile Arquivo, TipoSalvar
                End If
            Case vbNo
                Cancel = False
            Case vbCancel
                Cancel = True
        End Select
    End If
    GravaParametrosDoForm Me
End Sub

Private Sub Form_Resize()
    CentraTexto
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ArquivosRecentes
End Sub

Private Sub FormatarFonte_Click()
    On Error GoTo Erro
    If Not IsNull(txtRTF.SelFontName) Then
        Me.Dialogo.FontName = txtRTF.SelFontName
    End If
    If Not IsNull(txtRTF.SelFontSize) Then
        Me.Dialogo.FontSize = txtRTF.SelFontSize
    End If
    If Not IsNull(txtRTF.SelColor) Then
        Me.Dialogo.Color = txtRTF.SelColor
    End If
    If Not IsNull(txtRTF.SelBold) Then
        Me.Dialogo.FontBold = txtRTF.SelBold
    End If
    If Not IsNull(txtRTF.SelItalic) Then
        Me.Dialogo.FontItalic = txtRTF.SelItalic
    End If
    If Not IsNull(txtRTF.SelUnderline) Then
        Me.Dialogo.FontUnderline = txtRTF.SelUnderline
    End If
    If Not IsNull(txtRTF.SelStrikeThru) Then
        Me.Dialogo.FontStrikethru = txtRTF.SelStrikeThru
    End If
    Me.Dialogo.Flags = cdlCFBoth Or cdlCFEffects
    Me.Dialogo.CancelError = True
    Me.Dialogo.ShowFont
    txtRTF.SelFontName = Me.Dialogo.FontName
    txtRTF.SelFontSize = Me.Dialogo.FontSize
    txtRTF.SelColor = Me.Dialogo.Color
    txtRTF.SelBold = Me.Dialogo.FontBold
    txtRTF.SelItalic = Me.Dialogo.FontItalic
    txtRTF.SelUnderline = Me.Dialogo.FontUnderline
    txtRTF.SelStrikeThru = Me.Dialogo.FontStrikethru
    Exit Sub
Erro:
    If Err = cdlCancel Then
        Exit Sub
    Else
        MsgBox Err.Description, vbCritical, "ERRO"
        Exit Sub
    End If
End Sub

Private Sub FormatarItálico_Click()
    CheckComandos Me.FormatarItálico, 15
    Me.txtRTF.SelItalic = Me.FormatarItálico.Checked
End Sub

Private Sub FormatarNegrito_Click()
    CheckComandos Me.FormatarNegrito, 14
    Me.txtRTF.SelBold = Me.FormatarNegrito.Checked
End Sub

Private Sub FormatarOffSet_Click()
    On Error GoTo Erro
    If Not IsNull(Me.txtRTF.SelCharOffset) Then
        frmOffSet.txtOff.Text = Me.txtRTF.SelCharOffset
    End If
    SempreVisível frmOffSet, Me.CurrentX, Me.CurrentY, 200, 100
    Exit Sub
Erro:
    MsgBox Error$, vbCritical, "ERRO"
End Sub

Private Sub FormatarRiscado_Click()
    CheckComandos Me.FormatarRiscado, 17
    Me.txtRTF.SelStrikeThru = Me.FormatarRiscado.Checked
End Sub

Private Sub FormatarSublinhado_Click()
    CheckComandos Me.FormatarSublinhado, 16
    Me.txtRTF.SelUnderline = Me.FormatarSublinhado.Checked
End Sub

Private Sub mnuMarcadores_Click()
    CheckComandos Me.mnuMarcadores, 19
    Me.txtRTF.SelBullet = mnuMarcadores.Checked
End Sub

Private Sub Slider1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If frmMain.Slider1.Width <> frmMain.picSld.Width Then
        CentraTexto
    End If
End Sub

Private Sub Slider2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If frmMain.Slider1.Width <> frmMain.picSld.Width Then
        CentraTexto
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "NOVO": ArquivoNovo_Click
        Case "ABRIR": ArquivoAbrir_Click
        Case "SALVA": ArquivoSalvar_Click
        Case "IMPRIMIR": ArquivoImprimir_Click
        Case "COP": EditarCopiar_Click
        Case "REC": EditarRecortar_Click
        Case "COL": EditarColar_Click
        Case "FIND": EditarLocalizarSubstitur_Click
        Case "NEG": FormatarNegrito_Click
        Case "ITAL": FormatarItálico_Click
        Case "SUB": FormatarSublinhado_Click
        Case "RISC": FormatarRiscado_Click
        Case "MARCA": mnuMarcadores_Click
        Case "HELP": AjudaCréditos_Click
    End Select
End Sub

Private Sub Toolbar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If frmMain.Slider1.Width <> frmMain.picSld.Width Then
        CentraTexto
    End If
End Sub

Private Sub txtRTF_Change()
    FStado.Modificado = True
    Me.ArquivoSalvar.Enabled = True
    Me.Toolbar1.Buttons(4).Enabled = True
End Sub

Private Sub txtRTF_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If frmMain.Slider1.Width <> frmMain.picSld.Width Then
        CentraTexto
    End If
    Mens "Área de edição"
End Sub

Private Sub txtRTF_SelChange()
    If Me.txtRTF.SelLength = 0 Then
        Me.EditarCopiar.Enabled = False
        Me.EditarRecortar.Enabled = False
        Me.Toolbar1.Buttons(8).Enabled = False
        Me.Toolbar1.Buttons(9).Enabled = False
    Else
        Me.EditarCopiar.Enabled = True
        Me.EditarRecortar.Enabled = True
        Me.Toolbar1.Buttons(8).Enabled = True
        Me.Toolbar1.Buttons(9).Enabled = True
    End If
    If Not IsNull(txtRTF.SelBold) Then
        Me.FormatarNegrito.Checked = Me.txtRTF.SelBold
        If Me.FormatarNegrito.Checked Then
            Me.Toolbar1.Buttons(14).Value = tbrPressed
        Else
            Me.Toolbar1.Buttons(14).Value = tbrUnpressed
        End If
    End If
    If Not IsNull(txtRTF.SelItalic) Then
        Me.FormatarItálico.Checked = Me.txtRTF.SelItalic
        If Me.FormatarItálico.Checked Then
            Me.Toolbar1.Buttons(15).Value = tbrPressed
        Else
            Me.Toolbar1.Buttons(15).Value = tbrUnpressed
        End If
    End If
    If Not IsNull(txtRTF.SelUnderline) Then
        Me.FormatarSublinhado.Checked = Me.txtRTF.SelUnderline
        If Me.FormatarSublinhado.Checked Then
            Me.Toolbar1.Buttons(16).Value = tbrPressed
        Else
            Me.Toolbar1.Buttons(16).Value = tbrUnpressed
        End If
    End If
    If Not IsNull(txtRTF.SelStrikeThru) Then
        Me.FormatarRiscado.Checked = Me.txtRTF.SelStrikeThru
        If Me.FormatarRiscado.Checked Then
            Me.Toolbar1.Buttons(17).Value = tbrPressed
        Else
            Me.Toolbar1.Buttons(17).Value = tbrUnpressed
        End If
    End If
    If Not IsNull(txtRTF.SelBullet) Then
        Me.mnuMarcadores.Checked = Me.txtRTF.SelBullet
        If Me.mnuMarcadores.Checked Then
            Me.Toolbar1.Buttons(19).Value = tbrPressed
        Else
            Me.Toolbar1.Buttons(19).Value = tbrUnpressed
        End If
    End If
    If IsNull(txtRTF.SelIndent) Then
        Slider1.Enabled = False
        Slider2.Enabled = False
        Exit Sub
    Else
        Slider1.Enabled = True
        Slider2.Enabled = True
On Error Resume Next
        Slider1.Value = txtRTF.SelIndent * Slider1.Max / txtRTF.RightMargin
        Slider2.Value = (txtRTF.SelHangingIndent / txtRTF.RightMargin) * Slider2.Max + Slider1.Value
    End If
End Sub

Private Sub Slider1_Scroll()
    txtRTF.SelIndent = txtRTF.RightMargin * (Slider1.Value / Slider1.Max)
    Slider2_Scroll
End Sub


Private Sub Slider2_Scroll()
    txtRTF.SelHangingIndent = txtRTF.RightMargin * ((Slider2.Value - Slider1.Value) / Slider2.Max)
End Sub
