Attribute VB_Name = "basMCBPa"
Option Explicit

Public TipoSalvar

Private Const HWND_TOPMOST = -1
Private Const SWP_SHOWWINDOW = &H40

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Type EstadoForm
    Deletado As Boolean
    Modificado As Boolean
    Cor As Long
End Type

Public FStado As EstadoForm

Sub CentraTexto()
    On Error GoTo Fim
    frmMain.txtRTF.Height = frmMain.ScaleHeight - (frmMain.picSld.Height + frmMain.Toolbar1.Height)
    frmMain.txtRTF.Width = frmMain.ScaleWidth
    frmMain.txtRTF.Top = (frmMain.picSld.Height + frmMain.Toolbar1.Height)
    frmMain.txtRTF.Left = 0
    frmMain.txtRTF.RightMargin = frmMain.Slider1.Width
Fim:
End Sub

Public Sub ArquivosRecentes()
    Dim I As Integer
    
    If Trim(LerIni("Recentes", "Arquivo1", "")) = "" Then Exit Sub
    
    frmMain.mnu7.Visible = True
    
    For I = 0 To 3
        If Trim(LerIni("Recentes", "Arquivo" & Trim(Str(I + 1)), "")) <> "" Then
            frmMain.ArqRecentes(I).Caption = Trim(LerIni("Recentes", "Arquivo" & Trim(Str(I + 1)), ""))
            frmMain.ArqRecentes(I).Visible = True
        End If
    Next I
End Sub

Function EscreverIni(Secao As String, Chave As String, Valor As String)
     WritePrivateProfileString Secao, Chave, Valor, App.EXEName & ".ini"
End Function

Function LerIni(Secao As String, Chave As String, Padrao As String) As String
    Dim Valor As String
    Dim Tamanho As Long
    Valor = String(255, " ")
    Tamanho = GetPrivateProfileString(Secao, Chave, _
    Padrao, Valor, 255, App.EXEName & ".ini")
    LerIni = Left(Valor, Tamanho)
End Function

Public Function NomeDoArquivo(NomeArquivo As Variant)
    On Error Resume Next
    frmMain.Dialogo.CancelError = True
    frmMain.Dialogo.DefaultExt = "RTF"
    frmMain.Dialogo.Filter = "Arquivos RTF|*.RTF|Arquivos de texto|*.TXT|Todos os arquivos|*.*"
    frmMain.Dialogo.FileName = NomeArquivo
    frmMain.Dialogo.Flags = cdlOFNOverwritePrompt
    frmMain.Dialogo.ShowSave
    If Err <> cdlCancel Then
        If frmMain.Dialogo.FilterIndex = 1 Then
            TipoSalvar = rtfRTF
        Else
            TipoSalvar = rtfText
        End If
        NomeDoArquivo = frmMain.Dialogo.FileName
    Else
        NomeDoArquivo = ""
    End If
End Function

Public Sub NovoArquivo()
    Dim MSGResp As Integer
    Dim Resposta As Boolean
    Dim Arquivo As String
    Dim MSG As String
    
    If FStado.Modificado = True Then
        Arquivo = Right(frmMain.Caption, Len(frmMain.Caption) - 16)
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
    frmMain.txtRTF.Text = ""
    frmMain.Caption = "M.C.B. Editor - Desconhecido"
    FStado.Modificado = False
    frmMain.ArquivoSalvar.Enabled = False
    frmMain.Toolbar1.Buttons(4).Enabled = False
End Sub

Public Function SalvarArquivo() As Boolean
    Dim Arquivo As String

    If frmMain.Caption = "M.C.B. Editor - Desconhecido" Then
        Arquivo = NomeDoArquivo(Arquivo)
    Else
        Arquivo = Right(frmMain.Caption, Len(frmMain.Caption) - 16)
    End If
    If Trim(Arquivo) <> "" Then
        SalvarArquivoComo Arquivo
        SalvarArquivo = True
    Else
        SalvarArquivo = False
    End If
End Function


Public Sub SalvarArquivoComo(NomeArquivo)
    On Error Resume Next
    
    Screen.MousePointer = 11
    frmMain.txtRTF.SaveFile NomeArquivo, TipoSalvar
    
    Screen.MousePointer = 0
    
    If Err Then
        MsgBox Error, vbCritical, "ERRO"
    Else
        frmMain.Caption = "M.C.B. Editor - " & NomeArquivo
        FStado.Modificado = False
        frmMain.ArquivoSalvar.Enabled = False
        frmMain.Toolbar1.Buttons(4).Enabled = False
    End If
End Sub

Sub AbrirArquivoProc()
    Dim MSG As String
    Dim RetVal
    Dim MSGResp As String
    Dim Arquivo As String
    Dim Resposta As Integer
    Dim ArquivoAberto As String
    
    If FStado.Modificado = True Then
        Arquivo = Right(frmMain.Caption, Len(frmMain.Caption) - 16)
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
    On Error Resume Next
    frmMain.Dialogo.CancelError = True
    frmMain.Dialogo.Flags = cdlOFNFileMustExist
    frmMain.Dialogo.DefaultExt = "RTF"
    frmMain.Dialogo.Filter = "Arquivos RTF|*.RTF|Arquivos de texto|*.TXT|Todos os arquivos|*.*"
    frmMain.Dialogo.FileName = ""
    frmMain.Dialogo.ShowOpen
    If Err <> cdlCancel Then
        If frmMain.Dialogo.FilterIndex = 1 Then
            TipoSalvar = rtfRTF
        Else
            TipoSalvar = rtfText
        End If
        ArquivoAberto = frmMain.Dialogo.FileName
        AbrirArquivo ArquivoAberto
        AjeitaMenu ArquivoAberto
    End If
End Sub

Public Sub AbrirArquivo(NomeArquivo)
    Dim Indice As Integer
    
    On Error Resume Next
    
    Screen.MousePointer = 11
    
    If UCase(Right(NomeArquivo, 3)) = "RTF" Then
        frmMain.txtRTF.LoadFile NomeArquivo, rtfRTF
    Else
        frmMain.txtRTF.LoadFile NomeArquivo, rtfText
    End If
    
    If Err Then
        MsgBox "Não é possível ler " & NomeArquivo, vbCritical, "ERRO"
        Screen.MousePointer = 0
        Exit Sub
    End If

    frmMain.Caption = "M.C.B. Editor - " & UCase(NomeArquivo)
    FStado.Modificado = False
    frmMain.ArquivoSalvar.Enabled = False
    frmMain.Toolbar1.Buttons(4).Enabled = False
    Screen.MousePointer = 0
End Sub

Public Function RecenteJáExiste(NomeArquivo) As Boolean
    Dim I As Integer

    For I = 0 To 3
        If UCase(frmMain.ArqRecentes(I).Caption) = UCase(NomeArquivo) Then
            RecenteJáExiste = True
            Exit Function
        End If
    Next I
    RecenteJáExiste = False
End Function

Public Sub AjeitaMenu(NomeArquivo As String)
        Dim Sim As Boolean
        Sim = RecenteJáExiste(NomeArquivo)
        If Not Sim Then
            ConfiguraRecentes NomeArquivo
        End If
        ArquivosRecentes
End Sub

Public Sub ConfiguraRecentes(ArquivoAberto As String)
    Dim I As Integer
    Dim Arquivo As String
    Dim Chave As String
    
    For I = 4 To 1 Step -1
        Chave = "Arquivo" & Trim(Str(I))
        Arquivo = LerIni("Recentes", Chave, "")
        If Trim(Arquivo) <> "" And ((I + 1) <= 4) Then
            Chave = "Arquivo" & Trim(Str(I + 1))
            EscreverIni "Recentes", Chave, Arquivo
        End If
    Next I
    
    EscreverIni "Recentes", "Arquivo1", ArquivoAberto
    
End Sub

Public Sub AtribuiParametrosDoForm(F As Form)
    Dim Topo As String
    Dim Esquerda As String
    Dim Altura As String
    Dim Largura As String
    Dim Maximizado As String
    Maximizado = LerIni(F.Name, "Maximizado", "")
    Topo = LerIni(F.Name, "Top", "")
    Esquerda = LerIni(F.Name, "Left", "")
    Largura = LerIni(F.Name, "Width", "")
    Altura = LerIni(F.Name, "Height", "")
    If Trim(Maximizado) = "" Then Exit Sub
    If Maximizado = "1" Then
        F.WindowState = vbMaximized
    Else
        F.Top = CDbl(Topo)
        F.Left = CDbl(Esquerda)
        F.Height = CDbl(Altura)
        F.Width = CDbl(Largura)
    End If
End Sub

Public Sub Mens(Texto As String)
    frmMain.StatusBar1.Panels(1).Text = Texto
End Sub

Public Sub GravaParametrosDoForm(F As Form)
    If F.WindowState = vbMinimized Then Exit Sub
    If F.WindowState = vbMaximized Then
        EscreverIni F.Name, "Maximizado", "1"
    Else
        EscreverIni F.Name, "Maximizado", "0"
        EscreverIni F.Name, "Top", F.Top
        EscreverIni F.Name, "Left", F.Left
        EscreverIni F.Name, "Width", F.Width
        EscreverIni F.Name, "Height", F.Height
    End If
End Sub

Public Sub CheckComandos(ComandoMenu As Menu, NúmeroDoBotão As Integer)
    ComandoMenu.Checked = Not ComandoMenu.Checked
    If ComandoMenu.Checked Then
        frmMain.Toolbar1.Buttons(NúmeroDoBotão).Value = tbrPressed
    Else
        frmMain.Toolbar1.Buttons(NúmeroDoBotão).Value = tbrUnpressed
    End If
End Sub

Public Sub SempreVisível(F As Form, X As Long, Y As Long, cX As Long, cY As Long)
    SetWindowPos F.hwnd, HWND_TOPMOST, X, X, cX, cY, SWP_SHOWWINDOW
End Sub
