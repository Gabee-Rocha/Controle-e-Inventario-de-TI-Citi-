VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Cadastro 
   Caption         =   "Citi - Cadastro e Inventário de TI"
   ClientHeight    =   7410
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   14235
   OleObjectBlob   =   "Cadastro.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CtLab1() As Efeitos
Dim wPlan As Worksheet

Dim nAtualizaForm As New Classe1
Dim T
Dim frmUserWidth As Double
Dim frmUserWidthRatio As Double
Dim frmUserHeight As Double
Dim frmUserHeightRatio As Double
Dim r As Integer
Dim c As Integer
Dim ctl As Control


Private Sub BuscarBt2_Click()
' Habilita e desabilita algumas labels e seta o foco na barra de pesquisa
' ao clicar no botao buscar
Me.SalvarMov.Visible = False
Me.LabelLoja2.Visible = False
Me.Loja2Cb.Visible = False

Me.labelPesq.Visible = True
Me.PesqTx.Enabled = True
Me.PesqTx.Visible = True
Me.PesqTx.SetFocus

End Sub

Private Sub CatCb_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
HookListBoxScroll Me, Me.CatCb
End Sub


Private Sub FecharBt2_Click()

Call Fechar
Application.Visible = True

End Sub

Private Sub Imprimir1Bt2_Click()

'imprimir Termo Entrega

    Application.Visible = True
    Unload Me
    Sheets("Termo_Entrega").Visible = True
    Sheets("Termo_Entrega").Select
    Range("A1:I49").Select
    
    Selection.PrintPreview
    Sheets("Termo_Entrega").Visible = False
    Cadastro.Show
    
    
End Sub


Private Sub Imprimir2Bt2_Click()

'imprimir Termo Responsabilidade
    Application.Visible = True
    Unload Me
    Sheets("Termo_Responsabilidade").Visible = True
    Sheets("Termo_Responsabilidade").Select
    Range("A1:H47").Select
    
    Selection.PrintPreview
    Sheets("Termo_Responsabilidade").Visible = False
    Cadastro.Show
    
End Sub

Private Sub Imprimir3Bt2_Click()

'imprimir Termo Devolução
    Application.Visible = True
    Unload Me
    Sheets("Termo_Devolução").Visible = True
    Sheets("Termo_Devolução").Select
    Range("A1:I49").Select
    
    Selection.PrintPreview
    Sheets("Termo_Devolução").Visible = False
    Cadastro.Show

End Sub

Private Sub LabelExcluirRegBt2_Click()
' Efetua a exclusao do dado na planilda de acordo com a seleção
' do item no listbox
Dim indc
indc = Me.ListBox1.ListIndex + 2
If MsgBox("Deseja excluir o item " & Me.CodIgo.Text & " ?", vbExclamation + vbYesNo, "Popular listbox e deletar item") = vbYes Then
    Planilha1.Rows(indc).Delete
MsgBox ("Item Deletado com sucesso!"), vbInformation, "Popular listbox e deletar item"
End If

Call movi_excluir

ThisWorkbook.Save
Unload Cadastro
Cadastro.Show

End Sub

Private Sub LabelSalvarEdicaoBt2_Click()

Call editar

End Sub

Private Sub LabelSNBt2_Click()
' Verifica se os campos foram preenchidos
If Me.CatCb.Text = "" Or Me.MMTx.Text = "" Or _
    Me.LojaCb.Text = "" Or Me.SetorTx.Text = "" Or Me.UserTx.Text = "" Then
        MsgBox "Preencha os campos obrigatórios", vbExclamation, "Campo vazio!"

Exit Sub
End If

Call GravarNovo

End Sub

Private Sub LabelVoltar1Bt2_Click()

Call carregar_ListBox
Call LimpDados

Me.LabelSNBt.Visible = False
Me.LabelVoltar1Bt.Visible = False
Me.LabelSalvarEdicaoBt.Visible = False
Me.LabelExcluirRegBt.Visible = False
Me.labelPesq.Visible = False
Me.PesqTx.Visible = False
Me.LabelConta.Visible = False

Unload Cadastro
Cadastro.Show

End Sub


' Ao clicar em um item passa os para as labels os valores dessa seleção
Private Sub ListBox1_Click()

On Error Resume Next

With Me.ListBox1
Me.CodIgo.Value = .List(.ListIndex, 0)
Me.CatCb.Value = .List(.ListIndex, 1)
Me.MMTx.Value = .List(.ListIndex, 2)
Me.PatrTx.Value = .List(.ListIndex, 3)
Me.SerieTx.Value = .List(.ListIndex, 4)
Me.LojaCb.Value = .List(.ListIndex, 5)
Me.SetorTx.Value = .List(.ListIndex, 6)
Me.UserTx.Value = .List(.ListIndex, 8)

Call Grava_Entrega
Call Grava_Respons
Call Grava_Devolu

End With

End Sub


Private Sub Loja2Cb_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
HookListBoxScroll Me, Me.Loja2Cb
End Sub

Private Sub LojaCb_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
HookListBoxScroll Me, Me.LojaCb
End Sub

Private Sub MovimentarBt2_Click()
Me.LabelLoja2.Visible = True
Me.Loja2Cb.Visible = True
Me.SalvarMov.Visible = True
Me.labelPesq.Visible = False
Me.PesqTx.Visible = False

End Sub

' Seta opatrimonio para receber apenas numeros e até 6 caracteres
Private Sub PatrTx_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

Me.PatrTx.MaxLength = 6

If KeyAscii < 48 Or KeyAscii > 57 Then
    KeyAscii = 0
End If

End Sub

' Desativa Label de pesquisa por Loja ao passar o curso por cima
Private Sub PesqTx_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Me.labelPesq.Visible = False

Call Pesquisar
End Sub

' Seta O CAMPO quantidade para receber apenas numeros e até 3 caracteres
Private Sub QtdeTx_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

Me.QtdeTx.MaxLength = 3

If KeyAscii < 48 Or KeyAscii > 57 Then
    KeyAscii = 0
End If

End Sub

Private Sub RelBt2_Click()
Unload Cadastro
Relatorios.Show
End Sub

Private Sub SalvarMov_Click()
Call Mov_Loja
End Sub

Private Sub UserForm_Activate()
Set nAtualizaForm.Form = Me
End Sub

Private Sub UserForm_Initialize()

Dim ObjetoBt As Object
Dim conta As Long

Application.Visible = False
Cadastro.StartUpPosition = 2

Dim hWnd As Long

    'Vai para o topo do formulário
    ScrollTop = 0

    'Define os botões minimizar e maximizar do form
    hWnd = FindWindow(vbNullString, Me.Caption)
    SetWindowLong hWnd, -16, &H20000 Or &H10000 Or &H84C80080
    
    frmUserWidth = Me.InsideWidth
    frmUserHeight = Me.InsideHeight

ReDim CtLab1(1 To Me.Controls.Count)

For Each ObjetoBt In Me.Controls

    If TypeName(ObjetoBt) = "Label" Then
        conta = conta + 1
        Set CtLab1(conta) = New Efeitos
        Set CtLab1(conta).cntLb = ObjetoBt
        
    End If
Next ObjetoBt
Set ObjetoBt = Nothing
ReDim Preserve CtLab1(1 To conta) ' Mantem os valores que foram inseridos dentro do array

Call Design

' Carrega os valores de categorias
With Me.CatCb
    .AddItem "Desktop"
    .AddItem "Notebook"
    .AddItem "Impressora"
    .AddItem "Zebra"
    .AddItem "Monitor"
    .AddItem "Tablet"
    .AddItem "Celular"
    .AddItem "Adaptadores"
    .AddItem "Busca Preço"
    .AddItem "Leitor de Mão"
    .AddItem "Relogio Ponto"
    .AddItem "TV"
    .AddItem "Nobreak"
    .AddItem "VOIP"
    .AddItem "Switch"
    .AddItem "Firewall"
    .AddItem "Pinpad"
    .AddItem "Pos"
    .AddItem "Gaveta"
    .AddItem "Servidor"
    .AddItem "Periferico"
    .AddItem "Fonte ATX"
    .AddItem "Access Point"
    .AddItem "Outros"
    
End With

' Carrega os valores das Lojas
With Me.LojaCb
    .AddItem "ESTOQUE TI - PR"
    .AddItem "ESTOQUE TI - ES"
    .AddItem "ESTOQUE TI - BA"
    .AddItem "01 - LOJA"
    .AddItem "01 - RDS - LOJA"
    .AddItem "02 - LOJA"
    .AddItem "02 - RDS - LOJA"
    .AddItem "03 - CD"
    .AddItem "04 - LOJA"
    .AddItem "04 - RDS - LOJA"
    .AddItem "06 - LOJA"
    .AddItem "07 - LOJA"
    .AddItem "07 - RDS - LOJA"
    .AddItem "08 - LOJA"
    .AddItem "09 - LOJA"
    .AddItem "10 - LOJA"
    .AddItem "10 - RDS - LOJA"
    .AddItem "11 - LOJA"
    .AddItem "12 - RDS - LOJA"
    .AddItem "14 - LOJA"
    .AddItem "201 - RDS - LOJA"
    .AddItem "202 - RDS - LOJA"
    .AddItem "600 - D&G - LOJA"
    .AddItem "601 - D&G - LOJA"
    .AddItem "701 - LOJA"
    .AddItem "703 - LOJA"
    .AddItem "704 - LOJA"
    .AddItem "705 - LOJA"
    .AddItem "706 - CD ESPIRITO-SANTO"
    .AddItem "801 - LOJA"
    .AddItem "802 - LOJA"
    .AddItem "803 - LOJA"
    .AddItem "804 - CD BAHIA"
End With

' Carrega os valores das Lojas no campo do movimento
With Me.Loja2Cb
    .AddItem "ESTOQUE TI - PR"
    .AddItem "ESTOQUE TI - ES"
    .AddItem "ESTOQUE TI - BA"
    .AddItem "01 - LOJA"
    .AddItem "01 - RDS - LOJA"
    .AddItem "02 - LOJA"
    .AddItem "02 - RDS - LOJA"
    .AddItem "03 - CD"
    .AddItem "04 - LOJA"
    .AddItem "04 - RDS - LOJA"
    .AddItem "06 - LOJA"
    .AddItem "07 - LOJA"
    .AddItem "07 - RDS - LOJA"
    .AddItem "08 - LOJA"
    .AddItem "09 - LOJA"
    .AddItem "10 - LOJA"
    .AddItem "10 - RDS - LOJA"
    .AddItem "11 - LOJA"
    .AddItem "12 - RDS - LOJA"
    .AddItem "14 - LOJA"
    .AddItem "201 - RDS - LOJA"
    .AddItem "202 - RDS - LOJA"
    .AddItem "600 - D&G - LOJA"
    .AddItem "601 - D&G - LOJA"
    .AddItem "701 - LOJA"
    .AddItem "703 - LOJA"
    .AddItem "704 - LOJA"
    .AddItem "705 - LOJA"
    .AddItem "706 - CD ESPIRITO-SANTO"
    .AddItem "801 - LOJA"
    .AddItem "802 - LOJA"
    .AddItem "803 - LOJA"
    .AddItem "804 - CD BAHIA"
End With

Me.BuscarBt.Visible = True
Me.LabelSNBt.Visible = True
Me.LabelSalvarEdicaoBt.Visible = True
Me.LabelExcluirRegBt.Visible = True
Me.FecharBt.Visible = True
Me.LabelVoltar1Bt.Visible = True

Me.LabelLoja2.Visible = False
Me.Loja2Cb.Visible = False

Me.PesqTx.Visible = False
Me.labelPesq.Visible = False

Me.LabelConta.ForeColor = RGB(0, 191, 255)

Call carregar_ListBox
Call CriaCabecalhoLb(Me.ListBox1, Me.lbCabecalho, Array("Cód", "Categoria", "Marca e Modelo" _
        , "Patrimônio", "Serie", "Loja", "Local", "Data", "Usuário"))
End Sub

' Limpa os dados do controle
Sub LimpDados()

Dim nvCont As Object

For Each nvCont In Me.Controls
    If TypeName(nvCont) = "TextBox" Or _
        TypeName(nvCont) = "ComboBox" Or _
        TypeName(nvCont) = "OptionButton" Then
        nvCont.Value = ""
    End If
    
Next nvCont
Set nvCont = Nothing

End Sub

' Grava na planilha os dados
Sub GravarNovo()

Dim mov_grav As Worksheet
Dim wPlan As Worksheet
Dim nLin As Long
Dim vCod As Long
Dim mov_str As String
Dim lin_mov As Long

Dim Linha As Long
Dim lin As Long


Planilha1.Activate
Linha = Planilha1.Range("A2").End(xlDown).Row + 1
lin = 2

If Me.PatrTx.Value = Empty Then
GoTo Inicio
End If

While lin < Linha
    If Cells(lin, 4) = Val(Me.PatrTx.Value) Then
        MsgBox "Este Patrimônio já esta cadastrado!", vbCritical
        Exit Sub
    End If
    lin = lin + 1
Wend

Inicio:

mov_str = "Novo Inserido"
Set wPlan = Planilha1
Set mov_grav = Movimento

nLin = Application.WorksheetFunction.CountA(wPlan.Range("A:A"))
lin_mov = Application.WorksheetFunction.CountA(mov_grav.Range("A:A"))

If nLin = 1 Then
    vCod = 1
Else
    vCod = Application.WorksheetFunction.Large(wPlan.Range("A:A"), 1)
    vCod = vCod + 1
End If
nLin = nLin + 1

With wPlan
    .Cells(nLin, 1) = vCod
    .Cells(nLin, 2) = CatCb
    .Cells(nLin, 3) = MMTx
    .Cells(nLin, 4) = PatrTx
    .Cells(nLin, 5) = SerieTx
    .Cells(nLin, 6) = LojaCb
    .Cells(nLin, 7) = SetorTx
    .Cells(nLin, 8) = Date
    .Cells(nLin, 9) = UserTx
End With

lin_mov = lin_mov + 1

With mov_grav
    .Cells(lin_mov, 1) = vCod
    .Cells(lin_mov, 2) = mov_str
    .Cells(lin_mov, 3) = CatCb
    .Cells(lin_mov, 4) = MMTx
    .Cells(lin_mov, 5) = LojaCb
    .Cells(lin_mov, 6) = PatrTx
    .Cells(lin_mov, 7) = Date
    .Cells(lin_mov, 8) = UserTx
    .Cells(lin_mov, 10) = .Application.UserName
    
End With
MsgBox "Equipamento Cadastrado com sucesso!", vbInformation

ThisWorkbook.Save
Unload Cadastro
Cadastro.Show


End Sub

' Carregar List Box
Sub carregar_ListBox()

Me.ListBox1.Enabled = True
Dim LinhaListBox As Long
Dim lin As Long
Dim ultima As Long

LinhaListBox = 0
lin = 2

    Me.ListBox1.ColumnCount = 9
    Me.ListBox1.ColumnHeads = False
    Me.ListBox1.ColumnWidths = "60;100;260;120;170;150;150;110;260"
    Me.ListBox1.ForeColor = RGB(0, 191, 255)

    While Planilha1.Range("A" & lin).Value <> Empty
        Me.ListBox1.AddItem
        Me.ListBox1.List(LinhaListBox, 0) = Planilha1.Range("A" & lin).Value
        Me.ListBox1.List(LinhaListBox, 1) = Planilha1.Range("B" & lin).Value
        Me.ListBox1.List(LinhaListBox, 2) = Planilha1.Range("C" & lin).Value
        Me.ListBox1.List(LinhaListBox, 3) = Planilha1.Range("D" & lin).Value
        Me.ListBox1.List(LinhaListBox, 4) = Planilha1.Range("E" & lin).Value
        Me.ListBox1.List(LinhaListBox, 5) = Planilha1.Range("F" & lin).Value
        Me.ListBox1.List(LinhaListBox, 6) = Planilha1.Range("G" & lin).Value
        Me.ListBox1.List(LinhaListBox, 7) = Planilha1.Range("H" & lin).Value
        Me.ListBox1.List(LinhaListBox, 8) = Planilha1.Range("I" & lin).Value
    
        LinhaListBox = LinhaListBox + 1
        lin = lin + 1
    Wend
    Me.LabelConta.Caption = Me.ListBox1.ListCount & " Registro(s) Encontrado(s)"
    
End Sub

' Sub para editar o conteudo selecionado

Sub editar()

Dim resposta As VbMsgBoxResult
Dim valor As Long
Dim fila As Object
Dim Linha As Long

Dim mov_edi As Worksheet
Dim mov_stredi As String
Dim lin_movedi As Long

mov_stredi = "Editado"
Set mov_edi = Movimento

lin_movedi = Application.WorksheetFunction.CountA(mov_edi.Range("A:A"))

    If Me.CodIgo.Value = "" Then
        MsgBox "Selecione um Cadastro para Alterar!"
    Exit Sub
    End If
    valor = Me.CodIgo.Value
    resposta = MsgBox("Deseja realmente Alterar o item - " & valor & "?", vbYesNo)

    If resposta = vbNo Then
        MsgBox "O item não foi alterado!", vbExclamation
        Exit Sub

    Else
        Set fila = Planilha1.Range("A:A").Find(valor, lookat:=xlWhole)
        Linha = fila.Row
        Planilha1.Range("B" & Linha).Value = Me.CatCb.Value
        Planilha1.Range("C" & Linha).Value = Me.MMTx.Value
        Planilha1.Range("D" & Linha).Value = Me.PatrTx.Value
        Planilha1.Range("E" & Linha).Value = Me.SerieTx.Value
        Planilha1.Range("F" & Linha).Value = Me.LojaCb.Value
        Planilha1.Range("G" & Linha).Value = Me.SetorTx.Value
        Planilha1.Range("H" & Linha).Value = Date
        Planilha1.Range("I" & Linha).Value = Me.UserTx.Value
        
        lin_movedi = lin_movedi + 1
        
        With Movimento
            .Cells(lin_movedi, 1) = Me.CodIgo.Value
            .Cells(lin_movedi, 2) = mov_stredi
            .Cells(lin_movedi, 3) = CatCb
            .Cells(lin_movedi, 4) = MMTx
            .Cells(lin_movedi, 5) = LojaCb
            .Cells(lin_movedi, 6) = PatrTx
            .Cells(lin_movedi, 7) = Date
            .Cells(lin_movedi, 8) = UserTx
            .Cells(lin_movedi, 10) = Application.UserName
        End With

        Call LimpDados
        
        ListBox1.ListIndex = -1

        MsgBox "Cadastro Alterado com Sucesso!"
    End If
ThisWorkbook.Save
Unload Cadastro
Cadastro.Show
End Sub

' Seta as cores e referencias para textbox, combobox e listbox
Sub Design()

    Dim ef As Object
    Dim vcor1 As Variant
    Dim vcor2 As Variant

    vcor1 = RGB(20, 0, 26)
    vcor2 = RGB(0, 191, 255)
    Me.BackColor = vcor1
    For Each ef In Me.Controls
        If TypeName(ef) = "TextBox" Then
            ef.BorderColor = vcor2
        ElseIf TypeName(ef) = "ComboBox" Then
            ef.BorderColor = vcor2
        ElseIf TypeName(ef) = "ListBox" Then
            ForeColor = vcor2
            ef.BorderColor = vcor2
            ef.BackColor = vcor1
        End If

    Next ef
    Set ef = Nothing
End Sub

' Sub para fechar e salvar o arquivo
Private Sub Fechar()

    Dim nome As String
    nome = ThisWorkbook.Name
    ThisWorkbook.Save
    Windows(nome).Close
    
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
TiraEfeitos
End Sub

' Grava os valores no termo de devolução
Sub Grava_Devolu()
Dim wPlan As Worksheet

Set wPlan = Planilha5

With wPlan
    .Cells(21, 5) = CatCb.Text
    .Cells(22, 3) = MMTx.Text
    .Cells(23, 4) = PatrTx.Text
    .Cells(24, 3) = SerieTx.Text
    .Cells(15, 5) = Date
    .Cells(14, 3) = UserTx.Text
    .Cells(29, 2) = UserTx.Text
    .Cells(48, 3) = Date
End With

End Sub

' Grava os valores no termo de Responsabilidade
Sub Grava_Respons()
Dim wPlan As Worksheet

Set wPlan = Planilha4

With wPlan
    .Cells(13, 2) = CatCb.Text
    .Cells(13, 3) = MMTx.Text
    .Cells(13, 8) = PatrTx.Text
    .Cells(13, 5) = SerieTx.Text
    .Cells(14, 3) = UserTx.Text
    .Cells(43, 5) = UserTx.Text
    .Cells(46, 3) = Date
End With

End Sub

' Grava os valores no termo de Entrega
Sub Grava_Entrega()
Dim wPlan As Worksheet

Set wPlan = Planilha3

With wPlan
    .Cells(22, 4) = CatCb.Text
    .Cells(23, 3) = MMTx.Text
    .Cells(24, 3) = PatrTx.Text
    .Cells(25, 3) = SerieTx.Text
    .Cells(12, 5) = Date
    .Cells(26, 4) = UserTx.Text
    .Cells(37, 2) = UserTx.Text
    .Cells(45, 3) = Date
End With

End Sub

' Sub para pesquisar por Loja
Sub Pesquisar()

Dim Xcel As String
Dim Coluna As Integer
Dim LinhaListBox As Integer
Dim Linha As Integer

Linha = 2
LinhaListBox = 0

With Planilha1
With Me.ListBox1
.Clear
    While Planilha1.Cells(Linha, 1) <> Empty
    For Coluna = 1 To 9
    Xcel = Planilha1.Cells(Linha, Coluna)
    If InStr(1, UCase(Xcel), UCase(Me.PesqTx.Text)) > 0 Then
        .AddItem
        .List(LinhaListBox, 0) = Planilha1.Cells(Linha, 1)
        .List(LinhaListBox, 1) = Planilha1.Cells(Linha, 2)
        .List(LinhaListBox, 2) = Planilha1.Cells(Linha, 3)
        .List(LinhaListBox, 3) = Planilha1.Cells(Linha, 4)
        .List(LinhaListBox, 4) = Planilha1.Cells(Linha, 5)
        .List(LinhaListBox, 5) = Planilha1.Cells(Linha, 6)
        .List(LinhaListBox, 6) = Planilha1.Cells(Linha, 7)
        .List(LinhaListBox, 7) = Planilha1.Cells(Linha, 8)
        .List(LinhaListBox, 8) = Planilha1.Cells(Linha, 9)
        LinhaListBox = LinhaListBox + 1
        
        GoTo Proxima_Linha
        End If
    Next Coluna
Proxima_Linha:
    Linha = Linha + 1
    Wend
    Me.LabelConta.Caption = .ListCount & " Registro(s) Encontrado(s)"
End With
End With
End Sub


Public Sub CriaCabecalhoLb(lbPrincipal As MSForms.ListBox, lbCabecalho As MSForms.ListBox, cabecalho As Variant)
    
    With lbCabecalho
        'Iguala o numeros de colunas do ListBox Cabeçalho ao do ListBox Principal
        .ColumnCount = lbPrincipal.ColumnCount
        .ColumnWidths = lbPrincipal.ColumnWidths
        
        'Adiciona os elementos dos cabeçalhos
        .Clear
        .AddItem
        Dim i As Integer
        For i = 0 To UBound(cabecalho)
            .List(0, i) = cabecalho(i)
        Next i
        
        'Formata o visual
        .ZOrder (0)
        .Font.Size = 12
        .Font.Bold = True
        .SpecialEffect = fmSpecialEffectFlat
        .BackColor = RGB(35, 207, 222) 'RGB(229, 13, 90)
        .Height = 12
        
        'Alinha a posição e dimensões do ListBox Cabeçalho ao ListBox Principal
        .Width = lbPrincipal.Width
        .Left = lbPrincipal.Left
        .Top = lbPrincipal.Top - (.Height - 1)
    End With
    lbPrincipal.ZOrder (1)

End Sub
' sub para trocar de loja e gerar o relatorio
Sub Mov_Loja()

Dim resposta As VbMsgBoxResult
Dim valor As Long
Dim fila As Object
Dim Linha As Long

Dim mov_strmov As String
Dim lin_movmov As Long
Dim mov_gravmov As Worksheet

mov_strmov = "Movimentado"
Set mov_gravmov = Movimento

lin_movmov = Application.WorksheetFunction.CountA(mov_gravmov.Range("A:A"))
 
        If Me.CodIgo.Value = "" Then
            MsgBox "Selecione um Cadastro para Movimentar!"
        Exit Sub
        End If
        valor = Me.CodIgo.Value
        resposta = MsgBox("Deseja realmente Movimentar o item: " & valor & "?", vbYesNo)

    
        If resposta = vbNo Then
            Exit Sub
        Else
            Set fila = Planilha1.Range("A:A").Find(valor, lookat:=xlWhole)
            Linha = fila.Row
            Planilha1.Range("B" & Linha).Value = Me.CatCb.Value
            Planilha1.Range("C" & Linha).Value = Me.MMTx.Value
            Planilha1.Range("D" & Linha).Value = Me.PatrTx.Value
            Planilha1.Range("E" & Linha).Value = Me.SerieTx.Value
            Planilha1.Range("F" & Linha).Value = Me.Loja2Cb.Value
            Planilha1.Range("G" & Linha).Value = Me.SetorTx.Value
            Planilha1.Range("H" & Linha).Value = Date
            Planilha1.Range("I" & Linha).Value = Me.UserTx.Value
            
            lin_movmov = lin_movmov + 1
            
            With Movimento
                .Cells(lin_movmov, 1) = Me.CodIgo
                .Cells(lin_movmov, 2) = mov_strmov
                .Cells(lin_movmov, 3) = CatCb
                .Cells(lin_movmov, 4) = MMTx
                .Cells(lin_movmov, 5) = LojaCb
                .Cells(lin_movmov, 6) = PatrTx
                .Cells(lin_movmov, 7) = Date
                .Cells(lin_movmov, 8) = UserTx
                .Cells(lin_movmov, 9) = Loja2Cb
                .Cells(lin_movmov, 10) = Application.UserName
            End With
            
            Call LimpDados
            
            ListBox1.ListIndex = -1
            
            MsgBox "Equipamento Movido com Sucesso!"
        End If
        
ThisWorkbook.Save
Unload Cadastro
Cadastro.Show

End Sub

Sub movi_excluir()

Dim mov_strex As String
Dim lin_movex As Long
Dim mov_gravex As Worksheet

mov_strex = "Excluido"

Set mov_gravex = Movimento

lin_movex = Application.WorksheetFunction.CountA(mov_gravex.Range("A:A"))
lin_movex = lin_movex + 1

With Movimento
    .Cells(lin_movex, 1) = Me.CodIgo
    .Cells(lin_movex, 2) = mov_strex
    .Cells(lin_movex, 3) = CatCb
    .Cells(lin_movex, 4) = MMTx
    .Cells(lin_movex, 5) = LojaCb
    .Cells(lin_movex, 6) = PatrTx
    .Cells(lin_movex, 7) = Date
    .Cells(lin_movex, 8) = UserTx
    .Cells(lin_movex, 10) = Application.UserName
End With

End Sub

Private Sub ListBox1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
HookListBoxScroll Me, Me.ListBox1
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
UnhookListBoxScroll
End Sub

Private Sub UserForm_Resize()

If Me.InsideHeight < 1 Then Exit Sub
    
    frmUserWidthRatio = Me.InsideWidth / frmUserWidth
    frmUserHeightRatio = Me.InsideHeight / frmUserHeight
    
  ' Eliminate this section to prevent resizing of controls on form.
    ' Stick any control on the form at any location.
    For Each ctl In Me.Controls
        ctl.Width = frmUserWidthRatio * ctl.Width
        ctl.Left = frmUserWidthRatio * ctl.Left
        ctl.Height = frmUserHeightRatio * ctl.Height
        ctl.Top = frmUserHeightRatio * ctl.Top
    Next
    
    frmUserWidth = Me.InsideWidth
    frmUserHeight = Me.InsideHeight

End Sub

