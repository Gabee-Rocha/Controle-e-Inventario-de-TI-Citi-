VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Relatorios 
   Caption         =   "CTI - Relatorios"
   ClientHeight    =   8220.001
   ClientLeft      =   105
   ClientTop       =   465
   ClientWidth     =   17955
   OleObjectBlob   =   "Relatorios.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Relatorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ctLabl() As Efeitos2
Dim nAtualizaForm As New Classe1
Dim T
Dim frmUserWidth As Double
Dim frmUserWidthRatio As Double
Dim frmUserHeight As Double
Dim frmUserHeightRatio As Double
Dim r As Integer
Dim c As Integer
Dim ctl As Control

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
        .Font.Size = 9
        .Font.Bold = True
        .SpecialEffect = fmSpecialEffectFlat
        .BackColor = RGB(35, 207, 222) 'RGB(229, 13, 90)
        .Height = 13
        
        'Alinha a posição e dimensões do ListBox Cabeçalho ao ListBox Principal
        .Width = lbPrincipal.Width
        .Left = lbPrincipal.Left
        .Top = lbPrincipal.Top - (.Height - 1)
    
    End With
    lbPrincipal.ZOrder (1)
    
End Sub

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


Private Sub BuscarBt2_Click()
Me.PesquisaTx.Visible = True
Me.labelPesq.Visible = True
Me.PesquisaTx.Enabled = True
Me.PesquisaTx.SetFocus
End Sub


Private Sub RelEntradaBt2_Click()
Call Carregar_List_Entrada
End Sub


Private Sub PesquisaTx_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Me.labelPesq.Visible = False

Call Pesquisar
End Sub

Private Sub UserForm_Activate()
Set nAtualizaForm.Form = Me

Call CriaCabecalhoLb(Me.ListBox1, Me.lbCabecalho, Array("Cód", "Movimento", "Categoria" _
        , "Marca e Modelo", "Loja", "Patrimônio", "Data", "Usuário", "Movido Para", "Feito Por"))
End Sub


Private Sub UserForm_Initialize()
Dim ObjetoBt As Object
Dim conta As Long
Dim contador As Long

Relatorios.StartUpPosition = 2
contador = Me.ListBox1.ListCount


Dim hWnd As Long

    'Vai para o topo do formulário
    ScrollTop = 0

    'Define os botões minimizar e maximizar do form
    hWnd = FindWindow(vbNullString, Me.Caption)
    SetWindowLong hWnd, -16, &H20000 Or &H10000 Or &H84C80080
    
    frmUserWidth = Me.InsideWidth
    frmUserHeight = Me.InsideHeight
    
ReDim ctLabl(1 To Me.Controls.Count)

For Each ObjetoBt In Me.Controls

    If TypeName(ObjetoBt) = "Label" Then
        conta = conta + 1
        Set ctLabl(conta) = New Efeitos2
        Set ctLabl(conta).cntLb1 = ObjetoBt
    End If
Next ObjetoBt
Set ObjetoBt = Nothing
ReDim Preserve ctLabl(1 To conta) ' Mantem os valores que foram inseridos dentro do array

Call Design

Call Carregar_List_Entrada
Me.PesquisaTx.Visible = False
Me.labelPesq.Visible = False
Me.LabelContador.Visible = True
Me.LabelContador.ForeColor = RGB(0, 191, 255)

End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
TiraEfeitoRel
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

Sub Carregar_List_Entrada()

' Carrega o List Box com os valores do relátorio de Movimentação
Me.ListBox1.Enabled = True
Dim LinhaListBox As Long
Dim lin As Long
Dim ultima As Long

LinhaListBox = 0
lin = 2

    Me.ListBox1.ColumnCount = 10
    Me.ListBox1.ColumnHeads = False
    Me.ListBox1.ColumnWidths = "30;90;100;150;100;90;90;130;100;150"
    Me.ListBox1.ForeColor = RGB(0, 191, 255)

    While Movimento.Range("A" & lin).Value <> Empty
        Me.ListBox1.AddItem
        Me.ListBox1.List(LinhaListBox, 0) = Movimento.Range("A" & lin).Value
        Me.ListBox1.List(LinhaListBox, 1) = Movimento.Range("B" & lin).Value
        Me.ListBox1.List(LinhaListBox, 2) = Movimento.Range("C" & lin).Value
        Me.ListBox1.List(LinhaListBox, 3) = Movimento.Range("D" & lin).Value
        Me.ListBox1.List(LinhaListBox, 4) = Movimento.Range("E" & lin).Value
        Me.ListBox1.List(LinhaListBox, 5) = Movimento.Range("F" & lin).Value
        Me.ListBox1.List(LinhaListBox, 6) = Movimento.Range("G" & lin).Value
        Me.ListBox1.List(LinhaListBox, 7) = Movimento.Range("H" & lin).Value
        Me.ListBox1.List(LinhaListBox, 8) = Movimento.Range("I" & lin).Value
        Me.ListBox1.List(LinhaListBox, 9) = Movimento.Range("J" & lin).Value
        
        LinhaListBox = LinhaListBox + 1
        lin = lin + 1
    Wend
End Sub


Sub Pesquisar()

Dim Xcel As String
Dim Coluna As Integer
Dim LinhaListBox As Integer
Dim Linha As Integer

Linha = 2
LinhaListBox = 0

With Movimento
With Me.ListBox1
.Clear
    While Movimento.Cells(Linha, 1) <> Empty
    For Coluna = 1 To 10
    Xcel = Movimento.Cells(Linha, Coluna)
    If InStr(1, UCase(Xcel), UCase(Me.PesquisaTx.Text)) > 0 Then
        .AddItem
        .List(LinhaListBox, 0) = Movimento.Cells(Linha, 1)
        .List(LinhaListBox, 1) = Movimento.Cells(Linha, 2)
        .List(LinhaListBox, 2) = Movimento.Cells(Linha, 3)
        .List(LinhaListBox, 3) = Movimento.Cells(Linha, 4)
        .List(LinhaListBox, 4) = Movimento.Cells(Linha, 5)
        .List(LinhaListBox, 5) = Movimento.Cells(Linha, 6)
        .List(LinhaListBox, 6) = Movimento.Cells(Linha, 7)
        .List(LinhaListBox, 7) = Movimento.Cells(Linha, 8)
        .List(LinhaListBox, 8) = Movimento.Cells(Linha, 9)
        .List(LinhaListBox, 9) = Movimento.Cells(Linha, 10)
        LinhaListBox = LinhaListBox + 1
        
        GoTo Proxima_Linha
        End If
    Next Coluna
Proxima_Linha:
    Linha = Linha + 1
    Wend
    Me.LabelContador.Caption = .ListCount & " Registro(s) Encontrado(s)"
    
End With
End With
End Sub


Private Sub VoltarpagBt2_Click()
Unload Relatorios
Cadastro.Show
End Sub

' A Proximas duas subs são para simular o Scrool do Mouse
Private Sub ListBox1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
HookListBoxScroll Me, Me.ListBox1
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
UnhookListBoxScroll
End Sub

