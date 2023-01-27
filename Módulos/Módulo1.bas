Attribute VB_Name = "Módulo1"
Option Explicit

Function TiraEfeitos()

With Cadastro
    .BuscarBt.Visible = True
    .FecharBt.Visible = True
    .Imprimir1Bt.Visible = True
    .Imprimir2Bt.Visible = True
    .Imprimir3Bt.Visible = True
    .LabelSNBt.Visible = True
    .LabelSalvarEdicaoBt.Visible = True
    .LabelExcluirRegBt.Visible = True
    .LabelVoltar1Bt.Visible = True
    .MovimentarBt.Visible = True
    .RelBt.Visible = True
    
    .BuscarBt2.Visible = False
    .FecharBt2.Visible = False
    .Imprimir1Bt2.Visible = False
    .Imprimir2Bt2.Visible = False
    .Imprimir3Bt2.Visible = False
    .LabelSNBt2.Visible = False
    .LabelSalvarEdicaoBt2.Visible = False
    .LabelExcluirRegBt2.Visible = False
    .LabelVoltar1Bt2.Visible = False
    .MovimentarBt2.Visible = False
    .RelBt2.Visible = False
    
End With
End Function

Function TiraEfeitoRel()

With Relatorios
    
    .BuscarBt.Visible = True
    .VoltarpagBt.Visible = True
    
    .BuscarBt2.Visible = False
    .VoltarpagBt2.Visible = False
End With
End Function


