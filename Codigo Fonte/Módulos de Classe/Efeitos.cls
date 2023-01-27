VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Efeitos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents cntLb As MSForms.Label
Attribute cntLb.VB_VarHelpID = -1
Public WithEvents cntTx As MSForms.TextBox
Attribute cntTx.VB_VarHelpID = -1
Public WithEvents cntCb As MSForms.ComboBox
Attribute cntCb.VB_VarHelpID = -1


Private Sub cntLb_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

If Right(cntLb.Name, 2) = "Bt" Then
    With Cadastro
        .Controls(cntLb.Name).Visible = False
        .Controls(cntLb.Name & "2").Visible = True
        .Controls(cntLb.Name & "2").ZOrder (0)
        .Controls(cntLb.Name & "2").ForeColor = RGB(255, 255, 255)
        .Controls(cntLb.Name & "2").Left = .Controls(cntLb.Name).Left
        .Controls(cntLb.Name & "2").Top = .Controls(cntLb.Name).Top
    End With
End If

End Sub

Private Sub cntLb_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim nCarc As Integer
nCarc = Len(cntLb.Name)

If Right(cntLb.Name, 3) = "Bt2" Then
    With Cadastro.Controls(Left(cntLb.Name, nCarc - 1))
        .Visible = True
        .ZOrder (0)
    End With
End If

End Sub

