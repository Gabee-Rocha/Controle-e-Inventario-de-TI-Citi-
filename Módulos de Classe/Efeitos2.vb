Option Explicit
Public WithEvents cntLb1 As MSForms.Label
Public WithEvents cntTx As MSForms.TextBox
Public WithEvents cntCb As MSForms.ComboBox

Private Sub cntLb1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

If Right(cntLb1.Name, 2) = "Bt" Then
        Relatorios.Controls(cntLb1.Name).Visible = False
        Relatorios.Controls(cntLb1.Name & "2").Visible = True
        Relatorios.Controls(cntLb1.Name & "2").ZOrder (0)
        Relatorios.Controls(cntLb1.Name & "2").ForeColor = RGB(255, 255, 255)
        Relatorios.Controls(cntLb1.Name & "2").Left = Relatorios.Controls(cntLb1.Name).Left
        Relatorios.Controls(cntLb1.Name & "2").Top = Relatorios.Controls(cntLb1.Name).Top
End If

End Sub

Private Sub cntLb1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Dim nCarc As Integer

nCarc = Len(cntLb1.Name)

If Right(cntLb1.Name, 3) = "Bt2" Then
        With Relatorios.Controls(Left(cntLb1.Name, nCarc - 1))
            .Visible = True
            .ZOrder (0)
        End With
End If

End Sub
