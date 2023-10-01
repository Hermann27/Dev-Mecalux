Public Class FrmChoixTraitement

    Private Sub RbtFC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RbtFC.Click
        With Frm_FluxEntrantCritére
            .Critere = "CSO"
            .MdiParent = MenuPrincipal
            .Text = " Traitement Transformation Confirmation de commande "
            .Show()
        End With
    End Sub

    Private Sub RbtTdepot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RbtTdepot.Click
        With Frm_ConfirmationReception
            .Critere = "CRP"
            .MdiParent = MenuPrincipal
            .Text = " Traitement Transformation Confrimation de Reception "
            .Show()
        End With
    End Sub
End Class