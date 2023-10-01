Public Class FrmChoix

    Private Sub RbtArt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RbtArt.Click
        With FrmAllCorrespondance
            .MdiParent = MenuApplication1
            .Text &= "]<-->[Correspondance Article/Detail Article"
            .Critère = "PRO"
            .Show()
        End With
        Me.Close()
    End Sub

    Private Sub Rbtclt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Rbtclt.Click
        With FrmAllCorrespondance
            .MdiParent = MenuApplication1
            .Text &= "]<-->[Correspondance Client"
            .Critère = "ACC"
            .Show()
        End With
        Me.Close()
    End Sub

    Private Sub RbtFrss_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RbtFrss.Click
        With FrmAllCorrespondance
            .MdiParent = MenuApplication1
            .Text &= "]<-->[Correspondance Fournisseurs"
            .Critère = "SUP"
            .Show()
        End With
        Me.Close()
    End Sub

    Private Sub RbtCClt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RbtCClt.Click
        With FrmAllCorrespondance
            .MdiParent = MenuApplication1
            .Text &= "]<-->[Correspondance au BC->Client"
            .Critère = "SOR"
            .Show()
        End With
        Me.Close()
    End Sub
End Class
