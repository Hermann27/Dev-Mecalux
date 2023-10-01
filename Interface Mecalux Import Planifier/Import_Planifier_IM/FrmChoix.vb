Public Class FrmChoix
    Public choix As String = ""
    Private Sub RbtArt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RbtArt.Click
        If choix = "Correspondace" Then
            With FrmAllCorrespondance
                .MdiParent = MenuPrincipal
                .MaximizeBox = True
                .Text &= "]<-->[Correspondance Article/Detail Article"
                .Crit�re = "PRO"
                .Show()
            End With
        Else
            With Frm_TemplateALLCorrespondance
                .MdiParent = MenuPrincipal
                .MaximizeBox = True
                .Text &= "]<-->[Template Article/Detail Article"
                .Crit�re = "PRO"
                .Show()
            End With

        End If
        Me.Close()
    End Sub

    Private Sub Rbtclt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Rbtclt.Click
        If choix = "Correspondace" Then
            With FrmAllCorrespondance
                .MdiParent = MenuPrincipal
                .MaximizeBox = True
                .Text &= "]<-->[Correspondance Client"
                .Crit�re = "ACC"
                .Show()
            End With
        Else
            With Frm_TemplateALLCorrespondance
                .MdiParent = MenuPrincipal
                .MaximizeBox = True
                .Text &= "]<-->[Template Client"
                .Crit�re = "ACC"
                .Show()
            End With
        End If
        Me.Close()
    End Sub

    Private Sub RbtFrss_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RbtFrss.Click
        If choix = "Correspondace" Then
            With FrmAllCorrespondance
                .MdiParent = MenuPrincipal
                .MaximizeBox = True
                .Text &= "]<-->[Correspondance Fournisseurs"
                .Crit�re = "SUP"
                .Show()
            End With
        Else
            With Frm_TemplateALLCorrespondance
                .MdiParent = MenuPrincipal
                .MaximizeBox = True
                .Text &= "]<-->[Template Fournisseurs"
                .Crit�re = "SUP"
                .Show()
            End With
        End If
        Me.Close()
    End Sub

    Private Sub RbtCClt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RbtCClt.Click
        If choix = "Correspondace" Then
            With FrmAllCorrespondance
                .MdiParent = MenuPrincipal
                .MaximizeBox = True
                .Text &= "]<-->[Correspondance pour BC->Client"
                .Crit�re = "SOR"
                .Show()
            End With
        Else
            With Frm_TemplateALLCorrespondance
                .MdiParent = MenuPrincipal
                .MaximizeBox = True
                .Text &= "]<-->[Correspondance au BC->Client"
                .Crit�re = "SOR"
                .Show()
            End With
        End If
        Me.Close()
    End Sub

    Private Sub RbtFC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RbtFC.Click
        If choix = "Correspondace" Then
            With FrmAllCorrespondance
                .MdiParent = MenuPrincipal
                .MaximizeBox = True
                .Text &= "]<-->[Correspondance pour CF->Commande"
                .Crit�re = "CSO"
                .Show()
            End With
        Else
            With Frm_TemplateALLCorrespondance
                .MdiParent = MenuPrincipal
                .MaximizeBox = True
                .Text &= "]<-->[Template pour CF->Commande"
                .Crit�re = "CSO"
                .Show()
            End With
        End If
        Me.Close()
    End Sub

    Private Sub RbtCF_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RbtCF.Click
        If choix = "Correspondace" Then
            With FrmAllCorrespondance
                .MdiParent = MenuPrincipal
                .MaximizeBox = True
                .Text &= "]<-->[Correspondance du CF->Commande"
                .Crit�re = "PRE"
                .Show()
            End With
        Else
            With Frm_TemplateALLCorrespondance
                .MdiParent = MenuPrincipal
                .MaximizeBox = True
                .Text &= "]<-->[Template du CF->Commande"
                .Crit�re = "PRE"
                .Show()
            End With
        End If
        Me.Close()
    End Sub

    Private Sub RbtMvt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RbtMvt.Click
        If choix = "Correspondace" Then
            With FrmAllCorrespondance
                .MdiParent = MenuPrincipal
                .MaximizeBox = True
                .Text &= "]<-->[Correspondance au Mvt->E/S"
                .Crit�re = "VST"
                .Show()
            End With
        Else
            With Frm_TemplateALLCorrespondance
                .MdiParent = MenuPrincipal
                .MaximizeBox = True
                .Text &= "]<-->[Template du Mvt->E/S"
                .Crit�re = "VST"
                .Show()
            End With
        End If
        Me.Close()
    End Sub

    Private Sub RbtTdepot_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RbtTdepot.Click
        If choix = "Correspondace" Then
            With FrmAllCorrespondance
                .MdiParent = MenuPrincipal
                .MaximizeBox = True
                .Text &= "]<-->[Correspondance � la CF->R�ception"
                .Crit�re = "CRP"
                .Show()
            End With
        Else
            With Frm_TemplateALLCorrespondance
                .MdiParent = MenuPrincipal
                .MaximizeBox = True
                .Text &= "]<-->[Template � la CF->R�ception"
                .Crit�re = "CRP"
                .Show()
            End With
        End If
        Me.Close()
    End Sub

    Private Sub RtbPseudo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RtbPseudo.Click
      
        If choix = "Correspondace" Then
            With FrmAllCorrespondance
                .MdiParent = MenuPrincipal
                .MaximizeBox = True
                .Text &= "]<-->[Correspondance Pseudo"
                .Crit�re = "ALI"
                .Show()
            End With
        Else
            With Frm_TemplateALLCorrespondance
                .MdiParent = MenuPrincipal
                .MaximizeBox = True
                .Text &= "]<-->[Template Pseudo"
                .Crit�re = "ALI"
                .Show()
            End With
        End If
        Me.Close()
    End Sub
End Class