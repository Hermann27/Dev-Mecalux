Public Class Frm_OptionGrid
    Public choix As String = ""
    Public ShowsForm As String = ""
    Private Sub BtnAjouter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAjouter.Click
        Dim Position As Integer = DGVCE.Rows(DGVCE.CurrentRow.Index).Cells(1).Value
        Select Case choix
            Case "Entete"
                Select Case ShowsForm
                    Case "1"
                        Frm_FluxEntrantCrit�re.DataListeIntegrer.Columns.Item(Position).Visible = True
                        Frm_FluxEntrantCrit�re.Vsate = False
                        Frm_FluxEntrantCrit�re.PicLigne_Click(sender, e)
                    Case "2"
                        Frm_ConfirmationReception.DataListeIntegrer.Columns.Item(Position).Visible = True
                        Frm_ConfirmationReception.Vsate = False
                        Frm_ConfirmationReception.PicLigne_Click(sender, e)
                    Case "3"
                        Frm_MvtE_S.DataListeIntegrer.Columns.Item(Position).Visible = True
                        Frm_MvtE_S.Vsate = False
                        Frm_MvtE_S.PicLigne_Click(sender, e)
                End Select
            Case "Ligne"
                Select Case ShowsForm
                    Case "1"
                        Frm_FluxEntrantCrit�re.DataListeIntegrerLigne.Columns.Item(Position).Visible = True
                        Frm_FluxEntrantCrit�re.Vsate = False
                        Frm_FluxEntrantCrit�re.PictureBox1_Click(sender, e)
                    Case "2"
                        Frm_ConfirmationReception.DataListeIntegrerLigne.Columns.Item(Position).Visible = True
                        Frm_ConfirmationReception.Vsate = False
                        Frm_ConfirmationReception.PictureBox1_Click(sender, e)
                End Select
            Case "SousLigne"
                Select Case ShowsForm
                    Case "1"
                        Frm_FluxEntrantCrit�re.DataListeIntegrerD�tailLigne.Columns.Item(Position).Visible = True
                        Frm_FluxEntrantCrit�re.Vsate = False
                        Frm_FluxEntrantCrit�re.PictureBox2_Click(sender, e)
                    Case "2"
                        Frm_ConfirmationReception.DataListeIntegrerD�tailLigne.Columns.Item(Position).Visible = True
                        Frm_ConfirmationReception.Vsate = False
                        Frm_ConfirmationReception.PictureBox2_Click(sender, e)
                End Select
        End Select
    End Sub

    Private Sub Frm_OptionGrid_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Frm_FluxEntrantCrit�re.Vsate = True
    End Sub

    Private Sub BtnSup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSup.Click
        Dim Position As Integer = DGVCV.Rows(DGVCV.CurrentRow.Index).Cells(1).Value
        Select Case choix
            Case "Entete"
                Select Case ShowsForm
                    Case "1"
                        Frm_FluxEntrantCrit�re.DataListeIntegrer.Columns.Item(Position).Visible = False
                        Frm_FluxEntrantCrit�re.Vsate = False
                        Frm_FluxEntrantCrit�re.PicLigne_Click(sender, e)
                    Case "2"
                        Frm_ConfirmationReception.DataListeIntegrer.Columns.Item(Position).Visible = False
                        Frm_ConfirmationReception.Vsate = False
                        Frm_ConfirmationReception.PicLigne_Click(sender, e)
                    Case "3"
                        Frm_MvtE_S.DataListeIntegrer.Columns.Item(Position).Visible = False
                        Frm_MvtE_S.Vsate = False
                        Frm_MvtE_S.PicLigne_Click(sender, e)
                End Select
            Case "Ligne"
                Select Case ShowsForm
                    Case "1"
                        Frm_FluxEntrantCrit�re.DataListeIntegrerLigne.Columns.Item(Position).Visible = False
                        Frm_FluxEntrantCrit�re.Vsate = False
                        Frm_FluxEntrantCrit�re.PictureBox1_Click(sender, e)
                    Case "2"
                        Frm_ConfirmationReception.DataListeIntegrerLigne.Columns.Item(Position).Visible = False
                        Frm_ConfirmationReception.Vsate = False
                        Frm_ConfirmationReception.PictureBox1_Click(sender, e)
                End Select
            Case "SousLigne"
                Select Case ShowsForm
                    Case "1"
                        Frm_FluxEntrantCrit�re.DataListeIntegrerD�tailLigne.Columns.Item(Position).Visible = False
                        Frm_FluxEntrantCrit�re.Vsate = False
                        Frm_FluxEntrantCrit�re.PictureBox2_Click(sender, e)
                    Case "2"
                        Frm_ConfirmationReception.DataListeIntegrerD�tailLigne.Columns.Item(Position).Visible = False
                        Frm_ConfirmationReception.Vsate = False
                        Frm_ConfirmationReception.PictureBox2_Click(sender, e)
                End Select
        End Select
    End Sub
End Class