Imports System.Data.OleDb
Public Class PlanificationDate
    Private Sub PlanificationDate_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Me.Text = "PlanificationTraitement" Then
            If Idexe >= 0 Then
                If IsDate(Trim(PlanificationTraitement.DataListeIntegrer.Rows(Idexe).Cells("Critere1").Value)) = True Then
                    TimeDebut.Value = PlanificationTraitement.DataListeIntegrer.Rows(Idexe).Cells("Critere1").Value
                Else
                    TimeDebut.Value = DateAndTime.Now
                End If
                TimeFin.Value = Strings.FormatDateTime((Date.DaysInMonth(DateAndTime.Year(TimeDebut.Value), DateAndTime.Month(TimeDebut.Value)) & "/" & DateAndTime.Month(TimeDebut.Value) & "/" & DateAndTime.Year(TimeDebut.Value)), DateFormat.ShortDate)

                If IsDate(Trim(PlanificationTraitement.DataListeIntegrer.Rows(Idexe).Cells("Critere2").Value)) = True Then
                    TimeFin.Value = PlanificationTraitement.DataListeIntegrer.Rows(Idexe).Cells("Critere2").Value
                End If
            End If
        Else
            If Me.Text = "Planification" Then
                If IsDate(Trim(PlanificationSpecial.dgvTraitementEnr.Rows(Idexe).Cells("Critere1").Value)) = True Then
                    TimeDebut.Value = PlanificationSpecial.dgvTraitementEnr.Rows(Idexe).Cells("Critere1").Value
                Else
                    TimeDebut.Value = DateAndTime.Now
                End If
                TimeFin.Value = Strings.FormatDateTime((Date.DaysInMonth(DateAndTime.Year(TimeDebut.Value), DateAndTime.Month(TimeDebut.Value)) & "/" & DateAndTime.Month(TimeDebut.Value) & "/" & DateAndTime.Year(TimeDebut.Value)), DateFormat.ShortDate)

                If IsDate(Trim(PlanificationSpecial.dgvTraitementEnr.Rows(Idexe).Cells("Critere2").Value)) = True Then
                    TimeFin.Value = PlanificationSpecial.dgvTraitementEnr.Rows(Idexe).Cells("Critere2").Value
                End If
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub BTvalider_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTvalider.Click
        If Me.Text = "PlanificationTraitement" Then
            If Idexe >= 0 Then
                If TimeDebut.Value <= TimeFin.Value Then
                    PlanificationTraitement.DataListeIntegrer.Rows(Idexe).Cells("Critere1").Value = Strings.FormatDateTime(TimeDebut.Value, DateFormat.ShortDate)
                    PlanificationTraitement.DataListeIntegrer.Rows(Idexe).Cells("Critere2").Value = Strings.FormatDateTime(TimeFin.Value, DateFormat.ShortDate)
                    Me.Close()
                Else
                    MsgBox("Date de debut doit être inferieur à celle de Fin", MsgBoxStyle.Information, "Selection date")
                End If
            End If
        Else
            If Me.Text = "Planification" Then
                If Idexe >= 0 Then
                    If TimeDebut.Value <= TimeFin.Value Then
                        PlanificationSpecial.dgvTraitementEnr.Rows(Idexe).Cells("Critere1").Value = Strings.FormatDateTime(TimeDebut.Value, DateFormat.ShortDate)
                        PlanificationSpecial.dgvTraitementEnr.Rows(Idexe).Cells("Critere2").Value = Strings.FormatDateTime(TimeFin.Value, DateFormat.ShortDate)
                        Me.Close()
                    Else
                        MsgBox("Date de debut doit être inferieur à celle de Fin", MsgBoxStyle.Information, "Selection date")
                    End If
                End If
            End If
        End If
    End Sub
End Class