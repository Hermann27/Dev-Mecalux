Imports System.Data.OleDb
Public Class PlanificationHeure
    Private Sub PlanificationHeure_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Me.Text = "PlanificationTraitement" Then
            If Idexe >= 0 Then
                If Trim(PlanificationTraitement.DataListeIntegrer.Rows(Idexe).Cells("Heure1").Value) <> "" Then
                    TimeDebut.Text = Strings.FormatDateTime(PlanificationTraitement.DataListeIntegrer.Rows(Idexe).Cells("Heure1").Value, DateFormat.LongTime)
                Else
                    TimeDebut.Text = DateAndTime.TimeString
                End If
                If Trim(PlanificationTraitement.DataListeIntegrer.Rows(Idexe).Cells("Heure2").Value) <> "" Then
                    TimeFin.Text = Strings.FormatDateTime(PlanificationTraitement.DataListeIntegrer.Rows(Idexe).Cells("Heure2").Value, DateFormat.LongTime)
                Else
                    TimeFin.Text = DateAndTime.TimeString
                End If
            End If
        Else
            If Me.Text = "Planification" Then
                If Trim(PlanificationSpecial.dgvTraitementEnr.Rows(Idexe).Cells("Heure1").Value) <> "" Then
                    TimeDebut.Text = Strings.FormatDateTime(PlanificationSpecial.dgvTraitementEnr.Rows(Idexe).Cells("Heure1").Value, DateFormat.LongTime)
                Else
                    TimeDebut.Text = DateAndTime.TimeString
                End If
                If Trim(PlanificationSpecial.dgvTraitementEnr.Rows(Idexe).Cells("Heure2").Value) <> "" Then
                    TimeFin.Text = Strings.FormatDateTime(PlanificationSpecial.dgvTraitementEnr.Rows(Idexe).Cells("Heure2").Value, DateFormat.LongTime)
                Else
                    TimeFin.Text = DateAndTime.TimeString
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
                If Strings.FormatDateTime(TimeDebut.Value, DateFormat.ShortTime) <= Strings.FormatDateTime(TimeFin.Value, DateFormat.ShortTime) Then
                    If Format(DateAndTime.Hour(TimeDebut.Value), "00") & ":" & Format(DateAndTime.Minute(TimeDebut.Value), "00") & ":" & Format(DateAndTime.Minute(TimeDebut.Value), "00") = "00:00:00" Then
                        PlanificationTraitement.DataListeIntegrer.Rows(Idexe).Cells("Heure1").Value = ""
                    Else
                        PlanificationTraitement.DataListeIntegrer.Rows(Idexe).Cells("Heure1").Value = Strings.FormatDateTime(TimeDebut.Value, DateFormat.LongTime) ' Format(DateAndTime.Hour(TimeDebut.Value), "00") & ":" & Format(DateAndTime.Minute(TimeDebut.Value), "00") & ":" & Format(DateAndTime.Minute(TimeDebut.Value), "00")
                    End If
                    If Format(DateAndTime.Hour(TimeFin.Value), "00") & ":" & Format(DateAndTime.Minute(TimeFin.Value), "00") & ":" & Format(DateAndTime.Minute(TimeFin.Value), "00") = "00:00:00" Then
                        PlanificationTraitement.DataListeIntegrer.Rows(Idexe).Cells("Heure2").Value = ""
                    Else
                        PlanificationTraitement.DataListeIntegrer.Rows(Idexe).Cells("Heure2").Value = Strings.FormatDateTime(TimeFin.Value, DateFormat.LongTime) ' Format(DateAndTime.Hour(TimeFin.Value), "00") & ":" & Format(DateAndTime.Minute(TimeFin.Value), "00") & ":" & Format(DateAndTime.Minute(TimeFin.Value), "00")
                    End If
                    Me.Close()
                Else
                    MsgBox("Heure de debut doit être inferieur à celle de Fin", MsgBoxStyle.Information, "Selection d'heure")
                End If

            End If
        Else
            If Me.Text = "Planification" Then
                If Idexe >= 0 Then
                    If Strings.FormatDateTime(TimeDebut.Value, DateFormat.ShortTime) <= Strings.FormatDateTime(TimeFin.Value, DateFormat.ShortTime) Then
                        If Format(DateAndTime.Hour(TimeDebut.Value), "00") & ":" & Format(DateAndTime.Minute(TimeDebut.Value), "00") & ":" & Format(DateAndTime.Minute(TimeDebut.Value), "00") = "00:00:00" Then
                            PlanificationSpecial.dgvTraitementEnr.Rows(Idexe).Cells("Heure1").Value = ""
                        Else
                            PlanificationSpecial.dgvTraitementEnr.Rows(Idexe).Cells("Heure1").Value = Strings.FormatDateTime(TimeDebut.Value, DateFormat.LongTime) '   Format(DateAndTime.Hour(TimeDebut.Value), "00") & ":" & Format(DateAndTime.Minute(TimeDebut.Value), "00") & ":" & Format(DateAndTime.Minute(TimeDebut.Value), "00")
                        End If
                        If Format(DateAndTime.Hour(TimeFin.Value), "00") & ":" & Format(DateAndTime.Minute(TimeFin.Value), "00") & ":" & Format(DateAndTime.Minute(TimeFin.Value), "00") = "00:00:00" Then
                            PlanificationSpecial.dgvTraitementEnr.Rows(Idexe).Cells("Heure2").Value = ""
                        Else
                            PlanificationSpecial.dgvTraitementEnr.Rows(Idexe).Cells("Heure2").Value = Strings.FormatDateTime(TimeFin.Value, DateFormat.LongTime) ' Format(DateAndTime.Hour(TimeFin.Value), "00") & ":" & Format(DateAndTime.Minute(TimeFin.Value), "00") & ":" & Format(DateAndTime.Minute(TimeFin.Value), "00")
                        End If
                        Me.Close()
                    Else
                        MsgBox("Heure de debut doit être inferieur à celle de Fin", MsgBoxStyle.Information, "Selection d'heure")
                    End If
                End If
            End If
        End If
    End Sub
End Class