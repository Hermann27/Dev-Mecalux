Imports System.IO
Public Class frmChargement
    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        If ProgressBar1.Value < 100 Then
            ProgressBar1.Value = ProgressBar1.Value + 1
        End If
        If ProgressBar1.Value > 0 And ProgressBar1.Value < 20 Then
            Label1.Text = "Detection des Parametres De Connexion"
        End If
        If ProgressBar1.Value > 20 And ProgressBar1.Value < 40 Then
            Label1.Text = ""
        End If
        If ProgressBar1.Value > 40 And ProgressBar1.Value < 70 Then
            Label1.Text = "Mise a Jour des Fichiers"
        End If
        If ProgressBar1.Value > 70 And ProgressBar1.Value < 90 Then
            Label1.Text = "Connexion au Serveur"
        End If
        If ProgressBar1.Value > 90 And ProgressBar1.Value <= 100 Then
            Label4.Text = "Chargement des Objets Metiers 100"
        End If
        If ProgressBar1.Value > 95 And ProgressBar1.Value <= 100 Then
            Label3.Text = "Patienter Quelques Secondes ...SVP..."
        End If
        If ProgressBar1.Value = 100 Then
            Timer2.Enabled = False
            Timer3.Enabled = False
        End If
        Label2.Text = ProgressBar1.Value & " %"
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim Reponse As Object
        Dim RepCompta As Object
        Dim RepAccess As Object
        If ProgressBar1.Value = 100 Then
            AccessData = Connected()
            If StatutConsolider = "Oui" Then
                Comptabool = OpenBaseCpta(BaseCpta, PathsBaseCpta, Nom_Util, Mot_Pas)
                Sqlbool = ExcelConnect(NomServersql, NomBaseCpta, Nom_Utilsql, Mot_Passql)
            Else
                Comptabool = True
                Sqlbool = True
            End If
            If AccessData = True Then
                If Sqlbool = True Then
                    If Comptabool = True Then
                        MenuPrincipal.Show()
                        MenuPrincipal.Hide()
                        Me.Close()
                    Else
                        Timer2.Enabled = False
                        Timer3.Enabled = False
                        Timer1.Enabled = False
                        Reponse = MsgBox("Erreur d'ouverture de la base Comptable  " & Chr(13) & "" & Chr(13) & "Modifiez Le Fichier de Configuration", MsgBoxStyle.YesNo, "Connexion à la base Comptable")
                        If MsgBoxResult.Yes = Reponse Then
                            MenuPrincipal.T1.Enabled = False
                            MenuPrincipal.T3.Enabled = False
                            MenuPrincipal.Table1.Enabled = False
                            MenuPrincipal.Show()
                            Me.Close()
                        Else
                            End
                        End If
                    End If
                Else
                    Timer2.Enabled = False
                    Timer3.Enabled = False
                    Timer1.Enabled = False
                    RepCompta = MsgBox("Erreur de Connexion au Serveur SQL  " & Chr(13) & "" & Chr(13) & "Cliquer sur OK pour Continuer" & Chr(13) & "" & Chr(13) & "                 Ou Sur  " & Chr(13) & "" & Chr(13) & "Annuler Pour Quitter Le Programme", MsgBoxStyle.YesNo, "Connexion au Serveur SQL")
                    If RepCompta = MsgBoxResult.Yes Then
                        If Comptabool = True Then
                            MenuPrincipal.Show()
                            Me.Close()
                        Else
                            MenuPrincipal.T1.Enabled = False
                            MenuPrincipal.T3.Enabled = False
                            MenuPrincipal.Table1.Visible = False
                            Timer2.Enabled = False
                            Timer3.Enabled = False
                            Timer1.Enabled = False
                            Reponse = MsgBox("Erreur d'ouverture de la base Comptable  " & Chr(13) & "" & Chr(13) & "Modifiez Le Fichier de Configuration", MsgBoxStyle.YesNo, "Connexion à la base Comptable")
                            If MsgBoxResult.Yes = Reponse Then
                                MenuPrincipal.Show()
                                Me.Close()
                            Else
                                End
                            End If
                        End If
                    Else
                        End
                    End If
                End If
            Else
                MenuPrincipal.T1.Enabled = False
                MenuPrincipal.T3.Enabled = False
                MenuPrincipal.Table1.Visible = False
                Timer2.Enabled = False
                Timer3.Enabled = False
                Timer1.Enabled = False
                RepAccess = MsgBox("Erreur d'ouverture de la base Microsoft Access  " & Chr(13) & "" & Chr(13) & "Modifiez Le Fichier de Configuration", MsgBoxStyle.YesNo, "Connexion à la base Microsoft Access")
                If MsgBoxResult.Yes = RepAccess Then
                    MenuPrincipal.Show()
                    Me.Close()
                Else
                    End
                End If
            End If
        End If
    End Sub
    Private Sub Timer3_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer3.Tick
        Timer1.Enabled = True
        Timer2.Enabled = True
    End Sub

    Private Sub frmChargement_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Timer1.Enabled = False
        Timer2.Enabled = False
        Try
            Call LirefichierConfig()
        Catch ex As Exception

        End Try

    End Sub
End Class
