Imports System.Data.OleDb
Public Class Frm_EtatDeConexion

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Timer1.Enabled = False
        Timer2.Enabled = False
        Call LirefichierConfig()
        If Connected() = True Then
            VerificationConnexion()
        End If
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        If ProgressBar1.Value < 100 Then
            ProgressBar1.Value = ProgressBar1.Value + 1
        End If
        If ProgressBar1.Value > 0 And ProgressBar1.Value < 20 Then
        End If
        If ProgressBar1.Value > 20 And ProgressBar1.Value < 40 Then
        End If
        If ProgressBar1.Value > 40 And ProgressBar1.Value < 70 Then
        End If
        If ProgressBar1.Value > 70 And ProgressBar1.Value < 90 Then
        End If
        If ProgressBar1.Value > 90 And ProgressBar1.Value <= 100 Then
        End If
        If ProgressBar1.Value > 95 And ProgressBar1.Value <= 100 Then
        End If
        If ProgressBar1.Value = 100 Then
            Timer2.Enabled = False
            Timer3.Enabled = False
        End If
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If ProgressBar1.Value = 100 Then
            Timer2.Enabled = False
            Timer3.Enabled = False
            Timer1.Enabled = False
            'Me.Close()
        End If
    End Sub

    Private Sub Timer3_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer3.Tick
        Timer1.Enabled = True
        Timer2.Enabled = True
    End Sub
    Private Sub VerificationConnexion()
        Dim OleAdaptaterConso As OleDbDataAdapter
        Dim OleConsoDataset As DataSet
        Dim OledatableConso As DataTable
        Dim i As Integer
        OleAdaptaterConso = New OleDbDataAdapter("select * from PARAMETRE ", OleConnenection)
        OleConsoDataset = New DataSet
        OleAdaptaterConso.Fill(OleConsoDataset)
        OledatableConso = OleConsoDataset.Tables(0)
        If OledatableConso.Rows.Count <> 0 Then
            For i = 0 To OledatableConso.Rows.Count - 1
                DataConnexion.RowCount = i + 1
                If TestConnected(OledatableConso.Rows(i).Item("BaseDonnee"), OledatableConso.Rows(i).Item("MotPas"), OledatableConso.Rows(i).Item("NomUser"), OledatableConso.Rows(i).Item("Serveur")) = True Then
                    DataConnexion.Rows(i).Cells("Connexion").Value = "Connexion à la Société  " & OledatableConso.Rows(i).Item("Societe")
                    DataConnexion.Rows(i).Cells("Reussie").Value = "Reussie....."
                Else
                    DataConnexion.Rows(i).Cells("Connexion").Value = "Connexion à la Société  " & OledatableConso.Rows(i).Item("Societe")
                    DataConnexion.Rows(i).Cells("Reussie").Value = "Echec....."
                End If
            Next i
        End If
    End Sub
    Public Function TestConnected(ByRef FichierSageCpta As String, ByRef Mot_Psql As String, ByRef Nom_Utsql As String, ByRef servers As String) As Boolean
        Dim OleConnectedTest As OleDbConnection
        Try
            OleConnectedTest = New OleDbConnection("provider=SQLOLEDB;UID=" & Trim(Nom_Utsql) & ";Pwd=" & Trim(Mot_Psql) & ";Initial Catalog=" & Trim(FichierSageCpta) & ";Data Source=" & Trim(servers) & "")
            OleConnectedTest.Open()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class
