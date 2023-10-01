Imports System.IO
Public Class Frm_FluxEntrant
    Private Sub FichierJournal_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            LirefichierConfig()
            Me.WindowState = FormWindowState.Maximized
            OuvreLaListedeFichier(PathsfileExport)
        Catch ex As Exception
        End Try
    End Sub
    Private Sub OuvreLaListedeFichier(ByRef Directpath As String)
        Dim i As Integer
        Dim NomFichier As String
        Dim aLines() As String
        Dim jRow As Integer
        DGV.Rows.Clear()
        Try
            If Directory.Exists(Directpath) = True Then
                aLines = Directory.GetFiles(Directpath)
                For i = 0 To UBound(aLines)
                    DGV.RowCount = jRow + 1
                    NomFichier = Trim(aLines(i))
                    Do While InStr(Trim(NomFichier), "\") <> 0
                        NomFichier = Strings.Right(NomFichier, Strings.Len(Trim(NomFichier)) - InStr(Trim(NomFichier), "\"))
                    Loop

                    Select Case NomFichier.Substring(0, 3)
                        Case "CRP"
                            DGV.Rows(jRow).Cells("C1").Value = "Confirmation de Réceptions d'orders d'entrée"
                            DGV.Rows(jRow).Cells("C2").Value = "Exportation"
                            DGV.Rows(jRow).Cells("C3").Value = NomFichier.Substring(0, 3)
                            DGV.Rows(jRow).Cells("C4").Value = NomFichier.Substring(3, 2)
                            DGV.Rows(jRow).Cells("C5").Value = NomFichier
                            jRow = jRow + 1
                        Case "CSO"
                            DGV.Rows(jRow).Cells("C1").Value = "Confirmation de Commande"
                            DGV.Rows(jRow).Cells("C2").Value = "Exportation"
                            DGV.Rows(jRow).Cells("C3").Value = NomFichier.Substring(0, 3)
                            DGV.Rows(jRow).Cells("C4").Value = NomFichier.Substring(3, 2)
                            DGV.Rows(jRow).Cells("C5").Value = NomFichier
                            jRow = jRow + 1
                        Case ""
                    End Select
                Next i
                aLines = Nothing
            Else
                MsgBox("Ce Repertoire n'est pas valide! " & Directpath, MsgBoxStyle.Information, "Repertoire des Fichiers à Traiter")
            End If
        Catch ex As Exception
        End Try
    End Sub
End Class