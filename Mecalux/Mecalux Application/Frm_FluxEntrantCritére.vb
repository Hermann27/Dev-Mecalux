Imports System.IO
Public Class Frm_FluxEntrantCritére
    Public Critere As String = ""

    Private Sub FichierJournal_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            LirefichierConfig()
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
                    If Critere = NomFichier.Substring(0, 3) Then
                        Select Case NomFichier.Substring(0, 3)
                            Case "CRP"
                                DGV.Rows(jRow).Cells("C1").Value = "Confirmation de Réceptions d'orders d'entrée"
                                DGV.Rows(jRow).Cells("C2").Value = "Exportation"
                                DGV.Rows(jRow).Cells("C3").Value = NomFichier.Substring(0, 3)
                                DGV.Rows(jRow).Cells("C4").Value = NomFichier.Substring(3, 2)
                                DGV.Rows(jRow).Cells("C5").Value = NomFichier
                                DGV.Rows(jRow).Cells("C6").Value = False
                                DGV.Rows(jRow).Cells("C7").Value = My.Resources.btFermer22
                                DGV.Rows(jRow).Cells("C8").Value = aLines(i)
                                jRow = jRow + 1
                            Case "CSO"
                                DGV.Rows(jRow).Cells("C1").Value = "Confirmation de Commande"
                                DGV.Rows(jRow).Cells("C2").Value = "Exportation"
                                DGV.Rows(jRow).Cells("C3").Value = NomFichier.Substring(0, 3)
                                DGV.Rows(jRow).Cells("C4").Value = NomFichier.Substring(3, 2)
                                DGV.Rows(jRow).Cells("C5").Value = NomFichier
                                DGV.Rows(jRow).Cells("C6").Value = False
                                DGV.Rows(jRow).Cells("C7").Value = My.Resources.btFermer22
                                DGV.Rows(jRow).Cells("C8").Value = aLines(i)
                                jRow = jRow + 1
                            Case ""
                        End Select
                    End If
                Next i
                aLines = Nothing
            Else
                MsgBox("Ce Repertoire n'est pas valide! " & Directpath, MsgBoxStyle.Information, "Repertoire des Fichiers à Traiter")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        Dim i As Integer
        For i = 0 To DGV.RowCount - 1
            DGV.Rows(i).Cells("C6").Value = True
        Next i
    End Sub

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click
        Dim i As Integer
        For i = 0 To DGV.RowCount - 1
            DGV.Rows(i).Cells("C6").Value = False
        Next i
    End Sub

    Private Sub RadButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadButton1.Click
        Dim i As Integer
        For i = 0 To DGV.RowCount - 1
            DGV.Rows(i).Cells("C7").Value = My.Resources.Checked
            EcritureFlux(DGV.Rows(i).Cells("C8").Value)
        Next i
    End Sub
    Public Function EcritureFlux(ByVal Chemin As Object) As Boolean
        Try
            Dim aRows() As String = Nothing
            If GetArrayFile(Chemin, aRows) IsNot Nothing Then
                aRows = GetArrayFile(Chemin, aRows)
                For i As Integer = 0 To UBound(aRows)
                    Dim Ligne As String = aRows(i)

                Next
            End If
        Catch ex As Exception

        End Try
        Return False
    End Function

End Class