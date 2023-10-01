Imports System.Data.OleDb
Public Class Frm_ToPosting

    Private Sub Frm_ToPosting_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Connected() = True Then
        End If
    End Sub
    Private Sub MiseàjourLeSchema()
        Dim n As Integer
        Dim OleAdaptaterEnreg As OleDbDataAdapter
        Dim OleEnregDataset As DataSet
        Dim OledatableEnreg As DataTable
        Dim OleCommandEnreg As OleDbCommand
        Dim Insert As Boolean = False
        Dim Insertion As String
        Try
            If DataListeIntegrer.RowCount >= 0 Then
                For n = 0 To DataListeIntegrer.RowCount - 1
                    OleAdaptaterEnreg = New OleDbDataAdapter("select * From ARTICLE where  Fournisseur='" & DataListeIntegrer.Rows(n).Cells("CodeFo1").Value & "'and Code_Article_Fo='" & DataListeIntegrer.Rows(n).Cells("ArticleFo1").Value & "'", OleConnenection)
                    OleEnregDataset = New DataSet
                    OleAdaptaterEnreg.Fill(OleEnregDataset)
                    OledatableEnreg = OleEnregDataset.Tables(0)
                    If OledatableEnreg.Rows.Count <> 0 Then
                        Insertion = "UPDATE ARTICLE SET Code_Article_Dis='" & DataListeIntegrer.Rows(n).Cells("ArticleDis1").Value & "'where  Fournisseur='" & DataListeIntegrer.Rows(n).Cells("CodeFo1").Value & "'and Code_Article_Fo='" & DataListeIntegrer.Rows(n).Cells("ArticleFo1").Value & "'"
                        OleCommandEnreg = New OleDbCommand(Insertion)
                        OleCommandEnreg.Connection = OleConnenection
                        OleCommandEnreg.ExecuteNonQuery()
                        Insert = True
                    Else
                    End If
                    Me.Refresh()
                Next n
                If Insert = True Then
                    MsgBox("Mise à Jour Reussie", MsgBoxStyle.Information, "Mise à Jour des Articles")
                End If
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        MiseàjourLeSchema()
        AfficheSchemasIntegrer()
    End Sub
    Private Sub AfficheSchemasIntegrer()
        Dim i As Integer
        Dim OleAdaptaterschema As OleDbDataAdapter
        Dim OleSchemaDataset As DataSet
        Dim OledatableSchema As DataTable
        Frm_CorArticle.DataListeIntegrer.Rows.Clear()
        OleAdaptaterschema = New OleDbDataAdapter("select * from ARTICLE", OleConnenection)
        OleSchemaDataset = New DataSet
        OleAdaptaterschema.Fill(OleSchemaDataset)
        OledatableSchema = OleSchemaDataset.Tables(0)
        Frm_CorArticle.DataListeIntegrer.RowCount = OledatableSchema.Rows.Count
        For i = 0 To OledatableSchema.Rows.Count - 1
            Frm_CorArticle.DataListeIntegrer.Rows(i).Cells("CodeFo1").Value = OledatableSchema.Rows(i).Item("Fournisseur")
            Frm_CorArticle.DataListeIntegrer.Rows(i).Cells("ArticleFo1").Value = OledatableSchema.Rows(i).Item("Code_Article_Fo")
            Frm_CorArticle.DataListeIntegrer.Rows(i).Cells("ArticleDis1").Value = OledatableSchema.Rows(i).Item("Code_Article_Dis")
        Next i
    End Sub
End Class