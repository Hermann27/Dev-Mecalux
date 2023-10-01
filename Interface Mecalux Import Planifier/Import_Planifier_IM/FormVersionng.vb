Imports System.Data.OleDb
Public Class FormVersionng
    Public Sub AfficheSchemasConso()
        Try
            Dim i As Integer
            Dim OleAdaptaterschema As OleDbDataAdapter
            Dim OleSchemaDataset As DataSet
            Dim OledatableSchema As DataTable
            DataListeIntegrer.Rows.Clear()
            OleAdaptaterschema = New OleDbDataAdapter("select * from P_TABLECORRESP ", OleConnenectionArticle)
            OleSchemaDataset = New DataSet
            OleAdaptaterschema.Fill(OleSchemaDataset)
            OledatableSchema = OleSchemaDataset.Tables(0)
            DataListeIntegrer.RowCount = OledatableSchema.Rows.Count
            For i = 0 To OledatableSchema.Rows.Count - 1
                DataListeIntegrer.Rows(i).Cells("Table").Value = OledatableSchema.Rows(i).Item("Libelle")
                DataListeIntegrer.Rows(i).Cells("codeTable").Value = OledatableSchema.Rows(i).Item("CodeTbls")
                DataListeIntegrer.Rows(i).Cells("version").Value = OledatableSchema.Rows(i).Item("version")
                DataListeIntegrer.Rows(i).Cells("Mode").Value = OledatableSchema.Rows(i).Item("TypeEchange")
            Next i
        Catch ex As Exception
        End Try
    End Sub
    Public Function RenvoiVersion(ByVal codeTable As String) As String
        Try
            Dim OleAdaptaterschema As OleDbDataAdapter
            Dim OleSchemaDataset As DataSet
            Dim OledatableSchema As DataTable
            OleAdaptaterschema = New OleDbDataAdapter("select * from P_TABLECORRESP WHERE CodeTbls='" & codeTable & "'", OleConnenectionArticle)
            OleSchemaDataset = New DataSet
            OleAdaptaterschema.Fill(OleSchemaDataset)
            OledatableSchema = OleSchemaDataset.Tables(0)
            Return OledatableSchema.Rows(0).Item("version")
        Catch ex As Exception
            Return "Nand"
        End Try
    End Function
    Private Sub FormVersionng_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            LirefichierConfig()
            If Connected() = True Then
                AfficheSchemasConso
            End If
        Catch ex As Exception
        End Try
    End Sub
    Public code As String = ""
    Private Sub DataListeIntegrer_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataListeIntegrer.CellClick
        If e.RowIndex >= 0 Then
            Try
                code = DataListeIntegrer.Rows(e.RowIndex).Cells("codeTable").Value
                txtVersion.Text = RenvoiVersion(DataListeIntegrer.Rows(e.RowIndex).Cells("codeTable").Value)
                lblinfos.Text = code
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If
    End Sub
    Private Sub BTupdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTupdate.Click
        Try
            Dim Chaine As String = "UPDATE P_TABLECORRESP SET version='" & txtVersion.Text & "' WHERE CodeTbls='" & code & "'"
            Dim MaCommande As New OleDbCommand(Chaine, OleConnenectionArticle)
            MaCommande.ExecuteNonQuery()
            MsgBox("Mise à Jour Reussie", MsgBoxStyle.Information, "Mise à Jour Version du Fichier")
        Catch ex As Exception
            MsgBox("Echéc de Mise à Jour " & ex.Message, MsgBoxStyle.Information, "Mise à Jour Version du Fichier")
        End Try
    End Sub
End Class