Imports System.Data.OleDb
Imports System.IO
Public Class Frm_CorArticle
    Private Sub Article_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Connected() = True Then
            AfficheSchemasIntegrer()
            Me.WindowState = FormWindowState.Maximized
        End If
    End Sub

    Private Sub AfficheSchemasIntegrer()
        Dim i As Integer
        Dim OleAdaptaterschema As OleDbDataAdapter
        Dim OleSchemaDataset As DataSet
        Dim OledatableSchema As DataTable
        DataListeIntegrer.Rows.Clear()

        OleAdaptaterschema = New OleDbDataAdapter("select * from ARTICLE", OleConnenection)
        OleSchemaDataset = New DataSet
        OleAdaptaterschema.Fill(OleSchemaDataset)
        OledatableSchema = OleSchemaDataset.Tables(0)
        DataListeIntegrer.RowCount = OledatableSchema.Rows.Count
        For i = 0 To OledatableSchema.Rows.Count - 1
            DataListeIntegrer.Rows(i).Cells("CodeFo1").Value = OledatableSchema.Rows(i).Item("Fournisseur")
            DataListeIntegrer.Rows(i).Cells("ArticleFo1").Value = OledatableSchema.Rows(i).Item("Code_Article_Fo")
            DataListeIntegrer.Rows(i).Cells("ArticleDis1").Value = OledatableSchema.Rows(i).Item("Code_Article_Dis")
        Next i
    End Sub
    Private Sub RechercheSchemasIntegrer()
        Dim i As Integer
        Dim OleAdaptaterschema As OleDbDataAdapter
        Dim OleSchemaDataset As DataSet
        Dim OledatableSchema As DataTable
        DataListeIntegrer.Rows.Clear()
        OleAdaptaterschema = New OleDbDataAdapter("select * from ARTICLE Where Code_Article_Dis='" & Trim(TextRecher.Text) & "'", OleConnenection)
        OleSchemaDataset = New DataSet
        OleAdaptaterschema.Fill(OleSchemaDataset)
        OledatableSchema = OleSchemaDataset.Tables(0)
        DataListeIntegrer.RowCount = OledatableSchema.Rows.Count
        For i = 0 To OledatableSchema.Rows.Count - 1
            DataListeIntegrer.Rows(i).Cells("CodeFo1").Value = OledatableSchema.Rows(i).Item("Fournisseur")
            DataListeIntegrer.Rows(i).Cells("ArticleFo1").Value = OledatableSchema.Rows(i).Item("Code_Article_Fo")
            DataListeIntegrer.Rows(i).Cells("ArticleDis1").Value = OledatableSchema.Rows(i).Item("Code_Article_Dis")
        Next i
    End Sub
    Private Sub Delete_DataListeSch()
        Dim i As Integer
        Dim OleAdaptaterDelete As OleDbDataAdapter
        Dim OleDeleteDataset As DataSet
        Dim OledatableDelete As DataTable
        Dim OleCommandDelete As OleDbCommand
        Dim DelFile As String
        For i = 0 To DataListeIntegrer.RowCount - 1
            If DataListeIntegrer.Rows(i).Cells("Supprimer").Value = True Then
                OleAdaptaterDelete = New OleDbDataAdapter("select * from ARTICLE where  Fournisseur='" & DataListeIntegrer.Rows(i).Cells("CodeFo1").Value & "'and Code_Article_Fo='" & DataListeIntegrer.Rows(i).Cells("ArticleFo1").Value & "'", OleConnenection)
                OleDeleteDataset = New DataSet
                OleAdaptaterDelete.Fill(OleDeleteDataset)
                OledatableDelete = OleDeleteDataset.Tables(0)

                If OledatableDelete.Rows.Count <> 0 Then
                    DelFile = "Delete From ARTICLE where Fournisseur='" & DataListeIntegrer.Rows(i).Cells("CodeFo1").Value & "'and Code_Article_Fo='" & DataListeIntegrer.Rows(i).Cells("ArticleFo1").Value & "'"
                    OleCommandDelete = New OleDbCommand(DelFile)
                    OleCommandDelete.Connection = OleConnenection
                    OleCommandDelete.ExecuteNonQuery()
                End If
            End If
        Next i
        AfficheSchemasIntegrer()
    End Sub
    Private Sub EnregistrerLeSchema()
        Dim n As Integer
        Dim OleAdaptaterEnreg As OleDbDataAdapter
        Dim OleEnregDataset As DataSet
        Dim OledatableEnreg As DataTable
        Dim OleCommandEnreg As OleDbCommand
        Dim Insert As Boolean = False
        Dim Insertion As String
        If DataListeSchema.RowCount >= 1 Then
            For n = 0 To DataListeSchema.RowCount - 1
                OleAdaptaterEnreg = New OleDbDataAdapter("select * From ARTICLE WHERE  Fournisseur='" & DataListeSchema.Rows(n).Cells("CodeFo").Value & "'and Code_Article_Fo='" & DataListeSchema.Rows(n).Cells("ArticleFo").Value & "'", OleConnenection)
                OleEnregDataset = New DataSet
                OleAdaptaterEnreg.Fill(OleEnregDataset)
                OledatableEnreg = OleEnregDataset.Tables(0)
                If OledatableEnreg.Rows.Count <> 0 Then
                Else
                    If Trim(DataListeSchema.Rows(n).Cells("CodeFo").Value) <> "" And Trim(DataListeSchema.Rows(n).Cells("ArticleFo").Value) <> "" Then
                        Insertion = "Insert Into ARTICLE (Fournisseur,Code_Article_Fo,Code_Article_Dis) VALUES ('" & DataListeSchema.Rows(n).Cells("CodeFo").Value & "','" & DataListeSchema.Rows(n).Cells("ArticleFo").Value & "','" & DataListeSchema.Rows(n).Cells("ArticleDis").Value & "')"
                        OleCommandEnreg = New OleDbCommand(Insertion)
                        OleCommandEnreg.Connection = OleConnenection
                        OleCommandEnreg.ExecuteNonQuery()
                        Insert = True
                    End If
                End If
            Next n
            If Insert = True Then
                MsgBox("Insertion Reussie", MsgBoxStyle.Information, "Insertion des ARTICLES")
                DataListeSchema.Rows.Clear()
            End If
        End If
    End Sub
    Private Sub BT_Quit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Quit.Click
        Me.Close()
    End Sub
    Private Sub BT_Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Save.Click
        EnregistrerLeSchema()
        AfficheSchemasIntegrer()
    End Sub
    Private Sub BT_Delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Delete.Click
        Delete_DataListeSch()
    End Sub
    Private Sub BT_DelRow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_DelRow.Click
        Dim first As Integer
        Dim last As Integer
        first = DataListeSchema.Rows.GetFirstRow(DataGridViewElementStates.Displayed)
        last = DataListeSchema.Rows.GetLastRow(DataGridViewElementStates.Displayed)
        If last >= 0 Then
            If last - first >= 0 Then
                DataListeSchema.Rows.RemoveAt(DataListeSchema.CurrentRow.Index)
            End If
        End If
    End Sub
    Private Sub BT_ADD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_ADD.Click
        DataListeSchema.Rows.Add()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        LectureXml.ShowDialog()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        RechercheSchemasIntegrer()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        AfficheSchemasIntegrer()
    End Sub

    Private Sub BT_SelAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_SelAll.Click
        Dim i As Integer
        For i = 0 To DataListeIntegrer.RowCount - 1
            DataListeIntegrer.Rows(i).Cells("Supprimer").Value = True
        Next i

    End Sub

    Private Sub BT_DelAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_DelAll.Click
        Dim i As Integer
        For i = 0 To DataListeIntegrer.RowCount - 1
            DataListeIntegrer.Rows(i).Cells("Supprimer").Value = False
        Next i

    End Sub
    Private Sub DataListeIntegrer_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataListeIntegrer.CellClick
        If e.RowIndex >= 0 Then
            If DataListeIntegrer.Columns(e.ColumnIndex).Name = "Modifier" Then
                Frm_ToPosting.DataListeIntegrer.RowCount = 1
                Frm_ToPosting.DataListeIntegrer.Rows(0).Cells("CodeFo1").Value = DataListeIntegrer.Rows(e.RowIndex).Cells("CodeFo1").Value
                Frm_ToPosting.DataListeIntegrer.Rows(0).Cells("ArticleFo1").Value = DataListeIntegrer.Rows(e.RowIndex).Cells("ArticleFo1").Value
                Frm_ToPosting.DataListeIntegrer.Rows(0).Cells("ArticleDis1").Value = DataListeIntegrer.Rows(e.RowIndex).Cells("ArticleDis1").Value
                Frm_ToPosting.ShowDialog()
            End If
        End If
    End Sub
End Class