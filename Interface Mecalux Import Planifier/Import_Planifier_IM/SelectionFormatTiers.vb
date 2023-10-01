Imports System.Data.OleDb
Imports System.IO
Public Class SelectionFormatTiers
    Private Sub AfficheSchemasConso()
        Try
            Dim i As Integer
            Dim OleAdaptaterschema As OleDbDataAdapter
            Dim OleSchemaDataset As DataSet
            Dim OledatableSchema As DataTable
            DataListeIntegrer.Rows.Clear()
            OleAdaptaterschema = New OleDbDataAdapter("select * from FORMAT ORDER BY NomFormat ASC", OleConnenection)
            OleSchemaDataset = New DataSet
            OleAdaptaterschema.Fill(OleSchemaDataset)
            OledatableSchema = OleSchemaDataset.Tables(0)
            DataListeIntegrer.RowCount = OledatableSchema.Rows.Count
            For i = 0 To OledatableSchema.Rows.Count - 1
                DataListeIntegrer.Rows(i).Cells("Catego1").Value = OledatableSchema.Rows(i).Item("NomFormat")
                DataListeIntegrer.Rows(i).Cells("Compte1").Value = Afficheauuser(OledatableSchema.Rows(i).Item("Type"))
                DataListeIntegrer.Rows(i).Cells("Dossier").Value = OledatableSchema.Rows(i).Item("Chemin")
            Next i
        Catch ex As Exception
            MsgBox("Message Systeme: " & ex.Message, MsgBoxStyle.Information, "Sélection des Fichiers Formats")
        End Try
    End Sub
    Private Sub BT_Quit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Quit.Click
        Me.Close()
    End Sub
    Private Sub SelectionFormatTiers_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Me.Name = "EXTRACTION" Then
            AfficheSchemasExtraction()
        Else
            AfficheSchemasConso()
        End If
    End Sub
    Private Sub AfficheSchemasExtraction()
        Try
            Dim i As Integer
            Dim OleAdaptaterschema As OleDbDataAdapter
            Dim OleSchemaDataset As DataSet
            Dim OledatableSchema As DataTable
            DataListeIntegrer.Rows.Clear()
            OleAdaptaterschema = New OleDbDataAdapter("select * from WET_FORMAT ORDER BY NomFormat ASC", OleConnenection)
            OleSchemaDataset = New DataSet
            OleAdaptaterschema.Fill(OleSchemaDataset)
            OledatableSchema = OleSchemaDataset.Tables(0)
            DataListeIntegrer.RowCount = OledatableSchema.Rows.Count
            For i = 0 To OledatableSchema.Rows.Count - 1
                DataListeIntegrer.Rows(i).Cells("Catego1").Value = OledatableSchema.Rows(i).Item("NomFormat")
                DataListeIntegrer.Rows(i).Cells("Compte1").Value = OledatableSchema.Rows(i).Item("Type")
                DataListeIntegrer.Rows(i).Cells("Dossier").Value = OledatableSchema.Rows(i).Item("Chemin")
            Next i
        Catch ex As Exception
            MsgBox("Message Systeme: " & ex.Message, MsgBoxStyle.Information, "Sélection des Fichiers Formats")
        End Try
    End Sub
    Private Sub BT_Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Save.Click
        If Me.Name = "EXTRACTION" Then
            Dim test As Windows.Forms.DataGridViewRow
            For Each test In DataListeIntegrer.SelectedRows
                Dim OleAdaptaterInt As OleDbDataAdapter
                Dim OleIntDataset As DataSet
                Dim OledatableInt As DataTable
                OleAdaptaterInt = New OleDbDataAdapter("select * from WET_FORMAT Where Chemin='" & Trim(test.Cells("Dossier").Value) & "'", OleConnenection)
                OleIntDataset = New DataSet
                OleAdaptaterInt.Fill(OleIntDataset)
                OledatableInt = OleIntDataset.Tables(0)
                If OledatableInt.Rows.Count <> 0 Then
                    SchematextractionTiers.DataListeSchema.Rows(Idexe).Cells("NomFormat").Value = Trim(OledatableInt.Rows(0).Item("NomFormat"))
                    SchematextractionTiers.DataListeSchema.Rows(Idexe).Cells("TypeFormat").Value = Trim(OledatableInt.Rows(0).Item("Type"))
                    SchematextractionTiers.DataListeSchema.Rows(Idexe).Cells("Chemin").Value = Trim(OledatableInt.Rows(0).Item("Chemin"))
                    Me.Close()
                End If
            Next
        Else
            Dim test As Windows.Forms.DataGridViewRow
            For Each test In DataListeIntegrer.SelectedRows
                Dim OleAdaptaterInt As OleDbDataAdapter
                Dim OleIntDataset As DataSet
                Dim OledatableInt As DataTable
                OleAdaptaterInt = New OleDbDataAdapter("select * from FORMAT Where Chemin='" & Trim(test.Cells("Dossier").Value) & "'", OleConnenection)
                OleIntDataset = New DataSet
                OleAdaptaterInt.Fill(OleIntDataset)
                OledatableInt = OleIntDataset.Tables(0)
                If OledatableInt.Rows.Count <> 0 Then
                    SchematintegrerTiers.DataListeSchema.Rows(Idexe).Cells("NomFormat").Value = Trim(OledatableInt.Rows(0).Item("NomFormat"))
                    SchematintegrerTiers.DataListeSchema.Rows(Idexe).Cells("TypeFormat").Value = Afficheauuser(Trim(OledatableInt.Rows(0).Item("Type")))
                    SchematintegrerTiers.DataListeSchema.Rows(Idexe).Cells("Chemin").Value = Trim(OledatableInt.Rows(0).Item("Chemin"))
                    Me.Close()
                End If
            Next
        End If
    End Sub
    Private Sub DataListeIntegrer_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataListeIntegrer.CellDoubleClick
        If e.RowIndex >= 0 Then
            If Me.Name = "EXTRACTION" Then
                Dim OleAdaptaterInt As OleDbDataAdapter
                Dim OleIntDataset As DataSet
                Dim OledatableInt As DataTable
                OleAdaptaterInt = New OleDbDataAdapter("select * from WET_FORMAT Where Chemin='" & Trim(DataListeIntegrer.Rows(e.RowIndex).Cells("Dossier").Value) & "' ", OleConnenection)
                OleIntDataset = New DataSet
                OleAdaptaterInt.Fill(OleIntDataset)
                OledatableInt = OleIntDataset.Tables(0)
                If OledatableInt.Rows.Count <> 0 Then
                    SchematextractionTiers.DataListeSchema.Rows(Idexe).Cells("NomFormat").Value = Trim(OledatableInt.Rows(0).Item("NomFormat"))
                    SchematextractionTiers.DataListeSchema.Rows(Idexe).Cells("TypeFormat").Value = Trim(OledatableInt.Rows(0).Item("Type"))
                    SchematextractionTiers.DataListeSchema.Rows(Idexe).Cells("Chemin").Value = Trim(OledatableInt.Rows(0).Item("Chemin"))
                End If
                Me.Close()
            Else
                Dim OleAdaptaterInt As OleDbDataAdapter
                Dim OleIntDataset As DataSet
                Dim OledatableInt As DataTable
                OleAdaptaterInt = New OleDbDataAdapter("select * from FORMAT Where Chemin='" & Trim(DataListeIntegrer.Rows(e.RowIndex).Cells("Dossier").Value) & "' ", OleConnenection)
                OleIntDataset = New DataSet
                OleAdaptaterInt.Fill(OleIntDataset)
                OledatableInt = OleIntDataset.Tables(0)
                If OledatableInt.Rows.Count <> 0 Then
                    SchematintegrerTiers.DataListeSchema.Rows(Idexe).Cells("NomFormat").Value = Trim(OledatableInt.Rows(0).Item("NomFormat"))
                    SchematintegrerTiers.DataListeSchema.Rows(Idexe).Cells("TypeFormat").Value = Afficheauuser(Trim(OledatableInt.Rows(0).Item("Type")))
                    SchematintegrerTiers.DataListeSchema.Rows(Idexe).Cells("Chemin").Value = Trim(OledatableInt.Rows(0).Item("Chemin"))
                End If
                Me.Close()
            End If
        End If
    End Sub
End Class