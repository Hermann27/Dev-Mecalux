Imports System.Data.OleDb
Imports System.IO
Public Class SelectionFiltre
    Private Sub AfficheColonneFiltre()
        Try
            Dim i As Integer
            Dim OleAdaptaterschema As OleDbDataAdapter
            Dim OleSchemaDataset As DataSet
            Dim OledatableSchema As DataTable
            DataListeIntegrer.Rows.Clear()
            OleAdaptaterschema = New OleDbDataAdapter("select TOP 1 * from " & Trim(Txttable.Text) & " ", BaseSQLConnection)
            OleSchemaDataset = New DataSet
            OleAdaptaterschema.Fill(OleSchemaDataset)
            OledatableSchema = OleSchemaDataset.Tables(0)
            DataListeIntegrer.RowCount = OledatableSchema.Columns.Count
            For i = 0 To OledatableSchema.Columns.Count - 1
                DataListeIntegrer.Rows(i).Cells("Filtre").Value = OledatableSchema.Columns(i).ColumnName
                DataListeIntegrer.Rows(i).Cells("Selection").Value = False
            Next i
        Catch ex As Exception
            MsgBox("Message Systeme: " & ex.Message, MsgBoxStyle.Information, "Sélection du Filtre")
        End Try
    End Sub
    Private Sub BT_Quit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Quit.Click
        Me.Close()
    End Sub
    Private Sub SelectionFormatTiers_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If BaseSQLConnection.State = ConnectionState.Open Then
            AfficheColonneFiltre()
        End If
    End Sub
    Private Sub BT_Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Save.Click
        If Me.Name = "SelectionFiltreCA" Then
            For i As Integer = 0 To DataListeIntegrer.RowCount - 1
                If DataListeIntegrer.Rows(i).Cells("Selection").Value = True Then
                    SchematintegrerCA.DataListeSchema.Rows(Idexe).Cells("FeuilleExcel").Value = DataListeIntegrer.Rows(i).Cells("Filtre").Value
                    Me.Close()
                    Exit For
                End If
            Next i
        Else
            If Me.Name = "SelectionFiltreCA1" Then
                For i As Integer = 0 To DataListeIntegrer.RowCount - 1
                    If DataListeIntegrer.Rows(i).Cells("Selection").Value = True Then
                        SchematintegrerCA.DataListeIntegrer.Rows(Idexe).Cells("FeuilleExcel1").Value = DataListeIntegrer.Rows(i).Cells("Filtre").Value
                        Me.Close()
                        Exit For
                    End If
                Next i
            Else
                If Me.Name = "SelectionFiltreEC" Then
                    For i As Integer = 0 To DataListeIntegrer.RowCount - 1
                        If DataListeIntegrer.Rows(i).Cells("Selection").Value = True Then
                            SchematintegrerEcriture.DataListeSchema.Rows(Idexe).Cells("FeuilleExcel").Value = DataListeIntegrer.Rows(i).Cells("Filtre").Value
                            Me.Close()
                            Exit For
                        End If
                    Next i
                Else
                    If Me.Name = "SelectionFiltreEC1" Then
                        For i As Integer = 0 To DataListeIntegrer.RowCount - 1
                            If DataListeIntegrer.Rows(i).Cells("Selection").Value = True Then
                                SchematintegrerEcriture.DataListeIntegrer.Rows(Idexe).Cells("FeuilleExcel1").Value = DataListeIntegrer.Rows(i).Cells("Filtre").Value
                                Me.Close()
                                Exit For
                            End If
                        Next i
                    Else
                        If Me.Name = "SelectionFiltreTI" Then
                            For i As Integer = 0 To DataListeIntegrer.RowCount - 1
                                If DataListeIntegrer.Rows(i).Cells("Selection").Value = True Then
                                    SchematintegrerTiers.DataListeSchema.Rows(Idexe).Cells("FeuilleExcel").Value = DataListeIntegrer.Rows(i).Cells("Filtre").Value
                                    Me.Close()
                                    Exit For
                                End If
                            Next i
                        Else
                            If Me.Name = "SelectionFiltreTI1" Then
                                For i As Integer = 0 To DataListeIntegrer.RowCount - 1
                                    If DataListeIntegrer.Rows(i).Cells("Selection").Value = True Then
                                        SchematintegrerTiers.DataListeIntegrer.Rows(Idexe).Cells("FeuilleExcel1").Value = DataListeIntegrer.Rows(i).Cells("Filtre").Value
                                        Me.Close()
                                        Exit For
                                    End If
                                Next i
                            Else
                                If Me.Name = "SelectionFiltreMVT" Then
                                    For i As Integer = 0 To DataListeIntegrer.RowCount - 1
                                        If DataListeIntegrer.Rows(i).Cells("Selection").Value = True Then
                                            SchematintegrerMvt.DataListeSchema.Rows(Idexe).Cells("Feuille_Excel").Value = DataListeIntegrer.Rows(i).Cells("Filtre").Value
                                            Me.Close()
                                            Exit For
                                        End If
                                    Next i
                                Else
                                    If Me.Name = "SelectionFiltreMVT1" Then
                                        For i As Integer = 0 To DataListeIntegrer.RowCount - 1
                                            If DataListeIntegrer.Rows(i).Cells("Selection").Value = True Then
                                                SchematintegrerMvt.DataListeIntegrer.Rows(Idexe).Cells("Feuille_Excel1").Value = DataListeIntegrer.Rows(i).Cells("Filtre").Value
                                                Me.Close()
                                                Exit For
                                            End If
                                        Next i
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Sub
End Class