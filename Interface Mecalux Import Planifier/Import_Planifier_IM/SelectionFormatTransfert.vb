Imports Objets100Lib
Imports System
Imports System.Data.OleDb
Imports System.Collections
Imports System.Windows.Forms
Imports System.IO
Imports System.Xml
Imports Microsoft.VisualBasic
Public Class SelectionFormatTransfert
    Public xdoc As XmlDocument
    Public racine As XmlElement
    Public nodelist As XmlNodeList
    Public nodelist2 As XmlNodeList
    Private Sub AfficheSchemasIntegration()
        Try
            Dim i, j As Integer
            Dim OleAdaptaterschema As OleDbDataAdapter
            Dim OleSchemaDataset As DataSet
            Dim OledatableSchema As DataTable
            DataListeIntegrer.Rows.Clear()
            OleAdaptaterschema = New OleDbDataAdapter("select * from WIT_FORMAT ORDER BY NomFormat ASC", OleConnenection)
            OleSchemaDataset = New DataSet
            OleAdaptaterschema.Fill(OleSchemaDataset)
            OledatableSchema = OleSchemaDataset.Tables(0)
            j = 0
            For i = 0 To OledatableSchema.Rows.Count - 1
                DataListeIntegrer.RowCount = j + 1
                If AffichFormatModifiable(OledatableSchema.Rows(i).Item("Chemin")) = "Modification" Then
                    If ExistePieceModifiable(OledatableSchema.Rows(i).Item("Chemin")) = True Then
                        DataListeIntegrer.Rows(j).Cells("Catego1").Value = OledatableSchema.Rows(i).Item("NomFormat")
                        DataListeIntegrer.Rows(j).Cells("Compte1").Value = Afficheauuser(OledatableSchema.Rows(i).Item("Type"))
                        DataListeIntegrer.Rows(j).Cells("Dossier").Value = OledatableSchema.Rows(i).Item("Chemin")
                        DataListeIntegrer.Rows(j).Cells("Mode").Value = AffichFormatModifiable(OledatableSchema.Rows(i).Item("Chemin"))
                        j = j + 1
                    Else
                        MsgBox("Il n'existe aucun identifiant de N°pièce Coché pour le format en modification : " & OledatableSchema.Rows(i).Item("Chemin"), MsgBoxStyle.Information, " Cocher Colonne pièce")
                    End If
                Else
                    DataListeIntegrer.Rows(j).Cells("Catego1").Value = OledatableSchema.Rows(i).Item("NomFormat")
                    DataListeIntegrer.Rows(j).Cells("Compte1").Value = Afficheauuser(OledatableSchema.Rows(i).Item("Type"))
                    DataListeIntegrer.Rows(j).Cells("Dossier").Value = OledatableSchema.Rows(i).Item("Chemin")
                    DataListeIntegrer.Rows(j).Cells("Mode").Value = AffichFormatModifiable(OledatableSchema.Rows(i).Item("Chemin"))
                    j = j + 1
                End If
            Next i
        Catch ex As Exception
            MsgBox("Message Systeme: " & ex.Message, MsgBoxStyle.Information, "Sélection des Fichiers Formats")
        End Try
    End Sub
    Private Sub BT_Quit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Quit.Click
        Me.Close()
    End Sub

    Private Sub SelectionFormatTransfert_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.Name = ""
    End Sub
    Private Sub SelectionFormatTransfert_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        AfficheSchemasIntegration()
    End Sub
    Private Sub BT_Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Save.Click
        Dim test As Windows.Forms.DataGridViewRow
        For Each test In DataListeIntegrer.SelectedRows
            Dim OleAdaptaterInt As OleDbDataAdapter
            Dim OleIntDataset As DataSet
            Dim OledatableInt As DataTable
            OleAdaptaterInt = New OleDbDataAdapter("select * from WIT_FORMAT Where Chemin='" & Trim(test.Cells("Dossier").Value) & "'", OleConnenection)
            OleIntDataset = New DataSet
            OleAdaptaterInt.Fill(OleIntDataset)
            OledatableInt = OleIntDataset.Tables(0)
            If OledatableInt.Rows.Count <> 0 Then
                SchematintegrerTransfert.DataListeSchema.Rows(Idexe).Cells("NomFormat").Value = Trim(OledatableInt.Rows(0).Item("NomFormat"))
                SchematintegrerTransfert.DataListeSchema.Rows(Idexe).Cells("TypeFormat").Value = Afficheauuser(Trim(OledatableInt.Rows(0).Item("Type")))
                SchematintegrerTransfert.DataListeSchema.Rows(Idexe).Cells("Chemin").Value = Trim(OledatableInt.Rows(0).Item("Chemin"))
                SchematintegrerTransfert.DataListeSchema.Rows(Idexe).Cells("Mode").Value = AffichFormatModifiable(Trim(OledatableInt.Rows(0).Item("Chemin")))
                Me.Close()
            End If
        Next
    End Sub
    Private Sub DataListeIntegrer_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataListeIntegrer.CellDoubleClick
        If e.RowIndex >= 0 Then
            Dim OleAdaptaterInt As OleDbDataAdapter
            Dim OleIntDataset As DataSet
            Dim OledatableInt As DataTable
            OleAdaptaterInt = New OleDbDataAdapter("select * from WIT_FORMAT Where Chemin='" & Trim(DataListeIntegrer.Rows(e.RowIndex).Cells("Dossier").Value) & "' ", OleConnenection)
            OleIntDataset = New DataSet
            OleAdaptaterInt.Fill(OleIntDataset)
            OledatableInt = OleIntDataset.Tables(0)
            If OledatableInt.Rows.Count <> 0 Then
                SchematintegrerTransfert.DataListeSchema.Rows(Idexe).Cells("NomFormat").Value = Trim(OledatableInt.Rows(0).Item("NomFormat"))
                SchematintegrerTransfert.DataListeSchema.Rows(Idexe).Cells("TypeFormat").Value = Afficheauuser(Trim(OledatableInt.Rows(0).Item("Type")))
                SchematintegrerTransfert.DataListeSchema.Rows(Idexe).Cells("Chemin").Value = Trim(OledatableInt.Rows(0).Item("Chemin"))
                SchematintegrerTransfert.DataListeSchema.Rows(Idexe).Cells("Mode").Value = AffichFormatModifiable(Trim(OledatableInt.Rows(0).Item("Chemin")))
            End If
            Me.Close()
        End If
    End Sub
    Private Function AffichFormatModifiable(ByVal CheminFichier) As String
        Dim ModeFormat As Object = Nothing
        Dim i As Integer
        Try
            If File.Exists(CheminFichier) = True Then
                Dim FileXml As New XmlTextReader(Trim(CheminFichier))
                xdoc = New XmlDocument
                xdoc.Load(Trim(CheminFichier))
                racine = xdoc.DocumentElement
                nodelist = racine.ChildNodes
                For i = 0 To nodelist.Count - 1
                    If Trim(nodelist.ItemOf(i).Name) = "MODE_FORMAT" Then
                        ModeFormat = nodelist.ItemOf(i).InnerText
                        Exit For
                    End If
                Next i
            Else
                MsgBox("Nom du Format inexistant!", MsgBoxStyle.Information, "Format d'integration")
            End If
        Catch ex As Exception

        End Try
        AffichFormatModifiable = ModeFormat
    End Function
    Private Function ExistePieceModifiable(ByVal CheminFichier) As Boolean
        Dim ExistePiece As Boolean = False
        Dim i As Integer
        Try
            If File.Exists(CheminFichier) = True Then
                Dim FileXml As New XmlTextReader(Trim(CheminFichier))
                xdoc = New XmlDocument
                xdoc.Load(Trim(CheminFichier))
                racine = xdoc.DocumentElement
                nodelist = racine.ChildNodes
                For i = 0 To nodelist.Count - 1
                    If Trim(nodelist.ItemOf(i).Name) = "Piece" Then
                        If nodelist.ItemOf(i).InnerText = "oui" Then
                            ExistePiece = True
                            Exit For
                        End If
                    End If
                Next i
            Else
                MsgBox("Nom du Format inexistant!", MsgBoxStyle.Information, "Format d'integration")
            End If
        Catch ex As Exception

        End Try
        ExistePieceModifiable = ExistePiece
    End Function
End Class