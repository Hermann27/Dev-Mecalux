Imports System.Xml
Imports System.IO
Imports System
Imports System.Data.OleDb
Public Class AFormatStock
    Public xdoc As XmlDocument
    Public racine As XmlElement
    Public nodelist As XmlNodeList
    Public nodelist2 As XmlNodeList
    Private Sub BT_Del_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Del.Click
        DataListeFormat.Rows.Add()
    End Sub

    Private Sub BT_ADD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_ADD.Click
        Dim first As Integer
        Dim last As Integer
        first = DataListeFormat.Rows.GetFirstRow(DataGridViewElementStates.Displayed)
        last = DataListeFormat.Rows.GetLastRow(DataGridViewElementStates.Displayed)
        If last >= 0 Then
            If last - first >= 0 Then
                DataListeFormat.Rows.RemoveAt(DataListeFormat.CurrentRow.Index)
            End If
        End If
    End Sub
    Private Sub AFormatStock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim BaseBool As Boolean
        BaseBool = Connected()
        DataListeFormat.Rows.Clear()
        Me.WindowState = FormWindowState.Maximized
    End Sub

    Private Sub AjouterUnFormat_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        BT_Quit.Location = New Point(SplitContainer1.Panel2.Width / 2 - 20, SplitContainer1.Panel2.Height / 2 - 10)
        BT_Save.Location = New Point(SplitContainer1.Panel2.Width / 2 + 100, SplitContainer1.Panel2.Height / 2 - 10)

    End Sub
    Private Sub BT_Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Save.Click
        Dim n As Integer
        Dim EstCreation As Boolean = False
        Dim SaveFormat As String
        Dim OleCommandSave As OleDbCommand
        Dim OleAdaptaterFormat As OleDbDataAdapter
        Dim OleFormatDataset As DataSet
        Dim OledatableFormat As DataTable
        Try
            For n = 0 To DataListeFormat.RowCount - 1
                If DataListeFormat.Rows(n).Cells("TypeFormat").Value <> "" Then
                    OleAdaptaterFormat = New OleDbDataAdapter("select * from WIS_FORMAT where NomFormat='" & Trim(DataListeFormat.Rows(n).Cells("Format").Value) & "' and Type='" & Renvoietypeformat(DataListeFormat.Rows(n).Cells("TypeFormat").Value) & "'", OleConnenection)
                    OleFormatDataset = New DataSet
                    OleAdaptaterFormat.Fill(OleFormatDataset)
                    OledatableFormat = OleFormatDataset.Tables(0)
                    If OledatableFormat.Rows.Count = 0 Then
                        EstCreation = True
                        SaveFormat = "Insert Into WIS_FORMAT (Chemin,NomFormat,Type) VALUES ('" & DataListeFormat.Rows(n).Cells("Chemin").Value & "','" & DataListeFormat.Rows(n).Cells("Format").Value & "','" & Renvoietypeformat(DataListeFormat.Rows(n).Cells("TypeFormat").Value) & "')"
                        OleCommandSave = New OleDbCommand(SaveFormat)
                        OleCommandSave.Connection = OleConnenection
                        OleCommandSave.ExecuteNonQuery()
                    End If
                Else
                    MsgBox("Le Type de Format est Obligatoire", MsgBoxStyle.Information, "Ajouter un Format")
                End If
            Next n
            If EstCreation = True Then
                MsgBox("Format Crée", MsgBoxStyle.Information, "Ajouter un Format")
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub BT_Quit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Quit.Click
        Me.Close()
    End Sub

    Private Sub DataListeFormat_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataListeFormat.CellContentClick
        Dim i As Integer
        If e.RowIndex >= 0 Then
            If DataListeFormat.Columns(e.ColumnIndex).Name = "Fichier" Then
                If DataListeFormat.Rows(e.RowIndex).Cells("TypeFormat").Value = "" Then
                    MsgBox("Veuillez choisir d'abord le Type de format!", MsgBoxStyle.Information, "Selection Fichier Format")
                    Exit Sub
                Else
                    If DataListeFormat.Columns(e.ColumnIndex).Name = "Fichier" Then
                        OpenFileFormat.Filter = "Fichier texte (*.Xml)|*.Xml"
                        OpenFileFormat.FileName = Nothing
                        If OpenFileFormat.ShowDialog = Windows.Forms.DialogResult.OK Then
                            xdoc = New XmlDocument
                            xdoc.Load(Trim(OpenFileFormat.FileName))
                            racine = xdoc.DocumentElement
                            nodelist = racine.ChildNodes
                            For i = 0 To nodelist.Count - 1
                                If Trim(nodelist.ItemOf(i).Name) = "TypeFormat" Then
                                    If Renvoietypeformat(DataListeFormat.Rows(e.RowIndex).Cells("TypeFormat").Value) = nodelist.ItemOf(i).InnerText Then
                                        i = nodelist.Count - 1
                                        DataListeFormat.Rows(e.RowIndex).Cells("Chemin").Value = OpenFileFormat.FileName
                                        DataListeFormat.Rows(e.RowIndex).Cells("Format").Value = Trim(System.IO.Path.GetFileName(OpenFileFormat.FileName))
                                    Else
                                        MsgBox("Ce fichier ne correspond pas au format sélectionné!", MsgBoxStyle.Information, "Selection Fichier Format")
                                        i = nodelist.Count - 1
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If
            End If
        End If
    End Sub
End Class