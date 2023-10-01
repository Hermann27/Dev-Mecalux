Imports System.Data.OleDb
Public Class FrmCorrespondance

    Private Sub BT_FicCpta_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_FicCpta.Click
        DataListeSchema.Rows.Add()
    End Sub

    Private Sub BtnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDelete.Click
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
    Private Sub EnregistrerLeSchema()
        Dim n As Integer
        Dim OleCommandEnreg As OleDbCommand
        Dim Insert As Boolean = False
        Dim InsertTable As Boolean = False
        Dim Insertion As String
        Dim InsertionTable As String
        If Trim(txtCde.Text) <> "" And Trim(txttableSage.Text) <> "" And Trim(txtintulé.Text) <> "" And Trim(ComboType.Text) <> "" Then
            InsertionTable = "Insert Into P_TABLECORRESP (CodeTbls,Libelle,NomTableSage,TypeEchange) VALUES ('" & Join(Split(txtCde.Text, "'"), "''") & "','" & Join(Split(txtintulé.Text, "'"), "''") & "','" & Join(Split(txttableSage.Text, "'"), "''") & "','" & Join(Split(ComboType.Text, "'"), "''") & "')"
        Else
            MsgBox("Veuillez renseigner tous les champs SVP", MsgBoxStyle.Information, "Parametrage des Correspondance")
            Exit Sub
        End If
        If DataListeSchema.RowCount >= 1 Then
            For n = 0 To DataListeSchema.RowCount - 1
                If Trim(txtCde.Text) <> "" And Trim(txttableSage.Text) <> "" And Trim(txtintulé.Text) <> "" And Trim(ComboType.Text) <> "" Then
                    If Trim(DataListeSchema.Rows(n).Cells("Cols").Value) <> "" _
                         And Trim(DataListeSchema.Rows(n).Cells("Format").Value) <> "" Then
                        If Trim(DataListeSchema.Rows(n).Cells("Entete").Value) = "True" And Trim(DataListeSchema.Rows(n).Cells("Ligne").Value) = "True" Then
                            MsgBox("Faite un choix entre Entête|Ligne", MsgBoxStyle.Information, "Parametrage des Correspondance")
                            Exit Sub
                        End If
                        If Trim(DataListeSchema.Rows(n).Cells("Entete").Value) = "True" Then
                            If Trim(DataListeSchema.Rows(n).Cells("InfosLibre").Value) = "True" Then
                                Insertion = "Insert Into P_COLONNEST (Cols,ordre,description,Format,PositionG,DefaultValue,ChampSage,InfosLibre,CodeTbls,Entete,Ligne) VALUES ('" & Join(Split(DataListeSchema.Rows(n).Cells("Cols").Value, "'"), "''") & "'," & n + 1 & ",'" & Join(Split(DataListeSchema.Rows(n).Cells("Desc").Value, "'"), "''") & "','" & Join(Split(DataListeSchema.Rows(n).Cells("Format").Value, "'"), "''") & "','" & Join(Split(DataListeSchema.Rows(n).Cells("PositionG").Value, "'"), "''") & "','" & Join(Split(DataListeSchema.Rows(n).Cells("DefaultValue").Value, "'"), "''") & "','" & Join(Split(DataListeSchema.Rows(n).Cells("ChampSage").Value, "'"), "''") & "'," & True & ",'" & Join(Split(txtCde.Text, "'"), "''") & "'," & True & "," & False & ")"
                            Else
                                Insertion = "Insert Into P_COLONNEST (Cols,ordre,description,Format,PositionG,DefaultValue,ChampSage,InfosLibre,CodeTbls,Entete,Ligne) VALUES ('" & Join(Split(DataListeSchema.Rows(n).Cells("Cols").Value, "'"), "''") & "'," & n + 1 & ",'" & Join(Split(DataListeSchema.Rows(n).Cells("Desc").Value, "'"), "''") & "','" & Join(Split(DataListeSchema.Rows(n).Cells("Format").Value, "'"), "''") & "','" & Join(Split(DataListeSchema.Rows(n).Cells("PositionG").Value, "'"), "''") & "','" & Join(Split(DataListeSchema.Rows(n).Cells("DefaultValue").Value, "'"), "''") & "','" & Join(Split(DataListeSchema.Rows(n).Cells("ChampSage").Value, "'"), "''") & "'," & False & ",'" & Join(Split(txtCde.Text, "'"), "''") & "'," & True & "," & False & ")"
                            End If
                        ElseIf Trim(DataListeSchema.Rows(n).Cells("Ligne").Value) = "True" Then
                            If Trim(DataListeSchema.Rows(n).Cells("InfosLibre").Value) = "True" Then
                                Insertion = "Insert Into P_COLONNEST (Cols,ordre,description,Format,PositionG,DefaultValue,ChampSage,InfosLibre,CodeTbls,Entete,Ligne) VALUES ('" & Join(Split(DataListeSchema.Rows(n).Cells("Cols").Value, "'"), "''") & "'," & n + 1 & ",'" & Join(Split(DataListeSchema.Rows(n).Cells("Desc").Value, "'"), "''") & "','" & Join(Split(DataListeSchema.Rows(n).Cells("Format").Value, "'"), "''") & "','" & Join(Split(DataListeSchema.Rows(n).Cells("PositionG").Value, "'"), "''") & "','" & Join(Split(DataListeSchema.Rows(n).Cells("DefaultValue").Value, "'"), "''") & "','" & Join(Split(DataListeSchema.Rows(n).Cells("ChampSage").Value, "'"), "''") & "'," & True & ",'" & Join(Split(txtCde.Text, "'"), "''") & "'," & False & "," & True & ")"
                            Else
                                Insertion = "Insert Into P_COLONNEST (Cols,ordre,description,Format,PositionG,DefaultValue,ChampSage,InfosLibre,CodeTbls,Entete,Ligne) VALUES ('" & Join(Split(DataListeSchema.Rows(n).Cells("Cols").Value, "'"), "''") & "'," & n + 1 & ",'" & Join(Split(DataListeSchema.Rows(n).Cells("Desc").Value, "'"), "''") & "','" & Join(Split(DataListeSchema.Rows(n).Cells("Format").Value, "'"), "''") & "','" & Join(Split(DataListeSchema.Rows(n).Cells("PositionG").Value, "'"), "''") & "','" & Join(Split(DataListeSchema.Rows(n).Cells("DefaultValue").Value, "'"), "''") & "','" & Join(Split(DataListeSchema.Rows(n).Cells("ChampSage").Value, "'"), "''") & "'," & False & ",'" & Join(Split(txtCde.Text, "'"), "''") & "'," & False & "," & True & ")"
                            End If
                        Else
                            MsgBox("Faite un choix entre Entête/Ligne", MsgBoxStyle.Information, "Parametrage des Correspondance")
                            Exit Sub
                        End If

                        If InsertTable = False Then
                            OleCommandEnreg = New OleDbCommand(InsertionTable)
                            OleCommandEnreg.Connection = OleConnenection
                            OleCommandEnreg.ExecuteNonQuery()
                            InsertTable = True
                        End If

                        OleCommandEnreg = New OleDbCommand(Insertion)
                        OleCommandEnreg.Connection = OleConnenection
                        OleCommandEnreg.ExecuteNonQuery()
                        Insert = True
                    Else
                        MsgBox("Renseigner les champs obligatoire SVP {Colonne EasyWMS,Description,Format,Position Gauche}", MsgBoxStyle.Information, "Parametrage des Correspondance")
                        Exit Sub
                    End If
                End If

            Next n
            If Insert = True Then
                MsgBox("Insertion Reussie", MsgBoxStyle.Information, "Parametrage des Correspondance")
                Inititialiseur()
                DataListeSchema.Rows.Clear()
            End If
        End If
    End Sub
    Public Sub Inititialiseur()
        txtCde.Text = ""
        txtintulé.Text = ""
        ComboType.Text = ""
        txttableSage.Text = ""
    End Sub
    Private Sub FrmCorrespondance_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LirefichierConfig()
    End Sub

    Private Sub BtnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave.Click
        If Connected() Then
            EnregistrerLeSchema()
        Else
            MsgBox("Erreur de connexion au fichier access")
        End If
    End Sub
End Class
