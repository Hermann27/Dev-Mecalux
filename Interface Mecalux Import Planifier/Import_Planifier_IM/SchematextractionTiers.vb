Imports System.Data.OleDb
Imports System.IO
Imports System.Net.NetworkInformation
Public Class SchematextractionTiers
    Public FlagConnection As OleDbConnection
    Private Sub SchematextractionTiers_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Maximized
        Try
            If Connected() = True Then
                AfficheSchemasIntegrer()
                AfficheSocieteCpta()
                AfficheSocieteCial()
            End If
        Catch ex As Exception

        End Try
    End Sub
    Public Function SociéteConnection(ByRef Server As String, ByRef basededonne As String, ByRef utilisateur As String, ByRef motdepasse As String) As Boolean
        Try
            FlagConnection = New OleDbConnection("provider=SQLOLEDB;UID=" & Trim(utilisateur) & ";Pwd=" & Trim(motdepasse) & ";Initial Catalog=" & Trim(basededonne) & ";Data Source=" & Trim(Server) & "")
            FlagConnection.Open()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Private Sub AfficheSocieteCial()
        Dim OleAdaptater As OleDbDataAdapter
        Dim OleAfficheDataset As DataSet
        Dim Oledatable As DataTable
        Dim i As Integer
        BaseCial.Items.Clear()
        Try
            OleAdaptater = New OleDbDataAdapter("select * from PARAMETRE where nomtype='COMMERCIAL'", OleConnenection)
            OleAfficheDataset = New DataSet
            OleAdaptater.Fill(OleAfficheDataset)
            Oledatable = OleAfficheDataset.Tables(0)
            For i = 0 To Oledatable.Rows.Count - 1
                If Trim(Oledatable.Rows(i).Item("Societe")) <> "" Then
                    BaseCial.Items.AddRange(New String() {Oledatable.Rows(i).Item("Societe")})
                End If
            Next i
        Catch ex As Exception

        End Try
    End Sub
    Private Sub AfficheSocieteCpta()
        Dim OleAdaptater As OleDbDataAdapter
        Dim OleAfficheDataset As DataSet
        Dim Oledatable As DataTable
        Dim i As Integer
        BaseCpta.Items.Clear()
        Try
            OleAdaptater = New OleDbDataAdapter("select * from PARAMETRE where nomtype='COMPTABILITE'", OleConnenection)
            OleAfficheDataset = New DataSet
            OleAdaptater.Fill(OleAfficheDataset)
            Oledatable = OleAfficheDataset.Tables(0)
            For i = 0 To Oledatable.Rows.Count - 1
                If Trim(Oledatable.Rows(i).Item("Societe")) <> "" Then
                    BaseCpta.Items.AddRange(New String() {Oledatable.Rows(i).Item("Societe")})
                End If
            Next i
        Catch ex As Exception

        End Try
    End Sub
    Private Sub AfficheSchemasIntegrer()
        Dim i As Integer
        Dim OleAdaptaterschema As OleDbDataAdapter
        Dim OleSchemaDataset As DataSet
        Dim OledatableSchema As DataTable
        DataListeIntegrer.Rows.Clear()
        Try
            OleAdaptaterschema = New OleDbDataAdapter("select * from WET_SCHEMA", OleConnenection)
            OleSchemaDataset = New DataSet
            OleAdaptaterschema.Fill(OleSchemaDataset)
            OledatableSchema = OleSchemaDataset.Tables(0)
            DataListeIntegrer.RowCount = OledatableSchema.Rows.Count
            For i = 0 To OledatableSchema.Rows.Count - 1
                DataListeIntegrer.Rows(i).Cells("BaseCpta1").Value = OledatableSchema.Rows(i).Item("BaseCpta")
                DataListeIntegrer.Rows(i).Cells("BaseCial1").Value = OledatableSchema.Rows(i).Item("BaseCial")
                DataListeIntegrer.Rows(i).Cells("CheminForma").Value = OledatableSchema.Rows(i).Item("CheminFormat")
                DataListeIntegrer.Rows(i).Cells("FeuilleExcel1").Value = OledatableSchema.Rows(i).Item("Feuille_Excel")
                DataListeIntegrer.Rows(i).Cells("NameFormat").Value = OledatableSchema.Rows(i).Item("NomFormat")
                DataListeIntegrer.Rows(i).Cells("TypeFormat1").Value = OledatableSchema.Rows(i).Item("Type")
                DataListeIntegrer.Rows(i).Cells("Flag1").Value = OledatableSchema.Rows(i).Item("Chmpflag")
                DataListeIntegrer.Rows(i).Cells("Valeur1").Value = OledatableSchema.Rows(i).Item("Valflag")
                DataListeIntegrer.Rows(i).Cells("RepExtraction1").Value = OledatableSchema.Rows(i).Item("CheminExport")
                DataListeIntegrer.Rows(i).Cells("EstEntete1").Value = OledatableSchema.Rows(i).Item("EstEntete")
                DataListeIntegrer.Rows(i).Cells("IDDossier1").Value = OledatableSchema.Rows(i).Item("IDDossier")
                DataListeIntegrer.Rows(i).Cells("Cible1").Value = OledatableSchema.Rows(i).Item("Cible")
            Next i
        Catch ex As Exception

        End Try
    End Sub
    Private Sub Delete_DataListeSch()
        Dim i As Integer
        Dim OleAdaptaterDelete As OleDbDataAdapter
        Dim OleDeleteDataset As DataSet
        Dim OledatableDelete As DataTable
        Dim OleCommandDelete As OleDbCommand
        Dim DelFile As String
        Try
            For i = 0 To DataListeIntegrer.RowCount - 1
                If DataListeIntegrer.Rows(i).Cells("Supprimer").Value = True Then
                    OleAdaptaterDelete = New OleDbDataAdapter("select * from WET_SCHEMA where Cible='" & DataListeIntegrer.Rows(i).Cells("Cible1").Value & "' And  CheminFormat='" & DataListeIntegrer.Rows(i).Cells("CheminForma").Value & "' and BaseCpta='" & DataListeIntegrer.Rows(i).Cells("BaseCpta1").Value & "'  and Type='" & DataListeIntegrer.Rows(i).Cells("TypeFormat1").Value & "'", OleConnenection)
                    OleDeleteDataset = New DataSet
                    OleAdaptaterDelete.Fill(OleDeleteDataset)
                    OledatableDelete = OleDeleteDataset.Tables(0)
                    If OledatableDelete.Rows.Count <> 0 Then
                        DelFile = "Delete From WET_SCHEMA where Cible='" & DataListeIntegrer.Rows(i).Cells("Cible1").Value & "' And  CheminFormat='" & DataListeIntegrer.Rows(i).Cells("CheminForma").Value & "' and BaseCpta='" & DataListeIntegrer.Rows(i).Cells("BaseCpta1").Value & "' and Type='" & DataListeIntegrer.Rows(i).Cells("TypeFormat1").Value & "'"
                        OleCommandDelete = New OleDbCommand(DelFile)
                        OleCommandDelete.Connection = OleConnenection
                        OleCommandDelete.ExecuteNonQuery()
                        DelFile = "Delete From PLANIFICATION where (Intitule='Export Article') and IDDossier=" & CInt(DataListeIntegrer.Rows(i).Cells("IDDossier1").Value) & ""
                        OleCommandDelete = New OleDbCommand(DelFile)
                        OleCommandDelete.Connection = OleConnenection
                        OleCommandDelete.ExecuteNonQuery()
                    End If
                End If
            Next i
            AfficheSchemasIntegrer()
        Catch ex As Exception

        End Try
    End Sub
    Public Sub TranscodageIDExiste(ByRef IDtraitement As Integer, ByRef Categorietraitement As String)
        Dim OleAdaptater As OleDbDataAdapter
        Dim OleAfficheDataset As DataSet
        Dim Oledatable As DataTable
        Try
            OleAdaptater = New OleDbDataAdapter("select * from TRANSCODAGEEXPORT Where IDDossier=" & IDtraitement & " And Categorie='" & Categorietraitement & "'", OleConnenection)
            OleAfficheDataset = New DataSet
            OleAdaptater.Fill(OleAfficheDataset)
            Oledatable = OleAfficheDataset.Tables(0)
            If Oledatable.Rows.Count <> 0 Then
                MsgBox("          Transcodage Existant " & Chr(13) & "Catégorie de traitement :" & Categorietraitement & ", ID : " & IDtraitement & Chr(13) & "" & Chr(13) & " supprimez les transcodages si inutilisés !", MsgBoxStyle.Information, "Paramétrage de traitement")
            End If
        Catch ex As Exception
            MsgBox("Erreur système :" & ex.Message)
        End Try
    End Sub
    Private Sub EnregistrerLeSchema()
        Dim n As Integer
        Dim OleAdaptaterEnreg As OleDbDataAdapter
        Dim OleEnregDataset As DataSet
        Dim OledatableEnreg As DataTable
        Dim OleCommandEnreg As OleDbCommand
        Dim Insert As Boolean = False
        Dim Insertion As String
        Try
            If DataListeSchema.RowCount >= 1 Then
                For n = 0 To DataListeSchema.RowCount - 1
                    If DataListeSchema.Rows(n).Cells("Cible").Value = "Repertoire" Then
                        If IsNumeric(DataListeSchema.Rows(n).Cells("IDDossier").Value) = True Then
                            OleAdaptaterEnreg = New OleDbDataAdapter("select * From WET_SCHEMA WHERE Cible='" & DataListeSchema.Rows(n).Cells("Cible").Value & "' and CheminFormat='" & DataListeSchema.Rows(n).Cells("Chemin").Value & "' and BaseCpta='" & DataListeSchema.Rows(n).Cells("BaseCpta").Value & "'  and Type='" & DataListeSchema.Rows(n).Cells("TypeFormat").Value & "'", OleConnenection)
                            OleEnregDataset = New DataSet
                            OleAdaptaterEnreg.Fill(OleEnregDataset)
                            OledatableEnreg = OleEnregDataset.Tables(0)
                            If OledatableEnreg.Rows.Count <> 0 Then
                            Else
                                OleAdaptaterEnreg = New OleDbDataAdapter("select * From WET_SCHEMA WHERE IDDossier=" & CInt(DataListeSchema.Rows(n).Cells("IDDossier").Value) & "", OleConnenection)
                                OleEnregDataset = New DataSet
                                OleAdaptaterEnreg.Fill(OleEnregDataset)
                                OledatableEnreg = OleEnregDataset.Tables(0)
                                If OledatableEnreg.Rows.Count <> 0 Then

                                Else
                                    If Trim(DataListeSchema.Rows(n).Cells("Chemin").Value) <> "" And Trim(DataListeSchema.Rows(n).Cells("BaseCpta").Value) <> "" And DataListeSchema.Rows(n).Cells("TypeFormat").Value <> "" Then
                                        Insertion = "Insert Into WET_SCHEMA (BaseCpta,Cible,BaseCial,CheminFormat,NomFormat,Type,Feuille_Excel,Chmpflag,Valflag,CheminExport,EstEntete,IDDossier) VALUES ('" & DataListeSchema.Rows(n).Cells("BaseCpta").Value & "','" & DataListeSchema.Rows(n).Cells("Cible").Value & "','" & DataListeSchema.Rows(n).Cells("BaseCial").Value & "','" & DataListeSchema.Rows(n).Cells("Chemin").Value & "','" & DataListeSchema.Rows(n).Cells("NomFormat").Value & "','" & DataListeSchema.Rows(n).Cells("TypeFormat").Value & "','" & DataListeSchema.Rows(n).Cells("FeuilleExcel").Value & "','" & DataListeSchema.Rows(n).Cells("Flag").Value & "','" & DataListeSchema.Rows(n).Cells("Valeur").Value & "','" & DataListeSchema.Rows(n).Cells("RepExtraction").Value & "'," & DataListeSchema.Rows(n).Cells("EstEntete").Value & "," & CInt(DataListeSchema.Rows(n).Cells("IDDossier").Value) & ")"
                                        OleCommandEnreg = New OleDbCommand(Insertion)
                                        OleCommandEnreg.Connection = OleConnenection
                                        OleCommandEnreg.ExecuteNonQuery()
                                        TranscodageIDExiste(CInt(DataListeSchema.Rows(n).Cells("IDDossier").Value), "Tiers")
                                        Insert = True
                                    End If
                                End If
                            End If
                        Else
                            MsgBox("L'ID : " & DataListeSchema.Rows(n).Cells("IDDossier").Value & " du dossier doit être un entier !", MsgBoxStyle.Information, "Insertion des Schemas d'integration")
                        End If
                    ElseIf DataListeSchema.Rows(n).Cells("Cible").Value = "FTP" Then
                        If IsNumeric(DataListeSchema.Rows(n).Cells("IDDossier").Value) = True Then
                            OleAdaptaterEnreg = New OleDbDataAdapter("select * From WET_SCHEMA WHERE Cible='" & DataListeSchema.Rows(n).Cells("Cible").Value & "' and CheminFormat='" & DataListeSchema.Rows(n).Cells("Chemin").Value & "' and BaseCpta='" & DataListeSchema.Rows(n).Cells("BaseCpta").Value & "'  and Type='" & DataListeSchema.Rows(n).Cells("TypeFormat").Value & "'", OleConnenection)
                            OleEnregDataset = New DataSet
                            OleAdaptaterEnreg.Fill(OleEnregDataset)
                            OledatableEnreg = OleEnregDataset.Tables(0)
                            If OledatableEnreg.Rows.Count <> 0 Then
                            Else
                                OleAdaptaterEnreg = New OleDbDataAdapter("select * From WET_SCHEMA WHERE IDDossier=" & CInt(DataListeSchema.Rows(n).Cells("IDDossier").Value) & "", OleConnenection)
                                OleEnregDataset = New DataSet
                                OleAdaptaterEnreg.Fill(OleEnregDataset)
                                OledatableEnreg = OleEnregDataset.Tables(0)
                                If OledatableEnreg.Rows.Count <> 0 Then

                                Else
                                    If Trim(DataListeSchema.Rows(n).Cells("Chemin").Value) <> "" And Trim(DataListeSchema.Rows(n).Cells("BaseCpta").Value) <> "" And DataListeSchema.Rows(n).Cells("TypeFormat").Value <> "" Then
                                        If RetourneDirectoryFtp(Trim(DataListeSchema.Rows(n).Cells("RepExtraction").Value)) <> "" And RetournePassWordFtp(Trim(DataListeSchema.Rows(n).Cells("RepExtraction").Value)) <> "" And RetourneServeurFtp(Trim(DataListeSchema.Rows(n).Cells("RepExtraction").Value)) <> "" And RetourneUserFtp(Trim(DataListeSchema.Rows(n).Cells("RepExtraction").Value)) <> "" Then
                                            Insertion = "Insert Into WET_SCHEMA (BaseCpta,Cible,BaseCial,CheminFormat,NomFormat,Type,Feuille_Excel,Chmpflag,Valflag,CheminExport,EstEntete,IDDossier) VALUES ('" & DataListeSchema.Rows(n).Cells("BaseCpta").Value & "','" & DataListeSchema.Rows(n).Cells("Cible").Value & "','" & DataListeSchema.Rows(n).Cells("BaseCial").Value & "','" & DataListeSchema.Rows(n).Cells("Chemin").Value & "','" & DataListeSchema.Rows(n).Cells("NomFormat").Value & "','" & DataListeSchema.Rows(n).Cells("TypeFormat").Value & "','" & DataListeSchema.Rows(n).Cells("FeuilleExcel").Value & "','" & DataListeSchema.Rows(n).Cells("Flag").Value & "','" & DataListeSchema.Rows(n).Cells("Valeur").Value & "','" & DataListeSchema.Rows(n).Cells("RepExtraction").Value & "'," & DataListeSchema.Rows(n).Cells("EstEntete").Value & "," & CInt(DataListeSchema.Rows(n).Cells("IDDossier").Value) & ")"
                                            OleCommandEnreg = New OleDbCommand(Insertion)
                                            OleCommandEnreg.Connection = OleConnenection
                                            OleCommandEnreg.ExecuteNonQuery()
                                            TranscodageIDExiste(CInt(DataListeSchema.Rows(n).Cells("IDDossier").Value), "Tiers")
                                            Insert = True
                                        Else
                                            MsgBox("Impossible d'extraire un paramètre Ftp : " & Trim(DataListeSchema.Rows(n).Cells("RepExtraction").Value) & " Respectez le format de Saisie !", MsgBoxStyle.Information, "Insertion des Schemas d'integration")
                                        End If
                                    End If
                                End If
                            End If
                        Else
                            MsgBox("L'ID : " & DataListeSchema.Rows(n).Cells("IDDossier").Value & " du dossier doit être un entier !", MsgBoxStyle.Information, "Insertion des Schemas d'integration")
                        End If
                    End If
                Next n
                If Insert = True Then
                    MsgBox("Insertion Reussie", MsgBoxStyle.Information, "Insertion des Schemas d'integration")
                    DataListeIntegrer.Rows.Clear()
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub BT_Quit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Quit.Click
        Me.Close()
    End Sub
    Private Sub BT_Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Save.Click
        Try
            EnregistrerLeSchema()
            AfficheSchemasIntegrer()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub BT_DelRow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_DelRow.Click
        Dim first As Integer
        Dim last As Integer
        Try
            first = DataListeSchema.Rows.GetFirstRow(DataGridViewElementStates.Displayed)
            last = DataListeSchema.Rows.GetLastRow(DataGridViewElementStates.Displayed)
            If last >= 0 Then
                If last - first >= 0 Then
                    DataListeSchema.Rows.RemoveAt(DataListeSchema.CurrentRow.Index)
                End If
            End If
        Catch ex As Exception

        End Try
        For i As Integer = 0 To DataListeSchema.RowCount - 1
            DataListeSchema.Rows(i).Cells("IDDossier").Value = RenvoieID("WET_SCHEMA") + i
        Next i
    End Sub
    Private Sub BT_ADD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_ADD.Click
        Dim i As Integer = DataListeSchema.Rows.Add()
        DataListeSchema.Rows(i).Cells("EstEntete").Value = False
        DataListeSchema.Rows(i).Cells("IDDossier").Value = RenvoieID("WET_SCHEMA") + i
        If RenvoieID("WET_SCHEMA") = DataListeSchema.Rows(0).Cells("IDDossier").Value Then

        Else
            For j As Integer = 0 To DataListeSchema.RowCount - 1
                DataListeSchema.Rows(j).Cells("IDDossier").Value = RenvoieID("WET_SCHEMA") + j
            Next j
        End If
    End Sub
    Private Sub BT_Delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Delete.Click
        Delete_DataListeSch()
    End Sub

    Private Sub DataListeSchema_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataListeSchema.CellClick
        Idexe = 0
        Try
            If e.RowIndex >= 0 Then
                If DataListeSchema.Columns(e.ColumnIndex).Name = "Type" Then
                    Idexe = e.RowIndex
                    SelectionFormatTiers.Name = "EXTRACTION"
                    SelectionFormatTiers.ShowDialog()
                End If
                If DataListeSchema.Columns(e.ColumnIndex).Name <> "Cible" Then
                    If DataListeSchema.Rows(e.RowIndex).Cells("Cible").Value = "" Then
                        MsgBox("Veuillez Choisir une cible Svp", MsgBoxStyle.Information, "Choix d'une cible")
                        Exit Sub
                    End If
                End If
                If DataListeSchema.Columns(e.ColumnIndex).Name = "Flag" Or DataListeSchema.Columns(e.ColumnIndex).Name = "Flagligne" Then
                    If Trim(DataListeSchema.Rows(e.RowIndex).Cells("BaseCpta").Value) = "" Then
                        MsgBox("Veuillez choisir d'abord la Société Comptable !", MsgBoxStyle.Information, "Selection Champ Flag")
                        Exit Sub
                    End If
                End If
                If DataListeSchema.Columns(e.ColumnIndex).Name = "ChoixExtraction" Then
                    If DataListeSchema.Rows(e.RowIndex).Cells("Cible").Value = "Repertoire" Then
                        FolderRepListeFile.Description = "Repertoire d'export"
                        If FolderRepListeFile.ShowDialog = Windows.Forms.DialogResult.OK Then
                            DataListeSchema.Rows(e.RowIndex).Cells("RepExtraction").Value = FolderRepListeFile.SelectedPath & "\"
                        End If
                    Else
                        If DataListeSchema.Rows(e.RowIndex).Cells("Cible").Value = "FTP" Then
                            MsgBox(getPingTime(RetourneServeurFtping(Trim(DataListeSchema.Rows(e.RowIndex).Cells("RepExtraction").Value))), MsgBoxStyle.Information, "ping du serveur " & RetourneServeurFtping(Trim(DataListeSchema.Rows(e.RowIndex).Cells("RepExtraction").Value)))
                        End If
                    End If
                    
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub MiseàjourLeSchema()
        Dim n As Integer
        Dim OleAdaptaterEnreg As OleDbDataAdapter
        Dim OleEnregDataset As DataSet
        Dim OledatableEnreg As DataTable
        Dim OleCommandEnreg As OleDbCommand
        Dim Insert As Boolean = False
        Dim Insertion As String
        If DataListeIntegrer.RowCount >= 0 Then
            For n = 0 To DataListeIntegrer.RowCount - 1
                OleAdaptaterEnreg = New OleDbDataAdapter("select * From WET_SCHEMA where  Cible='" & DataListeIntegrer.Rows(n).Cells("Cible1").Value & "' And  CheminFormat='" & DataListeIntegrer.Rows(n).Cells("CheminForma").Value & "' and BaseCpta='" & DataListeIntegrer.Rows(n).Cells("BaseCpta1").Value & "'  and Type='" & DataListeIntegrer.Rows(n).Cells("TypeFormat1").Value & "'", OleConnenection)
                OleEnregDataset = New DataSet
                OleAdaptaterEnreg.Fill(OleEnregDataset)
                OledatableEnreg = OleEnregDataset.Tables(0)
                If OledatableEnreg.Rows.Count <> 0 Then
                    Insertion = "UPDATE WET_SCHEMA SET Feuille_Excel='" & DataListeIntegrer.Rows(n).Cells("FeuilleExcel1").Value & "',Valflag='" & DataListeIntegrer.Rows(n).Cells("Valeur1").Value & "',CheminExport='" & DataListeIntegrer.Rows(n).Cells("RepExtraction1").Value & "',EstEntete=" & DataListeIntegrer.Rows(n).Cells("EstEntete1").Value & " where  Cible='" & DataListeIntegrer.Rows(n).Cells("Cible1").Value & "' And  CheminFormat='" & DataListeIntegrer.Rows(n).Cells("CheminForma").Value & "' and BaseCpta='" & DataListeIntegrer.Rows(n).Cells("BaseCpta1").Value & "'  and Type='" & DataListeIntegrer.Rows(n).Cells("TypeFormat1").Value & "'"
                    OleCommandEnreg = New OleDbCommand(Insertion)
                    OleCommandEnreg.Connection = OleConnenection
                    OleCommandEnreg.ExecuteNonQuery()
                    Insert = True
                End If
            Next n
            If Insert = True Then
                MsgBox("Mise à Jour Reussie", MsgBoxStyle.Information, "Mise à Jour des Schemas d'extraction")
            End If
        End If
        AfficheSchemasIntegrer()
    End Sub

    Private Sub BT_Update_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Update.Click
        MiseàjourLeSchema()
    End Sub

    Private Sub DataListeSchema_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataListeSchema.CellContentClick
        If e.RowIndex >= 0 Then
            If DataListeSchema.Columns(e.ColumnIndex).Name = "Flag" Then
                If Trim(DataListeSchema.Rows(e.RowIndex).Cells("BaseCpta").Value) = "" Then
                    MsgBox("Veuillez choisir d'abord la Société Comptable !", MsgBoxStyle.Information, "Selection Champ Flag")
                    Exit Sub
                End If
            End If
        End If
    End Sub
    Private Sub DataListeSchema_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataListeSchema.CellEndEdit
        If e.RowIndex >= 0 Then
            Flag.Items.Clear()
            Dim i As Integer
            Dim OleAdaptaterEnreg As OleDbDataAdapter
            Dim OleEnregDataset As DataSet
            Dim OledatableEnreg As DataTable
            OleAdaptaterEnreg = New OleDbDataAdapter("select * From PARAMETRE where  Societe='" & Trim(DataListeSchema.Rows(e.RowIndex).Cells("BaseCpta").Value) & "'", OleConnenection)
            OleEnregDataset = New DataSet
            OleAdaptaterEnreg.Fill(OleEnregDataset)
            OledatableEnreg = OleEnregDataset.Tables(0)
            If OledatableEnreg.Rows.Count <> 0 Then
                If SociéteConnection(OledatableEnreg.Rows(0).Item("Serveur"), OledatableEnreg.Rows(0).Item("BaseDonnee"), OledatableEnreg.Rows(0).Item("NomUser"), OledatableEnreg.Rows(0).Item("MotPas")) = True Then
                    OleAdaptaterEnreg = New OleDbDataAdapter("select * From cbSysLibre where  CB_File='F_COMPTET'", FlagConnection)
                    OleEnregDataset = New DataSet
                    OleAdaptaterEnreg.Fill(OleEnregDataset)
                    OledatableEnreg = OleEnregDataset.Tables(0)
                    If OledatableEnreg.Rows.Count <> 0 Then
                        For i = 0 To OledatableEnreg.Rows.Count - 1
                            Flag.Items.Add(OledatableEnreg.Rows(i).Item("CB_Name"))
                        Next i
                    Else
                    End If
                Else
                    MsgBox("Echec de connexion SQL à la Société Comptable", MsgBoxStyle.Information, "Recupération Champ Flag")
                End If
            End If
        End If
    End Sub
    Private Sub DataListeSchema_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataListeSchema.DataError
        e.Cancel = True
    End Sub
    Private Sub DataListeIntegrer_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataListeIntegrer.CellClick
        If e.RowIndex >= 0 Then
            If DataListeIntegrer.Columns(e.ColumnIndex).Name = "ChoixExtraction1" Then
                If DataListeIntegrer.Rows(e.RowIndex).Cells("Cible1").Value = "Repertoire" Then
                    FolderRepListeFile.Description = "Repertoire d'export"
                    If FolderRepListeFile.ShowDialog = Windows.Forms.DialogResult.OK Then
                        DataListeIntegrer.Rows(e.RowIndex).Cells("RepExtraction1").Value = FolderRepListeFile.SelectedPath & "\"
                    End If
                ElseIf DataListeIntegrer.Rows(e.RowIndex).Cells("Cible1").Value = "FTP" Then
                    MsgBox(getPingTime(RetourneServeurFtping(Trim(DataListeIntegrer.Rows(e.RowIndex).Cells("RepExtraction1").Value))), MsgBoxStyle.Information, "ping du serveur " & RetourneServeurFtping(Trim(DataListeIntegrer.Rows(e.RowIndex).Cells("RepExtraction1").Value)))
                End If
            End If
        End If
    End Sub
    Private Sub DataListeIntegrer_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataListeIntegrer.DataError
        e.Cancel = True
    End Sub
    Private Function getPingTime(ByVal adresseIP As String) As String
        Dim monPing As New Ping
        Dim maReponsePing As PingReply
        Dim resultatPing As String = Nothing
        Try
            maReponsePing = monPing.Send(adresseIP, Nothing)
            resultatPing = "Réponse de " & adresseIP & " en " & maReponsePing.RoundtripTime.ToString & " ms."
            Return resultatPing
        Catch ex As Exception
            resultatPing = "Impossible de joindre l'hôte : " & ex.Message
            Return resultatPing
        End Try
    End Function    
End Class