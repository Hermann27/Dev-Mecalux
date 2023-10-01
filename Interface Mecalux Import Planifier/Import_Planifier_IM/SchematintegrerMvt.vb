Imports System.Data.OleDb
Imports System.IO
Imports System.Net.NetworkInformation
Public Class SchematintegrerMvt
    Private Sub SchematintegrerMvt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
    Private Sub AfficheSchemasIntegrer()
        Dim i As Integer
        Dim OleAdaptaterschema As OleDbDataAdapter
        Dim OleSchemaDataset As DataSet
        Dim OledatableSchema As DataTable
        DataListeIntegrer.Rows.Clear()
        Try
            OleAdaptaterschema = New OleDbDataAdapter("select * from SCHEMAS_IMPMOUV", OleConnenection)
            OleSchemaDataset = New DataSet
            OleAdaptaterschema.Fill(OleSchemaDataset)
            OledatableSchema = OleSchemaDataset.Tables(0)
            DataListeIntegrer.RowCount = OledatableSchema.Rows.Count
            For i = 0 To OledatableSchema.Rows.Count - 1
                DataListeIntegrer.Rows(i).Cells("NameFormat").Value = OledatableSchema.Rows(i).Item("NomFormat")
                DataListeIntegrer.Rows(i).Cells("NomRepxpor").Value = OledatableSchema.Rows(i).Item("NomFilexport")
                DataListeIntegrer.Rows(i).Cells("TypeFormat1").Value = Afficheauuser(OledatableSchema.Rows(i).Item("Type"))
                DataListeIntegrer.Rows(i).Cells("CheminForma").Value = OledatableSchema.Rows(i).Item("CheminFormat")
                DataListeIntegrer.Rows(i).Cells("CheminRepexpor").Value = OledatableSchema.Rows(i).Item("CheminFilexport")
                DataListeIntegrer.Rows(i).Cells("BaseCial1").Value = OledatableSchema.Rows(i).Item("BaseCial")
                DataListeIntegrer.Rows(i).Cells("BaseCpta1").Value = OledatableSchema.Rows(i).Item("BaseCpta")
                DataListeIntegrer.Rows(i).Cells("Deplace1").Value = OledatableSchema.Rows(i).Item("Deplace")
                DataListeIntegrer.Rows(i).Cells("Mode1").Value = OledatableSchema.Rows(i).Item("Mode")
                DataListeIntegrer.Rows(i).Cells("Feuille_Excel1").Value = OledatableSchema.Rows(i).Item("Feuille_Excel")
                DataListeIntegrer.Rows(i).Cells("Nom1").Value = OledatableSchema.Rows(i).Item("TriNom")
                DataListeIntegrer.Rows(i).Cells("Cr�ation1").Value = OledatableSchema.Rows(i).Item("TriCreation")
                DataListeIntegrer.Rows(i).Cells("Modification1").Value = OledatableSchema.Rows(i).Item("TriModification")
                DataListeIntegrer.Rows(i).Cells("IDDossier1").Value = OledatableSchema.Rows(i).Item("IDDossier")
                DataListeIntegrer.Rows(i).Cells("Categorie1").Value = OledatableSchema.Rows(i).Item("Categorie")
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
                    OleAdaptaterDelete = New OleDbDataAdapter("select * from SCHEMAS_IMPMOUV where IDDossier=" & CInt(DataListeIntegrer.Rows(i).Cells("IDDossier1").Value) & "", OleConnenection)
                    OleDeleteDataset = New DataSet
                    OleAdaptaterDelete.Fill(OleDeleteDataset)
                    OledatableDelete = OleDeleteDataset.Tables(0)
                    If OledatableDelete.Rows.Count <> 0 Then
                        DelFile = "Delete From SCHEMAS_IMPMOUV where IDDossier=" & CInt(DataListeIntegrer.Rows(i).Cells("IDDossier1").Value) & ""
                        OleCommandDelete = New OleDbCommand(DelFile)
                        OleCommandDelete.Connection = OleConnenection
                        OleCommandDelete.ExecuteNonQuery()
                        DelFile = "Delete From PLANIFICATION where (Intitule='Import Document Vente' Or Intitule='Import Document Achat') and IDDossier=" & CInt(DataListeIntegrer.Rows(i).Cells("IDDossier1").Value) & ""
                        OleCommandDelete = New OleDbCommand(DelFile)
                        OleCommandDelete.Connection = OleConnenection
                        OleCommandDelete.ExecuteNonQuery()
                        Dim OleAdaptaterEnreg As OleDbDataAdapter
                        Dim OleEnregDataset As DataSet
                        Dim OledatableEnreg As DataTable
                        OleAdaptaterEnreg = New OleDbDataAdapter("select * From FTPREPERTOIRE WHERE IDDossier=" & CInt(DataListeIntegrer.Rows(i).Cells("IDDossier1").Value) & " And  Traitement='IMPORT DOCUMENT'", OleConnenection)
                        OleEnregDataset = New DataSet
                        OleAdaptaterEnreg.Fill(OleEnregDataset)
                        OledatableEnreg = OleEnregDataset.Tables(0)
                        If OledatableEnreg.Rows.Count <> 0 Then
                            If Directory.Exists(OledatableEnreg.Rows(0).Item("Repertoire")) = True Then
                                DelFile = "Delete From FTPREPERTOIRE where IDDossier=" & CInt(DataListeIntegrer.Rows(i).Cells("IDDossier1").Value) & " And  Traitement='IMPORT DOCUMENT'"
                                OleCommandDelete = New OleDbCommand(DelFile)
                                OleCommandDelete.Connection = OleConnenection
                                OleCommandDelete.ExecuteNonQuery()
                                Directory.Delete(OledatableEnreg.Rows(0).Item("Repertoire"), True)
                            Else
                                DelFile = "Delete From FTPREPERTOIRE where IDDossier=" & CInt(DataListeIntegrer.Rows(i).Cells("IDDossier1").Value) & " And  Traitement='IMPORT DOCUMENT'"
                                OleCommandDelete = New OleDbCommand(DelFile)
                                OleCommandDelete.Connection = OleConnenection
                                OleCommandDelete.ExecuteNonQuery()
                            End If
                        End If
                    End If
                End If
            Next i
            AfficheSchemasIntegrer()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, "Param�trage de traitement")
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
                            If Trim(DataListeSchema.Rows(n).Cells("Categorie").Value) <> "" Then
                                OleAdaptaterEnreg = New OleDbDataAdapter("select * From SCHEMAS_IMPMOUV WHERE Cible='" & DataListeSchema.Rows(n).Cells("Cible").Value & "' and CheminFormat='" & DataListeSchema.Rows(n).Cells("Chemin").Value & "' and CheminFilexport='" & DataListeSchema.Rows(n).Cells("CheminExport").Value & "' and Type='" & Renvoietypeformat(DataListeSchema.Rows(n).Cells("TypeFormat").Value) & "'", OleConnenection)
                                OleEnregDataset = New DataSet
                                OleAdaptaterEnreg.Fill(OleEnregDataset)
                                OledatableEnreg = OleEnregDataset.Tables(0)
                                If OledatableEnreg.Rows.Count <> 0 Then
                                Else
                                    OleAdaptaterEnreg = New OleDbDataAdapter("select * From SCHEMAS_IMPMOUV WHERE IDDossier=" & CInt(DataListeSchema.Rows(n).Cells("IDDossier").Value) & " ", OleConnenection)
                                    OleEnregDataset = New DataSet
                                    OleAdaptaterEnreg.Fill(OleEnregDataset)
                                    OledatableEnreg = OleEnregDataset.Tables(0)
                                    If OledatableEnreg.Rows.Count <> 0 Then
                                    Else
                                        If Trim(DataListeSchema.Rows(n).Cells("Chemin").Value) <> "" And Trim(DataListeSchema.Rows(n).Cells("BaseCial").Value) <> "" And Trim(DataListeSchema.Rows(n).Cells("BaseCpta").Value) <> "" And Trim(DataListeSchema.Rows(n).Cells("CheminExport").Value) <> "" And DataListeSchema.Rows(n).Cells("TypeFormat").Value <> "" Then
                                            Insertion = "Insert Into SCHEMAS_IMPMOUV (Cible,BaseCial,BaseCpta,CheminFormat,CheminFilexport,NomFormat,NomFilexport,Type,Deplace,Mode,Feuille_Excel,TriNom,TriCreation,TriModification,IDDossier,Categorie) VALUES ('" & Trim(DataListeSchema.Rows(n).Cells("Cible").Value) & "','" & DataListeSchema.Rows(n).Cells("BaseCial").Value & "','" & DataListeSchema.Rows(n).Cells("BaseCpta").Value & "','" & DataListeSchema.Rows(n).Cells("Chemin").Value & "','" & DataListeSchema.Rows(n).Cells("CheminExport").Value & "','" & DataListeSchema.Rows(n).Cells("NomFormat").Value & "','" & DataListeSchema.Rows(n).Cells("DossierExport").Value & "','" & Renvoietypeformat(DataListeSchema.Rows(n).Cells("TypeFormat").Value) & "'," & DataListeSchema.Rows(n).Cells("Deplace").Value & ",'" & DataListeSchema.Rows(n).Cells("Mode").Value & "','" & DataListeSchema.Rows(n).Cells("Feuille_Excel").Value & "'," & DataListeSchema.Rows(n).Cells("Nom").Value & "," & DataListeSchema.Rows(n).Cells("Cr�ation").Value & "," & DataListeSchema.Rows(n).Cells("Modification").Value & "," & CInt(DataListeSchema.Rows(n).Cells("IDDossier").Value) & ",'" & Trim(DataListeSchema.Rows(n).Cells("Categorie").Value) & "')"
                                            OleCommandEnreg = New OleDbCommand(Insertion)
                                            OleCommandEnreg.Connection = OleConnenection
                                            OleCommandEnreg.ExecuteNonQuery()
                                            Insert = True
                                        End If
                                    End If
                                End If
                            Else
                                MsgBox("La Cat�gorie Vente/Achat : " & DataListeSchema.Rows(n).Cells("Categorie").Value & " n'est pas renseign�e !", MsgBoxStyle.Information, "Insertion des Schemas d'integration")
                            End If
                        Else
                            MsgBox("L'ID : " & DataListeSchema.Rows(n).Cells("IDDossier").Value & " du dossier doit �tre un entier !", MsgBoxStyle.Information, "Insertion des Schemas d'integration")
                        End If
                    Else
                        If DataListeSchema.Rows(n).Cells("Cible").Value = "FTP" Then
                            If IsNumeric(DataListeSchema.Rows(n).Cells("IDDossier").Value) = True Then
                                If Trim(DataListeSchema.Rows(n).Cells("Categorie").Value) <> "" Then
                                    OleAdaptaterEnreg = New OleDbDataAdapter("select * From SCHEMAS_IMPMOUV WHERE Cible='" & DataListeSchema.Rows(n).Cells("Cible").Value & "' and CheminFormat='" & DataListeSchema.Rows(n).Cells("Chemin").Value & "' and CheminFilexport='" & DataListeSchema.Rows(n).Cells("CheminExport").Value & "' and Type='" & Renvoietypeformat(DataListeSchema.Rows(n).Cells("TypeFormat").Value) & "'", OleConnenection)
                                    OleEnregDataset = New DataSet
                                    OleAdaptaterEnreg.Fill(OleEnregDataset)
                                    OledatableEnreg = OleEnregDataset.Tables(0)
                                    If OledatableEnreg.Rows.Count <> 0 Then
                                    Else
                                        OleAdaptaterEnreg = New OleDbDataAdapter("select * From SCHEMAS_IMPMOUV WHERE IDDossier=" & CInt(DataListeSchema.Rows(n).Cells("IDDossier").Value) & " ", OleConnenection)
                                        OleEnregDataset = New DataSet
                                        OleAdaptaterEnreg.Fill(OleEnregDataset)
                                        OledatableEnreg = OleEnregDataset.Tables(0)
                                        If OledatableEnreg.Rows.Count <> 0 Then
                                        Else
                                            If Trim(DataListeSchema.Rows(n).Cells("Chemin").Value) <> "" And Trim(DataListeSchema.Rows(n).Cells("BaseCial").Value) <> "" And Trim(DataListeSchema.Rows(n).Cells("BaseCpta").Value) <> "" And Trim(DataListeSchema.Rows(n).Cells("CheminExport").Value) <> "" And DataListeSchema.Rows(n).Cells("TypeFormat").Value <> "" Then
                                                If RetourneDirectoryFtp(Trim(DataListeSchema.Rows(n).Cells("CheminExport").Value)) <> "" And RetournePassWordFtp(Trim(DataListeSchema.Rows(n).Cells("CheminExport").Value)) <> "" And RetourneServeurFtp(Trim(DataListeSchema.Rows(n).Cells("CheminExport").Value)) <> "" And RetourneUserFtp(Trim(DataListeSchema.Rows(n).Cells("CheminExport").Value)) <> "" Then
                                                    Insertion = "Insert Into SCHEMAS_IMPMOUV (Cible,BaseCial,BaseCpta,CheminFormat,CheminFilexport,NomFormat,NomFilexport,Type,Deplace,Mode,Feuille_Excel,TriNom,TriCreation,TriModification,IDDossier,Categorie) VALUES ('" & Trim(DataListeSchema.Rows(n).Cells("Cible").Value) & "','" & DataListeSchema.Rows(n).Cells("BaseCial").Value & "','" & DataListeSchema.Rows(n).Cells("BaseCpta").Value & "','" & DataListeSchema.Rows(n).Cells("Chemin").Value & "','" & DataListeSchema.Rows(n).Cells("CheminExport").Value & "','" & DataListeSchema.Rows(n).Cells("NomFormat").Value & "','" & DataListeSchema.Rows(n).Cells("DossierExport").Value & "','" & Renvoietypeformat(DataListeSchema.Rows(n).Cells("TypeFormat").Value) & "'," & DataListeSchema.Rows(n).Cells("Deplace").Value & ",'" & DataListeSchema.Rows(n).Cells("Mode").Value & "','" & DataListeSchema.Rows(n).Cells("Feuille_Excel").Value & "'," & DataListeSchema.Rows(n).Cells("Nom").Value & "," & DataListeSchema.Rows(n).Cells("Cr�ation").Value & "," & DataListeSchema.Rows(n).Cells("Modification").Value & "," & CInt(DataListeSchema.Rows(n).Cells("IDDossier").Value) & ",'" & Trim(DataListeSchema.Rows(n).Cells("Categorie").Value) & "')"
                                                    OleCommandEnreg = New OleDbCommand(Insertion)
                                                    OleCommandEnreg.Connection = OleConnenection
                                                    OleCommandEnreg.ExecuteNonQuery()
                                                    Insert = True
                                                    If Directory.Exists(PatchImportftp & "DOCUMENT" & DataListeSchema.Rows(n).Cells("IDDossier").Value) = False Then
                                                        Directory.CreateDirectory(PatchImportftp & "DOCUMENT" & DataListeSchema.Rows(n).Cells("IDDossier").Value)
                                                        Insertion = "Insert Into FTPREPERTOIRE (IDDossier,Traitement,Repertoire) VALUES (" & CInt(DataListeSchema.Rows(n).Cells("IDDossier").Value) & ",'IMPORT DOCUMENT','" & PatchImportftp & "DOCUMENT" & DataListeSchema.Rows(n).Cells("IDDossier").Value & "')"
                                                        OleCommandEnreg = New OleDbCommand(Insertion)
                                                        OleCommandEnreg.Connection = OleConnenection
                                                        OleCommandEnreg.ExecuteNonQuery()
                                                    Else
                                                        Dim OleAdaptaterFtp As OleDbDataAdapter
                                                        Dim OleFtpDataset As DataSet
                                                        Dim OledatableFtp As DataTable
                                                        OleAdaptaterFtp = New OleDbDataAdapter("select * From FTPREPERTOIRE WHERE IDDossier=" & CInt(CInt(DataListeSchema.Rows(n).Cells("IDDossier").Value)) & " And  Traitement='IMPORT DOCUMENT'", OleConnenection)
                                                        OleFtpDataset = New DataSet
                                                        OleAdaptaterFtp.Fill(OleFtpDataset)
                                                        OledatableFtp = OleFtpDataset.Tables(0)
                                                        If OledatableFtp.Rows.Count <> 0 Then
                                                        Else
                                                            Insertion = "Insert Into FTPREPERTOIRE (IDDossier,Traitement,Repertoire) VALUES (" & CInt(DataListeSchema.Rows(n).Cells("IDDossier").Value) & ",'IMPORT DOCUMENT','" & PatchImportftp & "DOCUMENT" & DataListeSchema.Rows(n).Cells("IDDossier").Value & "')"
                                                            OleCommandEnreg = New OleDbCommand(Insertion)
                                                            OleCommandEnreg.Connection = OleConnenection
                                                            OleCommandEnreg.ExecuteNonQuery()
                                                        End If
                                                    End If
                                                Else
                                                    MsgBox("Impossible d'extraire un param�tre Ftp : " & Trim(DataListeSchema.Rows(n).Cells("CheminExport").Value) & " Respectez le format de Saisie !", MsgBoxStyle.Information, "Insertion des Schemas d'integration")
                                                End If
                                            End If
                                        End If
                                    End If
                                Else
                                    MsgBox("La Cat�gorie Vente/Achat : " & DataListeSchema.Rows(n).Cells("Categorie").Value & " n'est pas renseign�e !", MsgBoxStyle.Information, "Insertion des Schemas d'integration")
                                End If
                            Else
                                MsgBox("L'ID : " & DataListeSchema.Rows(n).Cells("IDDossier").Value & " du dossier doit �tre un entier !", MsgBoxStyle.Information, "Insertion des Schemas d'integration")
                            End If
                        Else
                            If DataListeSchema.Rows(n).Cells("Cible").Value = "BaseSQL" Then
                                If IsNumeric(DataListeSchema.Rows(n).Cells("IDDossier").Value) = True Then
                                    If Trim(DataListeSchema.Rows(n).Cells("Categorie").Value) <> "" Then
                                        OleAdaptaterEnreg = New OleDbDataAdapter("select * From SCHEMAS_IMPMOUV WHERE IDDossier=" & CInt(DataListeSchema.Rows(n).Cells("IDDossier").Value) & " ", OleConnenection)
                                        OleEnregDataset = New DataSet
                                        OleAdaptaterEnreg.Fill(OleEnregDataset)
                                        OledatableEnreg = OleEnregDataset.Tables(0)
                                        If OledatableEnreg.Rows.Count <> 0 Then
                                        Else
                                            If Trim(DataListeSchema.Rows(n).Cells("Chemin").Value) <> "" And Trim(DataListeSchema.Rows(n).Cells("BaseCial").Value) <> "" And Trim(DataListeSchema.Rows(n).Cells("BaseCpta").Value) <> "" And Trim(DataListeSchema.Rows(n).Cells("CheminExport").Value) <> "" And DataListeSchema.Rows(n).Cells("TypeFormat").Value <> "" Then
                                                If RetourneServeurSQL(Trim(DataListeSchema.Rows(n).Cells("CheminExport").Value)) <> "" And RetourneBaseSQL(Trim(DataListeSchema.Rows(n).Cells("CheminExport").Value)) <> "" And RetourneTableSQL(Trim(DataListeSchema.Rows(n).Cells("CheminExport").Value)) <> "" Then
                                                    Insertion = "Insert Into SCHEMAS_IMPMOUV (Cible,BaseCial,BaseCpta,CheminFormat,CheminFilexport,NomFormat,NomFilexport,Type,Deplace,Mode,Feuille_Excel,TriNom,TriCreation,TriModification,IDDossier,Categorie) VALUES ('" & Trim(DataListeSchema.Rows(n).Cells("Cible").Value) & "','" & DataListeSchema.Rows(n).Cells("BaseCial").Value & "','" & DataListeSchema.Rows(n).Cells("BaseCpta").Value & "','" & DataListeSchema.Rows(n).Cells("Chemin").Value & "','" & DataListeSchema.Rows(n).Cells("CheminExport").Value & "','" & DataListeSchema.Rows(n).Cells("NomFormat").Value & "','" & DataListeSchema.Rows(n).Cells("DossierExport").Value & "','" & Renvoietypeformat(DataListeSchema.Rows(n).Cells("TypeFormat").Value) & "'," & DataListeSchema.Rows(n).Cells("Deplace").Value & ",'" & DataListeSchema.Rows(n).Cells("Mode").Value & "','" & DataListeSchema.Rows(n).Cells("Feuille_Excel").Value & "'," & DataListeSchema.Rows(n).Cells("Nom").Value & "," & DataListeSchema.Rows(n).Cells("Cr�ation").Value & "," & DataListeSchema.Rows(n).Cells("Modification").Value & "," & CInt(DataListeSchema.Rows(n).Cells("IDDossier").Value) & ",'" & Trim(DataListeSchema.Rows(n).Cells("Categorie").Value) & "')"
                                                    OleCommandEnreg = New OleDbCommand(Insertion)
                                                    OleCommandEnreg.Connection = OleConnenection
                                                    OleCommandEnreg.ExecuteNonQuery()
                                                    Insert = True
                                                Else
                                                    MsgBox("Impossible d'extraire un param�tre SQL Server/Serveur inexistant : " & Trim(DataListeSchema.Rows(n).Cells("CheminExport").Value) & " Respectez le format de Saisie !", MsgBoxStyle.Information, "Insertion des Schemas d'integration")
                                                End If
                                            End If
                                        End If
                                    Else
                                        MsgBox("La Cat�gorie Vente/Achat : " & DataListeSchema.Rows(n).Cells("Categorie").Value & " n'est pas renseign�e !", MsgBoxStyle.Information, "Insertion des Schemas d'integration")
                                    End If
                                Else
                                    MsgBox("L'ID : " & DataListeSchema.Rows(n).Cells("IDDossier").Value & " du dossier doit �tre un entier !", MsgBoxStyle.Information, "Insertion des Schemas d'integration")
                                End If
                            End If
                        End If
                    End If
                    
                Next n
                If Insert = True Then
                    MsgBox("Insertion Reussie", MsgBoxStyle.Information, "Insertion des Schemas d'integration")
                    DataListeIntegrer.Rows.Clear()
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, "Param�trage de traitement")
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
            DataListeSchema.Rows(i).Cells("IDDossier").Value = RenvoieID("SCHEMAS_IMPMOUV") + i
        Next i
    End Sub
    Private Sub BT_ADD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_ADD.Click
        Dim i As Integer = DataListeSchema.Rows.Add()
        DataListeSchema.Rows(i).Cells("Deplace").Value = False
        DataListeSchema.Rows(i).Cells("Nom").Value = False
        DataListeSchema.Rows(i).Cells("Cr�ation").Value = False
        DataListeSchema.Rows(i).Cells("Modification").Value = False
        DataListeSchema.Rows(i).Cells("IDDossier").Value = RenvoieID("SCHEMAS_IMPMOUV") + i
        If RenvoieID("SCHEMAS_IMPMOUV") = DataListeSchema.Rows(0).Cells("IDDossier").Value Then
        Else
            For j As Integer = 0 To DataListeSchema.RowCount - 1
                DataListeSchema.Rows(j).Cells("IDDossier").Value = RenvoieID("SCHEMAS_IMPMOUV") + j
            Next j
        End If
    End Sub
    Private Sub BT_Delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Delete.Click
        Delete_DataListeSch()
    End Sub

    Private Sub DataListeSchema_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataListeSchema.CellClick
        Me.Cursor = Cursors.WaitCursor
        Idexe = 0
        Try
            If e.RowIndex >= 0 Then
                If DataListeSchema.Columns(e.ColumnIndex).Name = "Type" Then
                    Idexe = e.RowIndex
                    SelectionFormatMvt.ShowDialog()
                End If
                If DataListeSchema.Columns(e.ColumnIndex).Name = "RechercheFichier" Then
                    If DataListeSchema.Rows(e.RowIndex).Cells("Cible").Value = "Repertoire" Then
                        FolderRepListeFile.Description = "Repertoire des Fichiers � traiter"
                        If FolderRepListeFile.ShowDialog = Windows.Forms.DialogResult.OK Then
                            DataListeSchema.Rows(e.RowIndex).Cells("CheminExport").Value = FolderRepListeFile.SelectedPath & "\"
                            DataListeSchema.Rows(e.RowIndex).Cells("DossierExport").Value = System.IO.Path.GetFileName(Trim(FolderRepListeFile.SelectedPath))
                        End If
                    ElseIf DataListeSchema.Rows(e.RowIndex).Cells("Cible").Value = "FTP" Then
                        MsgBox(getPingTime(RetourneServeurFtping(Trim(DataListeSchema.Rows(e.RowIndex).Cells("CheminExport").Value))), MsgBoxStyle.Information, "ping du serveur " & RetourneServeurFtping(Trim(DataListeSchema.Rows(e.RowIndex).Cells("CheminExport").Value)))
                    ElseIf DataListeSchema.Rows(e.RowIndex).Cells("Cible").Value = "BaseSQL" Then
                        If BaseSQLConnexion(RetourneServeurSQL(Trim(DataListeSchema.Rows(e.RowIndex).Cells("CheminExport").Value)), RetourneBaseSQL(Trim(DataListeSchema.Rows(e.RowIndex).Cells("CheminExport").Value)), RetourneUserSQL(RetourneServeurSQL(Trim(DataListeSchema.Rows(e.RowIndex).Cells("CheminExport").Value))), RetournePasseSQL(RetourneServeurSQL(Trim(DataListeSchema.Rows(e.RowIndex).Cells("CheminExport").Value)))) = True Then
                            If ExisteTableSQL(RetourneTableSQL(Trim(DataListeSchema.Rows(e.RowIndex).Cells("CheminExport").Value))) = True Then
                                SelectionFiltre.Name = "SelectionFiltreMVT"
                                SelectionFiltre.Txttable.Text = RetourneTableSQL(Trim(DataListeSchema.Rows(e.RowIndex).Cells("CheminExport").Value))
                                Idexe = e.RowIndex
                                SelectionFiltre.Show()
                            Else
                                MsgBox("Connexion � la Base SQL : " & RetourneBaseSQL(Trim(DataListeSchema.Rows(e.RowIndex).Cells("CheminExport").Value)) & " R�ussie !" & Chr(13) & "" & Chr(13) & "La table SQL : " & RetourneTableSQL(Trim(DataListeSchema.Rows(e.RowIndex).Cells("CheminExport").Value)) & " n'existe pas dans la base!")
                            End If
                        Else
                            MsgBox("Echec de Connexion � la Base SQL : " & RetourneBaseSQL(Trim(DataListeSchema.Rows(e.RowIndex).Cells("CheminExport").Value)) & " V�rifiez les param�tres de Connexion !")
                        End If
                    End If
                End If

                If DataListeSchema.Columns(e.ColumnIndex).Name <> "Cible" Then
                    If DataListeSchema.Rows(e.RowIndex).Cells("Cible").Value = "" Then
                        MsgBox("Veuillez Choisir une cible Svp", MsgBoxStyle.Information, "Choix d'une cible")
                        Exit Sub
                    End If
                End If
            End If
        Catch ex As Exception

        End Try
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub BT_Update_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Update.Click
        Mise�jourLeSchema()
    End Sub
    Private Sub Mise�jourLeSchema()
        Dim n As Integer
        Dim OleAdaptaterEnreg As OleDbDataAdapter
        Dim OleEnregDataset As DataSet
        Dim OledatableEnreg As DataTable
        Dim OleCommandEnreg As OleDbCommand
        Dim Insert As Boolean = False
        Dim Insertion As String
        If DataListeIntegrer.RowCount >= 0 Then
            For n = 0 To DataListeIntegrer.RowCount - 1
                OleAdaptaterEnreg = New OleDbDataAdapter("select * From SCHEMAS_IMPMOUV where   IDDossier=" & CInt(DataListeIntegrer.Rows(n).Cells("IDDossier1").Value) & "", OleConnenection)
                OleEnregDataset = New DataSet
                OleAdaptaterEnreg.Fill(OleEnregDataset)
                OledatableEnreg = OleEnregDataset.Tables(0)
                If OledatableEnreg.Rows.Count <> 0 Then
                    Insertion = "UPDATE SCHEMAS_IMPMOUV SET Feuille_Excel='" & Join(Split(DataListeIntegrer.Rows(n).Cells("Feuille_Excel1").Value, "'"), "''") & "',Deplace=" & DataListeIntegrer.Rows(n).Cells("Deplace1").Value & ",TriNom=" & DataListeIntegrer.Rows(n).Cells("Nom1").Value & ",TriCreation=" & DataListeIntegrer.Rows(n).Cells("Cr�ation1").Value & ",TriModification=" & DataListeIntegrer.Rows(n).Cells("Modification1").Value & " where IDDossier=" & CInt(DataListeIntegrer.Rows(n).Cells("IDDossier1").Value) & ""
                    OleCommandEnreg = New OleDbCommand(Insertion)
                    OleCommandEnreg.Connection = OleConnenection
                    OleCommandEnreg.ExecuteNonQuery()
                    Insert = True
                Else
                End If
            Next n
            If Insert = True Then
                MsgBox("Mise � Jour Reussie", MsgBoxStyle.Information, "Mise � Jour des Schemas d'integration")
            End If
        End If
        AfficheSchemasIntegrer()
    End Sub
    Private Function getPingTime(ByVal adresseIP As String) As String
        Dim monPing As New Ping
        Dim maReponsePing As PingReply
        Dim resultatPing As String = Nothing
        Try
            maReponsePing = monPing.Send(adresseIP, Nothing)
            resultatPing = "R�ponse de " & adresseIP & " en " & maReponsePing.RoundtripTime.ToString & " ms."
            Return resultatPing
        Catch ex As Exception
            resultatPing = "Impossible de joindre l'h�te : " & ex.Message
            Return resultatPing
        End Try
    End Function
    Private Sub DataListeIntegrer_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataListeIntegrer.CellClick
        Me.Cursor = Cursors.WaitCursor
        Idexe = 0
        Try
            If e.RowIndex >= 0 Then
                If DataListeIntegrer.Columns(e.ColumnIndex).Name = "RechercheFichier1" Then
                    If DataListeIntegrer.Rows(e.RowIndex).Cells("Cible1").Value = "Repertoire" Then
                    ElseIf DataListeIntegrer.Rows(e.RowIndex).Cells("Cible1").Value = "FTP" Then
                        MsgBox(getPingTime(RetourneServeurFtping(Trim(DataListeIntegrer.Rows(e.RowIndex).Cells("CheminRepexpor").Value))), MsgBoxStyle.Information, "ping du serveur " & RetourneServeurFtping(Trim(DataListeIntegrer.Rows(e.RowIndex).Cells("CheminRepexpor").Value)))
                    ElseIf DataListeIntegrer.Rows(e.RowIndex).Cells("Cible1").Value = "BaseSQL" Then
                        If BaseSQLConnexion(RetourneServeurSQL(Trim(DataListeIntegrer.Rows(e.RowIndex).Cells("CheminRepexpor").Value)), RetourneBaseSQL(Trim(DataListeIntegrer.Rows(e.RowIndex).Cells("CheminRepexpor").Value)), RetourneUserSQL(RetourneServeurSQL(Trim(DataListeIntegrer.Rows(e.RowIndex).Cells("CheminRepexpor").Value))), RetournePasseSQL(RetourneServeurSQL(Trim(DataListeIntegrer.Rows(e.RowIndex).Cells("CheminRepexpor").Value)))) = True Then
                            If ExisteTableSQL(RetourneTableSQL(Trim(DataListeIntegrer.Rows(e.RowIndex).Cells("CheminRepexpor").Value))) = True Then
                                SelectionFiltre.Name = "SelectionFiltreMVT1"
                                SelectionFiltre.Txttable.Text = RetourneTableSQL(Trim(DataListeIntegrer.Rows(e.RowIndex).Cells("CheminRepexpor").Value))
                                Idexe = e.RowIndex
                                SelectionFiltre.Show()
                            Else
                                MsgBox("Connexion � la Base SQL : " & RetourneBaseSQL(Trim(DataListeIntegrer.Rows(e.RowIndex).Cells("CheminRepexpor").Value)) & " R�ussie !" & Chr(13) & "" & Chr(13) & "La table SQL : " & RetourneTableSQL(Trim(DataListeIntegrer.Rows(e.RowIndex).Cells("CheminRepexpor").Value)) & " n'existe pas dans la base!")
                            End If
                        Else
                            MsgBox("Echec de Connexion � la Base SQL : " & RetourneBaseSQL(Trim(DataListeIntegrer.Rows(e.RowIndex).Cells("CheminRepexpor").Value)) & " V�rifiez les param�tres de Connexion !")
                        End If
                    End If
                End If
            End If
        Catch ex As Exception

        End Try
        Me.Cursor = Cursors.Default
    End Sub
End Class