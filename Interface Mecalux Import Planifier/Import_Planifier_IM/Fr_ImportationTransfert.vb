Imports Objets100Lib
Imports System
Imports System.Data.OleDb
Imports System.Collections
Imports System.Windows.Forms
Imports System.IO
Imports System.Xml
Imports Microsoft.VisualBasic
Public Class Fr_ImportationTransfert
    Public ProgresMax, IndexPrec, numLigne, numColonne, NbreLigne, NbreTotal, iRow As Integer
    Public Result, sColumnsSepar, DecFormat, IDDepotLigne, LigneIntituleDepot As Object
    Public EntetePieceInterne, EntetePiecePrecedent, EnteteReference, EntetePlanAnalytique As Object
    Public EnteteSoucheDocument, EnteteTyPeDocument, LigneDatedeFabrication, LigneDatedeLivraison, LigneDatedePeremption As Object
    Public LigneDesignationArticle, LigneNSerieLot, LigneCodeArticle, PieceArticle, EnteteDateDocument As Object
    Public LignePoidsBrut, LignePoidsNet, LignePrixUnitaire, LigneQuantite, LigneReference, EnteteCodeAffaire As Object
    Public IDDepotEnteteOrigine, EnteteIntituleDepotOrigine, EnteteIntituleDepotDestination, IDDepotEnteteDestination As Object
    Public PuStock As Double
    Public OleSocieteConnect As OleDbConnection
    'Variable d'exception du deplacement de fichier
    Public exceptionTrouve As Boolean = False
    Public ExisteLecture As Boolean = True
    Public Filebool As Boolean
    Public NomFichier As String
    Public infoListe As List(Of Integer)
    Public infoLigne As List(Of Integer)
    Public ListePiece As List(Of String)
    Public ListeStock As List(Of String)
    Public Document As IBODocumentStock3 = Nothing
    Public LigneDocument As IBODocumentStockLigne3 = Nothing
    Public DocumentInfolibre As IBODocumentStock3 = Nothing
    Public PlanAna As IBPAnalytique3
    Public TraitementID As Integer
    Public Sub Fr_ImportationTransfert_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Connected() = True Then
                ListBox1.Items.Clear()
                AfficheSchemasIntegrer()
                Affichagefichier()
                Initialiser()
                Datagridaffiche.Rows.Clear()
                Datagridaffiche.Columns.Clear()
                Me.WindowState = FormWindowState.Maximized
            End If
            encours.Close()
        Catch ex As Exception
            encours.Close()
        End Try
    End Sub
    Public Function SocieteConnected(ByRef BaseConsolide As String, ByRef Mot_Psql As String, ByRef Nom_Utsql As String, ByRef Serveur As String) As Boolean
        Try
            OleSocieteConnect = New OleDbConnection("provider=SQLOLEDB;UID=" & Trim(Nom_Utsql) & ";Pwd=" & Trim(Mot_Psql) & ";Initial Catalog=" & Trim(BaseConsolide) & ";Data Source=" & Trim(Serveur) & "")
            OleSocieteConnect.Open()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Private Sub AfficheSchemasIntegrer()
        Dim i As Integer
        Dim Insertion As String
        Dim OleInsertCmd As OleDbCommand
        Dim OleAdaptaterschema As OleDbDataAdapter
        Dim OleSchemaDataset As DataSet
        Dim OledatableSchema As DataTable

        Try
            Insertion = "DELETE FROM TRI_IMPORTFICHIER"
            OleInsertCmd = New OleDbCommand(Insertion)
            OleInsertCmd.Connection = OleConnenection
            OleInsertCmd.ExecuteNonQuery()
            OleAdaptaterschema = New OleDbDataAdapter("select * from WIT_SCHEMA", OleConnenection)
            OleSchemaDataset = New DataSet
            OleAdaptaterschema.Fill(OleSchemaDataset)
            OledatableSchema = OleSchemaDataset.Tables(0)
            For i = 0 To OledatableSchema.Rows.Count - 1
                If Trim(OledatableSchema.Rows(i).Item("Cible")) = "FTP" Then
                    Dim OleAdaptaterftp As OleDbDataAdapter
                    Dim OleftpDataset As DataSet
                    Dim Oledatableftp As DataTable
                    OleAdaptaterftp = New OleDbDataAdapter("select * from FTPREPERTOIRE WHERE IDDossier=" & OledatableSchema.Rows(i).Item("IDDossier") & " And Traitement='IMPORT DOCUMENTTRANSFERT'", OleConnenection)
                    OleftpDataset = New DataSet
                    OleAdaptaterftp.Fill(OleftpDataset)
                    Oledatableftp = OleftpDataset.Tables(0)
                    If Oledatableftp.Rows.Count <> 0 Then
                        'blaise
                        If Directory.Exists(Oledatableftp.Rows(0).Item("Repertoire")) = True Then
                            DowloadFtp("WIT_SCHEMA", OledatableSchema.Rows(i).Item("IDDossier"), Oledatableftp.Rows(0).Item("Repertoire"))
                            OuvreLaListedeFichier(OledatableSchema.Rows(i).Item("Feuille_Excel"), OledatableSchema.Rows(i).Item("Mode"), OledatableSchema.Rows(i).Item("Type"), OledatableSchema.Rows(i).Item("NomFormat"), OledatableSchema.Rows(i).Item("CheminFormat"), Oledatableftp.Rows(0).Item("Repertoire"), Oledatableftp.Rows(0).Item("Repertoire"), Oledatableftp.Rows(0).Item("Repertoire"), OledatableSchema.Rows(i).Item("BaseCpta"), OledatableSchema.Rows(i).Item("BaseCial"), OledatableSchema.Rows(i).Item("Deplace"), OledatableSchema.Rows(i).Item("IDDossier"), OledatableSchema.Rows(i).Item("Cible"))
                        End If
                    End If
                Else
                    OuvreLaListedeFichier(OledatableSchema.Rows(i).Item("Feuille_Excel"), OledatableSchema.Rows(i).Item("Mode"), OledatableSchema.Rows(i).Item("Type"), OledatableSchema.Rows(i).Item("NomFormat"), OledatableSchema.Rows(i).Item("CheminFormat"), OledatableSchema.Rows(i).Item("NomFilexport"), OledatableSchema.Rows(i).Item("CheminFilexport"), OledatableSchema.Rows(i).Item("NomFilexport"), OledatableSchema.Rows(i).Item("BaseCpta"), OledatableSchema.Rows(i).Item("BaseCial"), OledatableSchema.Rows(i).Item("Deplace"), OledatableSchema.Rows(i).Item("IDDossier"), OledatableSchema.Rows(i).Item("Cible"))
                End If
            Next i
        Catch ex As Exception

        End Try
    End Sub
    Private Sub OuvreLaListedeFichier(ByRef Feuillexcel As String, ByRef ModeCreation As String, ByRef TypeFormalism As String, ByRef Formatname As String, ByRef PathFormat As String, ByRef NameDirectory As String, ByRef PathDirectory As String, ByRef Repert As String, ByRef BaseCpta As String, ByRef BaseCial As String, ByRef Deplacer As Boolean, ByRef Iddossier As Integer, ByRef CibleRe As String)
        Dim i As Integer
        Dim aLines() As String
        Dim Insertion As String
        Dim Datecreat, DateModif As Date
        Dim OleInsertCmd As OleDbCommand
        Try
            If Directory.Exists(PathDirectory) = True Then
                aLines = Directory.GetFiles(PathDirectory)
                For i = 0 To UBound(aLines)
                    Datecreat = Strings.FormatDateTime(File.GetCreationTime(aLines(i)), DateFormat.ShortDate)
                    DateModif = Strings.FormatDateTime(File.GetLastWriteTime(aLines(i)), DateFormat.ShortDate)
                    Insertion = "Insert into TRI_IMPORTFICHIER (Cible,PathFormat,CheminImport,CheminFichier,TypeFormat,Comptable,Commercial,NomFormat,Fichier,Mode,FeuilleExcel,Deplace,Dossier,DateCreation,DateModif) VALUES ('" & Join(Split(CibleRe, "'"), "''") & "','" & Join(Split(PathFormat, "'"), "''") & "','" & Join(Split(PathDirectory, "'"), "''") & "','" & Join(Split(aLines(i), "'"), "''") & "','" & Join(Split(TypeFormalism, "'"), "''") & "','" & Join(Split(BaseCpta, "'"), "''") & "','" & Join(Split(BaseCial, "'"), "''") & "','" & Join(Split(Formatname, "'"), "''") & "','" & Join(Split(System.IO.Path.GetFileName(Trim(aLines(i))), "'"), "''") & "','" & Join(Split(ModeCreation, "'"), "''") & "','" & Join(Split(Feuillexcel, "'"), "''") & "'," & Deplacer & ",'" & Join(Split(Repert, "'"), "''") & "','" & Datecreat & "','" & DateModif & "')"
                    OleInsertCmd = New OleDbCommand(Insertion)
                    OleInsertCmd.Connection = OleConnenection
                    OleInsertCmd.ExecuteNonQuery()
                Next i
                aLines = Nothing
            Else
                MsgBox("Ce Repertoire n'est pas valide! " & PathDirectory, MsgBoxStyle.Information, "Repertoire des Fichiers à Traiter")
            End If
        Catch ex As Exception
        End Try
    End Sub
    Private Sub Affichagefichier()
        Dim OleAdapsum As OleDbDataAdapter
        Dim OlesumDst As DataSet
        Dim OlesumData As DataTable
        Dim i, j As Integer
        iRow = 0
        DataListeIntegrer.Rows.Clear()
        OleAdapsum = New OleDbDataAdapter("select * from WIT_SCHEMA", OleConnenection)
        OlesumDst = New DataSet
        OleAdapsum.Fill(OlesumDst)
        OlesumData = OlesumDst.Tables(0)
        If OlesumData.Rows.Count <> 0 Then
            For i = 0 To OlesumData.Rows.Count - 1
                If OlesumData.Rows(i).Item("Cible") = "FTP" Then
                    Dim OleAdaptaterftp As OleDbDataAdapter
                    Dim OleftpDataset As DataSet
                    Dim Oledatableftp As DataTable
                    OleAdaptaterftp = New OleDbDataAdapter("select * from FTPREPERTOIRE WHERE IDDossier=" & OlesumData.Rows(i).Item("IDDossier") & " And Traitement='IMPORT DOCUMENTTRANSFERT'", OleConnenection)
                    OleftpDataset = New DataSet
                    OleAdaptaterftp.Fill(OleftpDataset)
                    Oledatableftp = OleftpDataset.Tables(0)
                    If Oledatableftp.Rows.Count <> 0 Then
                        If OlesumData.Rows(i).Item("TriNom") = True Then
                            If OlesumData.Rows(i).Item("TriCreation") = True Then
                                If OlesumData.Rows(i).Item("TriModification") = True Then
                                    Dim OleAdapsumAna As OleDbDataAdapter
                                    Dim OlesumAnaDst As DataSet
                                    Dim OlesumAnaData As DataTable
                                    OleAdapsumAna = New OleDbDataAdapter("select * from TRI_IMPORTFICHIER Where Cible='" & OlesumData.Rows(i).Item("Cible") & "' And  PathFormat='" & Trim(OlesumData.Rows(i).Item("CheminFormat")) & "' And CheminImport='" & Trim(Oledatableftp.Rows(0).Item("Repertoire")) & "' Order by Fichier ASC,DateCreation ASC,DateModif ASC", OleConnenection)
                                    OlesumAnaDst = New DataSet
                                    OleAdapsumAna.Fill(OlesumAnaDst)
                                    OlesumAnaData = OlesumAnaDst.Tables(0)
                                    If OlesumAnaData.Rows.Count <> 0 Then
                                        For j = 0 To OlesumAnaData.Rows.Count - 1
                                            DataListeIntegrer.RowCount = iRow + 1
                                            DataListeIntegrer.Rows(iRow).Cells("ID").Value = OlesumData.Rows(i).Item("IDDossier")
                                            DataListeIntegrer.Rows(iRow).Cells("Cible").Value = OlesumData.Rows(i).Item("Cible")
                                            DataListeIntegrer.Rows(iRow).Cells("FichierExport").Value = OlesumAnaData.Rows(j).Item("Fichier")
                                            DataListeIntegrer.Rows(iRow).Cells("CheminExport").Value = OlesumAnaData.Rows(j).Item("CheminFichier")
                                            DataListeIntegrer.Rows(iRow).Cells("Comptable").Value = OlesumAnaData.Rows(j).Item("Comptable")
                                            DataListeIntegrer.Rows(iRow).Cells("Commercial").Value = OlesumAnaData.Rows(j).Item("Commercial")
                                            DataListeIntegrer.Rows(iRow).Cells("Chemin").Value = OlesumAnaData.Rows(j).Item("PathFormat")
                                            DataListeIntegrer.Rows(iRow).Cells("NomFormat").Value = OlesumAnaData.Rows(j).Item("NomFormat")
                                            DataListeIntegrer.Rows(iRow).Cells("TypeFormat").Value = Afficheauuser(OlesumAnaData.Rows(j).Item("TypeFormat"))
                                            DataListeIntegrer.Rows(iRow).Cells("Mode").Value = OlesumAnaData.Rows(j).Item("Mode")
                                            DataListeIntegrer.Rows(iRow).Cells("FeuilleExcel").Value = OlesumAnaData.Rows(j).Item("FeuilleExcel")
                                            DataListeIntegrer.Rows(iRow).Cells("Deplace").Value = OlesumAnaData.Rows(j).Item("Deplace")
                                            DataListeIntegrer.Rows(iRow).Cells("Dossier").Value = OlesumAnaData.Rows(j).Item("Dossier")
                                            DataListeIntegrer.Rows(iRow).Cells("Valider").Value = True
                                            DataListeIntegrer.Rows(iRow).Cells("Modification").Value = Strings.FormatDateTime(OlesumAnaData.Rows(j).Item("DateModif"), DateFormat.ShortDate)
                                            DataListeIntegrer.Rows(iRow).Cells("Creation").Value = Strings.FormatDateTime(OlesumAnaData.Rows(j).Item("DateCreation"), DateFormat.ShortDate)
                                            iRow = iRow + 1
                                        Next j
                                    End If
                                Else
                                    Dim OleAdapsumAna As OleDbDataAdapter
                                    Dim OlesumAnaDst As DataSet
                                    Dim OlesumAnaData As DataTable
                                    OleAdapsumAna = New OleDbDataAdapter("select * from TRI_IMPORTFICHIER Where Cible='" & OlesumData.Rows(i).Item("Cible") & "' And  PathFormat='" & Trim(OlesumData.Rows(i).Item("CheminFormat")) & "' And CheminImport='" & Trim(Oledatableftp.Rows(0).Item("Repertoire")) & "' Order by Fichier ASC,DateCreation ASC", OleConnenection)
                                    OlesumAnaDst = New DataSet
                                    OleAdapsumAna.Fill(OlesumAnaDst)
                                    OlesumAnaData = OlesumAnaDst.Tables(0)
                                    If OlesumAnaData.Rows.Count <> 0 Then
                                        For j = 0 To OlesumAnaData.Rows.Count - 1
                                            DataListeIntegrer.RowCount = iRow + 1
                                            DataListeIntegrer.Rows(iRow).Cells("ID").Value = OlesumData.Rows(i).Item("IDDossier")
                                            DataListeIntegrer.Rows(iRow).Cells("Cible").Value = OlesumData.Rows(i).Item("Cible")
                                            DataListeIntegrer.Rows(iRow).Cells("FichierExport").Value = OlesumAnaData.Rows(j).Item("Fichier")
                                            DataListeIntegrer.Rows(iRow).Cells("CheminExport").Value = OlesumAnaData.Rows(j).Item("CheminFichier")
                                            DataListeIntegrer.Rows(iRow).Cells("Comptable").Value = OlesumAnaData.Rows(j).Item("Comptable")
                                            DataListeIntegrer.Rows(iRow).Cells("Commercial").Value = OlesumAnaData.Rows(j).Item("Commercial")
                                            DataListeIntegrer.Rows(iRow).Cells("Chemin").Value = OlesumAnaData.Rows(j).Item("PathFormat")
                                            DataListeIntegrer.Rows(iRow).Cells("NomFormat").Value = OlesumAnaData.Rows(j).Item("NomFormat")
                                            DataListeIntegrer.Rows(iRow).Cells("TypeFormat").Value = Afficheauuser(OlesumAnaData.Rows(j).Item("TypeFormat"))
                                            DataListeIntegrer.Rows(iRow).Cells("Mode").Value = OlesumAnaData.Rows(j).Item("Mode")
                                            DataListeIntegrer.Rows(iRow).Cells("FeuilleExcel").Value = OlesumAnaData.Rows(j).Item("FeuilleExcel")
                                            DataListeIntegrer.Rows(iRow).Cells("Deplace").Value = OlesumAnaData.Rows(j).Item("Deplace")
                                            DataListeIntegrer.Rows(iRow).Cells("Dossier").Value = OlesumAnaData.Rows(j).Item("Dossier")
                                            DataListeIntegrer.Rows(iRow).Cells("Valider").Value = True
                                            DataListeIntegrer.Rows(iRow).Cells("Modification").Value = Strings.FormatDateTime(OlesumAnaData.Rows(j).Item("DateModif"), DateFormat.ShortDate)
                                            DataListeIntegrer.Rows(iRow).Cells("Creation").Value = Strings.FormatDateTime(OlesumAnaData.Rows(j).Item("DateCreation"), DateFormat.ShortDate)
                                            iRow = iRow + 1
                                        Next j
                                    End If
                                End If
                            Else
                                If OlesumData.Rows(i).Item("TriModification") = True Then
                                    Dim OleAdapsumAna As OleDbDataAdapter
                                    Dim OlesumAnaDst As DataSet
                                    Dim OlesumAnaData As DataTable
                                    OleAdapsumAna = New OleDbDataAdapter("select * from TRI_IMPORTFICHIER Where Cible='" & OlesumData.Rows(i).Item("Cible") & "' And  PathFormat='" & Trim(OlesumData.Rows(i).Item("CheminFormat")) & "' And CheminImport='" & Trim(Oledatableftp.Rows(0).Item("Repertoire")) & "' Order by Fichier ASC,DateModif ASC", OleConnenection)
                                    OlesumAnaDst = New DataSet
                                    OleAdapsumAna.Fill(OlesumAnaDst)
                                    OlesumAnaData = OlesumAnaDst.Tables(0)
                                    If OlesumAnaData.Rows.Count <> 0 Then
                                        For j = 0 To OlesumAnaData.Rows.Count - 1
                                            DataListeIntegrer.RowCount = iRow + 1
                                            DataListeIntegrer.Rows(iRow).Cells("ID").Value = OlesumData.Rows(i).Item("IDDossier")
                                            DataListeIntegrer.Rows(iRow).Cells("Cible").Value = OlesumData.Rows(i).Item("Cible")
                                            DataListeIntegrer.Rows(iRow).Cells("FichierExport").Value = OlesumAnaData.Rows(j).Item("Fichier")
                                            DataListeIntegrer.Rows(iRow).Cells("CheminExport").Value = OlesumAnaData.Rows(j).Item("CheminFichier")
                                            DataListeIntegrer.Rows(iRow).Cells("Comptable").Value = OlesumAnaData.Rows(j).Item("Comptable")
                                            DataListeIntegrer.Rows(iRow).Cells("Commercial").Value = OlesumAnaData.Rows(j).Item("Commercial")
                                            DataListeIntegrer.Rows(iRow).Cells("Chemin").Value = OlesumAnaData.Rows(j).Item("PathFormat")
                                            DataListeIntegrer.Rows(iRow).Cells("NomFormat").Value = OlesumAnaData.Rows(j).Item("NomFormat")
                                            DataListeIntegrer.Rows(iRow).Cells("TypeFormat").Value = Afficheauuser(OlesumAnaData.Rows(j).Item("TypeFormat"))
                                            DataListeIntegrer.Rows(iRow).Cells("Mode").Value = OlesumAnaData.Rows(j).Item("Mode")
                                            DataListeIntegrer.Rows(iRow).Cells("FeuilleExcel").Value = OlesumAnaData.Rows(j).Item("FeuilleExcel")
                                            DataListeIntegrer.Rows(iRow).Cells("Deplace").Value = OlesumAnaData.Rows(j).Item("Deplace")
                                            DataListeIntegrer.Rows(iRow).Cells("Dossier").Value = OlesumAnaData.Rows(j).Item("Dossier")
                                            DataListeIntegrer.Rows(iRow).Cells("Valider").Value = True
                                            DataListeIntegrer.Rows(iRow).Cells("Modification").Value = Strings.FormatDateTime(OlesumAnaData.Rows(j).Item("DateModif"), DateFormat.ShortDate)
                                            DataListeIntegrer.Rows(iRow).Cells("Creation").Value = Strings.FormatDateTime(OlesumAnaData.Rows(j).Item("DateCreation"), DateFormat.ShortDate)
                                            iRow = iRow + 1
                                        Next j
                                    End If
                                Else
                                    Dim OleAdapsumAna As OleDbDataAdapter
                                    Dim OlesumAnaDst As DataSet
                                    Dim OlesumAnaData As DataTable
                                    OleAdapsumAna = New OleDbDataAdapter("select * from TRI_IMPORTFICHIER Where Cible='" & OlesumData.Rows(i).Item("Cible") & "' And  PathFormat='" & Trim(OlesumData.Rows(i).Item("CheminFormat")) & "' And CheminImport='" & Trim(Oledatableftp.Rows(0).Item("Repertoire")) & "' Order by Fichier ASC", OleConnenection)
                                    OlesumAnaDst = New DataSet
                                    OleAdapsumAna.Fill(OlesumAnaDst)
                                    OlesumAnaData = OlesumAnaDst.Tables(0)
                                    If OlesumAnaData.Rows.Count <> 0 Then
                                        For j = 0 To OlesumAnaData.Rows.Count - 1
                                            DataListeIntegrer.RowCount = iRow + 1
                                            DataListeIntegrer.Rows(iRow).Cells("ID").Value = OlesumData.Rows(i).Item("IDDossier")
                                            DataListeIntegrer.Rows(iRow).Cells("Cible").Value = OlesumData.Rows(i).Item("Cible")
                                            DataListeIntegrer.Rows(iRow).Cells("FichierExport").Value = OlesumAnaData.Rows(j).Item("Fichier")
                                            DataListeIntegrer.Rows(iRow).Cells("CheminExport").Value = OlesumAnaData.Rows(j).Item("CheminFichier")
                                            DataListeIntegrer.Rows(iRow).Cells("Comptable").Value = OlesumAnaData.Rows(j).Item("Comptable")
                                            DataListeIntegrer.Rows(iRow).Cells("Commercial").Value = OlesumAnaData.Rows(j).Item("Commercial")
                                            DataListeIntegrer.Rows(iRow).Cells("Chemin").Value = OlesumAnaData.Rows(j).Item("PathFormat")
                                            DataListeIntegrer.Rows(iRow).Cells("NomFormat").Value = OlesumAnaData.Rows(j).Item("NomFormat")
                                            DataListeIntegrer.Rows(iRow).Cells("TypeFormat").Value = Afficheauuser(OlesumAnaData.Rows(j).Item("TypeFormat"))
                                            DataListeIntegrer.Rows(iRow).Cells("Mode").Value = OlesumAnaData.Rows(j).Item("Mode")
                                            DataListeIntegrer.Rows(iRow).Cells("FeuilleExcel").Value = OlesumAnaData.Rows(j).Item("FeuilleExcel")
                                            DataListeIntegrer.Rows(iRow).Cells("Deplace").Value = OlesumAnaData.Rows(j).Item("Deplace")
                                            DataListeIntegrer.Rows(iRow).Cells("Dossier").Value = OlesumAnaData.Rows(j).Item("Dossier")
                                            DataListeIntegrer.Rows(iRow).Cells("Valider").Value = True
                                            DataListeIntegrer.Rows(iRow).Cells("Modification").Value = Strings.FormatDateTime(OlesumAnaData.Rows(j).Item("DateModif"), DateFormat.ShortDate)
                                            DataListeIntegrer.Rows(iRow).Cells("Creation").Value = Strings.FormatDateTime(OlesumAnaData.Rows(j).Item("DateCreation"), DateFormat.ShortDate)
                                            iRow = iRow + 1
                                        Next j
                                    End If
                                End If
                            End If
                        Else
                            If OlesumData.Rows(i).Item("TriCreation") = True Then
                                If OlesumData.Rows(i).Item("TriModification") = True Then
                                    Dim OleAdapsumAna As OleDbDataAdapter
                                    Dim OlesumAnaDst As DataSet
                                    Dim OlesumAnaData As DataTable
                                    OleAdapsumAna = New OleDbDataAdapter("select * from TRI_IMPORTFICHIER Where Cible='" & OlesumData.Rows(i).Item("Cible") & "' And  PathFormat='" & Trim(OlesumData.Rows(i).Item("CheminFormat")) & "' And CheminImport='" & Trim(Oledatableftp.Rows(0).Item("Repertoire")) & "' Order by DateCreation ASC,DateModif ASC", OleConnenection)
                                    OlesumAnaDst = New DataSet
                                    OleAdapsumAna.Fill(OlesumAnaDst)
                                    OlesumAnaData = OlesumAnaDst.Tables(0)
                                    If OlesumAnaData.Rows.Count <> 0 Then
                                        For j = 0 To OlesumAnaData.Rows.Count - 1
                                            DataListeIntegrer.RowCount = iRow + 1
                                            DataListeIntegrer.Rows(iRow).Cells("ID").Value = OlesumData.Rows(i).Item("IDDossier")
                                            DataListeIntegrer.Rows(iRow).Cells("Cible").Value = OlesumData.Rows(i).Item("Cible")
                                            DataListeIntegrer.Rows(iRow).Cells("FichierExport").Value = OlesumAnaData.Rows(j).Item("Fichier")
                                            DataListeIntegrer.Rows(iRow).Cells("CheminExport").Value = OlesumAnaData.Rows(j).Item("CheminFichier")
                                            DataListeIntegrer.Rows(iRow).Cells("Comptable").Value = OlesumAnaData.Rows(j).Item("Comptable")
                                            DataListeIntegrer.Rows(iRow).Cells("Commercial").Value = OlesumAnaData.Rows(j).Item("Commercial")
                                            DataListeIntegrer.Rows(iRow).Cells("Chemin").Value = OlesumAnaData.Rows(j).Item("PathFormat")
                                            DataListeIntegrer.Rows(iRow).Cells("NomFormat").Value = OlesumAnaData.Rows(j).Item("NomFormat")
                                            DataListeIntegrer.Rows(iRow).Cells("TypeFormat").Value = Afficheauuser(OlesumAnaData.Rows(j).Item("TypeFormat"))
                                            DataListeIntegrer.Rows(iRow).Cells("Mode").Value = OlesumAnaData.Rows(j).Item("Mode")
                                            DataListeIntegrer.Rows(iRow).Cells("FeuilleExcel").Value = OlesumAnaData.Rows(j).Item("FeuilleExcel")
                                            DataListeIntegrer.Rows(iRow).Cells("Deplace").Value = OlesumAnaData.Rows(j).Item("Deplace")
                                            DataListeIntegrer.Rows(iRow).Cells("Dossier").Value = OlesumAnaData.Rows(j).Item("Dossier")
                                            DataListeIntegrer.Rows(iRow).Cells("Valider").Value = True
                                            DataListeIntegrer.Rows(iRow).Cells("Modification").Value = Strings.FormatDateTime(OlesumAnaData.Rows(j).Item("DateModif"), DateFormat.ShortDate)
                                            DataListeIntegrer.Rows(iRow).Cells("Creation").Value = Strings.FormatDateTime(OlesumAnaData.Rows(j).Item("DateCreation"), DateFormat.ShortDate)
                                            iRow = iRow + 1
                                        Next j
                                    End If
                                Else
                                    Dim OleAdapsumAna As OleDbDataAdapter
                                    Dim OlesumAnaDst As DataSet
                                    Dim OlesumAnaData As DataTable
                                    OleAdapsumAna = New OleDbDataAdapter("select * from TRI_IMPORTFICHIER Where Cible='" & OlesumData.Rows(i).Item("Cible") & "' And  PathFormat='" & Trim(OlesumData.Rows(i).Item("CheminFormat")) & "' And CheminImport='" & Trim(Oledatableftp.Rows(0).Item("Repertoire")) & "' Order by DateCreation ASC", OleConnenection)
                                    OlesumAnaDst = New DataSet
                                    OleAdapsumAna.Fill(OlesumAnaDst)
                                    OlesumAnaData = OlesumAnaDst.Tables(0)
                                    If OlesumAnaData.Rows.Count <> 0 Then
                                        For j = 0 To OlesumAnaData.Rows.Count - 1
                                            DataListeIntegrer.RowCount = iRow + 1
                                            DataListeIntegrer.Rows(iRow).Cells("ID").Value = OlesumData.Rows(i).Item("IDDossier")
                                            DataListeIntegrer.Rows(iRow).Cells("Cible").Value = OlesumData.Rows(i).Item("Cible")
                                            DataListeIntegrer.Rows(iRow).Cells("FichierExport").Value = OlesumAnaData.Rows(j).Item("Fichier")
                                            DataListeIntegrer.Rows(iRow).Cells("CheminExport").Value = OlesumAnaData.Rows(j).Item("CheminFichier")
                                            DataListeIntegrer.Rows(iRow).Cells("Comptable").Value = OlesumAnaData.Rows(j).Item("Comptable")
                                            DataListeIntegrer.Rows(iRow).Cells("Commercial").Value = OlesumAnaData.Rows(j).Item("Commercial")
                                            DataListeIntegrer.Rows(iRow).Cells("Chemin").Value = OlesumAnaData.Rows(j).Item("PathFormat")
                                            DataListeIntegrer.Rows(iRow).Cells("NomFormat").Value = OlesumAnaData.Rows(j).Item("NomFormat")
                                            DataListeIntegrer.Rows(iRow).Cells("TypeFormat").Value = Afficheauuser(OlesumAnaData.Rows(j).Item("TypeFormat"))
                                            DataListeIntegrer.Rows(iRow).Cells("Mode").Value = OlesumAnaData.Rows(j).Item("Mode")
                                            DataListeIntegrer.Rows(iRow).Cells("FeuilleExcel").Value = OlesumAnaData.Rows(j).Item("FeuilleExcel")
                                            DataListeIntegrer.Rows(iRow).Cells("Deplace").Value = OlesumAnaData.Rows(j).Item("Deplace")
                                            DataListeIntegrer.Rows(iRow).Cells("Dossier").Value = OlesumAnaData.Rows(j).Item("Dossier")
                                            DataListeIntegrer.Rows(iRow).Cells("Valider").Value = True
                                            DataListeIntegrer.Rows(iRow).Cells("Modification").Value = Strings.FormatDateTime(OlesumAnaData.Rows(j).Item("DateModif"), DateFormat.ShortDate)
                                            DataListeIntegrer.Rows(iRow).Cells("Creation").Value = Strings.FormatDateTime(OlesumAnaData.Rows(j).Item("DateCreation"), DateFormat.ShortDate)
                                            iRow = iRow + 1
                                        Next j
                                    End If
                                End If
                            Else
                                If OlesumData.Rows(i).Item("TriModification") = True Then
                                    Dim OleAdapsumAna As OleDbDataAdapter
                                    Dim OlesumAnaDst As DataSet
                                    Dim OlesumAnaData As DataTable
                                    OleAdapsumAna = New OleDbDataAdapter("select * from TRI_IMPORTFICHIER Where Cible='" & OlesumData.Rows(i).Item("Cible") & "' And  PathFormat='" & Trim(OlesumData.Rows(i).Item("CheminFormat")) & "' And CheminImport='" & Trim(Oledatableftp.Rows(0).Item("Repertoire")) & "' Order by DateModif ASC", OleConnenection)
                                    OlesumAnaDst = New DataSet
                                    OleAdapsumAna.Fill(OlesumAnaDst)
                                    OlesumAnaData = OlesumAnaDst.Tables(0)
                                    If OlesumAnaData.Rows.Count <> 0 Then
                                        For j = 0 To OlesumAnaData.Rows.Count - 1
                                            DataListeIntegrer.RowCount = iRow + 1
                                            DataListeIntegrer.Rows(iRow).Cells("ID").Value = OlesumData.Rows(i).Item("IDDossier")
                                            DataListeIntegrer.Rows(iRow).Cells("Cible").Value = OlesumData.Rows(i).Item("Cible")
                                            DataListeIntegrer.Rows(iRow).Cells("FichierExport").Value = OlesumAnaData.Rows(j).Item("Fichier")
                                            DataListeIntegrer.Rows(iRow).Cells("CheminExport").Value = OlesumAnaData.Rows(j).Item("CheminFichier")
                                            DataListeIntegrer.Rows(iRow).Cells("Comptable").Value = OlesumAnaData.Rows(j).Item("Comptable")
                                            DataListeIntegrer.Rows(iRow).Cells("Commercial").Value = OlesumAnaData.Rows(j).Item("Commercial")
                                            DataListeIntegrer.Rows(iRow).Cells("Chemin").Value = OlesumAnaData.Rows(j).Item("PathFormat")
                                            DataListeIntegrer.Rows(iRow).Cells("NomFormat").Value = OlesumAnaData.Rows(j).Item("NomFormat")
                                            DataListeIntegrer.Rows(iRow).Cells("TypeFormat").Value = Afficheauuser(OlesumAnaData.Rows(j).Item("TypeFormat"))
                                            DataListeIntegrer.Rows(iRow).Cells("Mode").Value = OlesumAnaData.Rows(j).Item("Mode")
                                            DataListeIntegrer.Rows(iRow).Cells("FeuilleExcel").Value = OlesumAnaData.Rows(j).Item("FeuilleExcel")
                                            DataListeIntegrer.Rows(iRow).Cells("Deplace").Value = OlesumAnaData.Rows(j).Item("Deplace")
                                            DataListeIntegrer.Rows(iRow).Cells("Dossier").Value = OlesumAnaData.Rows(j).Item("Dossier")
                                            DataListeIntegrer.Rows(iRow).Cells("Valider").Value = True
                                            DataListeIntegrer.Rows(iRow).Cells("Modification").Value = Strings.FormatDateTime(OlesumAnaData.Rows(j).Item("DateModif"), DateFormat.ShortDate)
                                            DataListeIntegrer.Rows(iRow).Cells("Creation").Value = Strings.FormatDateTime(OlesumAnaData.Rows(j).Item("DateCreation"), DateFormat.ShortDate)
                                            iRow = iRow + 1
                                        Next j
                                    End If
                                Else
                                    Dim OleAdapsumAna As OleDbDataAdapter
                                    Dim OlesumAnaDst As DataSet
                                    Dim OlesumAnaData As DataTable
                                    OleAdapsumAna = New OleDbDataAdapter("select * from TRI_IMPORTFICHIER Where Cible='" & OlesumData.Rows(i).Item("Cible") & "' And  PathFormat='" & Trim(OlesumData.Rows(i).Item("CheminFormat")) & "' And CheminImport='" & Trim(Oledatableftp.Rows(0).Item("Repertoire")) & "'", OleConnenection)
                                    OlesumAnaDst = New DataSet
                                    OleAdapsumAna.Fill(OlesumAnaDst)
                                    OlesumAnaData = OlesumAnaDst.Tables(0)
                                    If OlesumAnaData.Rows.Count <> 0 Then
                                        For j = 0 To OlesumAnaData.Rows.Count - 1
                                            DataListeIntegrer.RowCount = iRow + 1
                                            DataListeIntegrer.Rows(iRow).Cells("ID").Value = OlesumData.Rows(i).Item("IDDossier")
                                            DataListeIntegrer.Rows(iRow).Cells("Cible").Value = OlesumData.Rows(i).Item("Cible")
                                            DataListeIntegrer.Rows(iRow).Cells("FichierExport").Value = OlesumAnaData.Rows(j).Item("Fichier")
                                            DataListeIntegrer.Rows(iRow).Cells("CheminExport").Value = OlesumAnaData.Rows(j).Item("CheminFichier")
                                            DataListeIntegrer.Rows(iRow).Cells("Comptable").Value = OlesumAnaData.Rows(j).Item("Comptable")
                                            DataListeIntegrer.Rows(iRow).Cells("Commercial").Value = OlesumAnaData.Rows(j).Item("Commercial")
                                            DataListeIntegrer.Rows(iRow).Cells("Chemin").Value = OlesumAnaData.Rows(j).Item("PathFormat")
                                            DataListeIntegrer.Rows(iRow).Cells("NomFormat").Value = OlesumAnaData.Rows(j).Item("NomFormat")
                                            DataListeIntegrer.Rows(iRow).Cells("TypeFormat").Value = Afficheauuser(OlesumAnaData.Rows(j).Item("TypeFormat"))
                                            DataListeIntegrer.Rows(iRow).Cells("Mode").Value = OlesumAnaData.Rows(j).Item("Mode")
                                            DataListeIntegrer.Rows(iRow).Cells("FeuilleExcel").Value = OlesumAnaData.Rows(j).Item("FeuilleExcel")
                                            DataListeIntegrer.Rows(iRow).Cells("Deplace").Value = OlesumAnaData.Rows(j).Item("Deplace")
                                            DataListeIntegrer.Rows(iRow).Cells("Dossier").Value = OlesumAnaData.Rows(j).Item("Dossier")
                                            DataListeIntegrer.Rows(iRow).Cells("Valider").Value = True
                                            DataListeIntegrer.Rows(iRow).Cells("Modification").Value = Strings.FormatDateTime(OlesumAnaData.Rows(j).Item("DateModif"), DateFormat.ShortDate)
                                            DataListeIntegrer.Rows(iRow).Cells("Creation").Value = Strings.FormatDateTime(OlesumAnaData.Rows(j).Item("DateCreation"), DateFormat.ShortDate)
                                            iRow = iRow + 1
                                        Next j
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else
                    If OlesumData.Rows(i).Item("TriNom") = True Then
                        If OlesumData.Rows(i).Item("TriCreation") = True Then
                            If OlesumData.Rows(i).Item("TriModification") = True Then
                                Dim OleAdapsumAna As OleDbDataAdapter
                                Dim OlesumAnaDst As DataSet
                                Dim OlesumAnaData As DataTable
                                OleAdapsumAna = New OleDbDataAdapter("select * from TRI_IMPORTFICHIER Where Cible='" & OlesumData.Rows(i).Item("Cible") & "' And  PathFormat='" & Trim(OlesumData.Rows(i).Item("CheminFormat")) & "' And CheminImport='" & Trim(OlesumData.Rows(i).Item("CheminFilexport")) & "' Order by Fichier ASC,DateCreation ASC,DateModif ASC", OleConnenection)
                                OlesumAnaDst = New DataSet
                                OleAdapsumAna.Fill(OlesumAnaDst)
                                OlesumAnaData = OlesumAnaDst.Tables(0)
                                If OlesumAnaData.Rows.Count <> 0 Then
                                    For j = 0 To OlesumAnaData.Rows.Count - 1
                                        DataListeIntegrer.RowCount = iRow + 1
                                        DataListeIntegrer.Rows(iRow).Cells("ID").Value = OlesumData.Rows(i).Item("IDDossier")
                                        DataListeIntegrer.Rows(iRow).Cells("Cible").Value = OlesumData.Rows(i).Item("Cible")
                                        DataListeIntegrer.Rows(iRow).Cells("FichierExport").Value = OlesumAnaData.Rows(j).Item("Fichier")
                                        DataListeIntegrer.Rows(iRow).Cells("CheminExport").Value = OlesumAnaData.Rows(j).Item("CheminFichier")
                                        DataListeIntegrer.Rows(iRow).Cells("Comptable").Value = OlesumAnaData.Rows(j).Item("Comptable")
                                        DataListeIntegrer.Rows(iRow).Cells("Commercial").Value = OlesumAnaData.Rows(j).Item("Commercial")
                                        DataListeIntegrer.Rows(iRow).Cells("Chemin").Value = OlesumAnaData.Rows(j).Item("PathFormat")
                                        DataListeIntegrer.Rows(iRow).Cells("NomFormat").Value = OlesumAnaData.Rows(j).Item("NomFormat")
                                        DataListeIntegrer.Rows(iRow).Cells("TypeFormat").Value = Afficheauuser(OlesumAnaData.Rows(j).Item("TypeFormat"))
                                        DataListeIntegrer.Rows(iRow).Cells("Mode").Value = OlesumAnaData.Rows(j).Item("Mode")
                                        DataListeIntegrer.Rows(iRow).Cells("FeuilleExcel").Value = OlesumAnaData.Rows(j).Item("FeuilleExcel")
                                        DataListeIntegrer.Rows(iRow).Cells("Deplace").Value = OlesumAnaData.Rows(j).Item("Deplace")
                                        DataListeIntegrer.Rows(iRow).Cells("Dossier").Value = OlesumAnaData.Rows(j).Item("Dossier")
                                        DataListeIntegrer.Rows(iRow).Cells("Valider").Value = True
                                        DataListeIntegrer.Rows(iRow).Cells("Modification").Value = Strings.FormatDateTime(OlesumAnaData.Rows(j).Item("DateModif"), DateFormat.ShortDate)
                                        DataListeIntegrer.Rows(iRow).Cells("Creation").Value = Strings.FormatDateTime(OlesumAnaData.Rows(j).Item("DateCreation"), DateFormat.ShortDate)
                                        iRow = iRow + 1
                                    Next j
                                End If
                            Else
                                Dim OleAdapsumAna As OleDbDataAdapter
                                Dim OlesumAnaDst As DataSet
                                Dim OlesumAnaData As DataTable
                                OleAdapsumAna = New OleDbDataAdapter("select * from TRI_IMPORTFICHIER Where Cible='" & OlesumData.Rows(i).Item("Cible") & "' And  PathFormat='" & Trim(OlesumData.Rows(i).Item("CheminFormat")) & "' And CheminImport='" & Trim(OlesumData.Rows(i).Item("CheminFilexport")) & "' Order by Fichier ASC,DateCreation ASC", OleConnenection)
                                OlesumAnaDst = New DataSet
                                OleAdapsumAna.Fill(OlesumAnaDst)
                                OlesumAnaData = OlesumAnaDst.Tables(0)
                                If OlesumAnaData.Rows.Count <> 0 Then
                                    For j = 0 To OlesumAnaData.Rows.Count - 1
                                        DataListeIntegrer.RowCount = iRow + 1
                                        DataListeIntegrer.Rows(iRow).Cells("ID").Value = OlesumData.Rows(i).Item("IDDossier")
                                        DataListeIntegrer.Rows(iRow).Cells("Cible").Value = OlesumData.Rows(i).Item("Cible")
                                        DataListeIntegrer.Rows(iRow).Cells("FichierExport").Value = OlesumAnaData.Rows(j).Item("Fichier")
                                        DataListeIntegrer.Rows(iRow).Cells("CheminExport").Value = OlesumAnaData.Rows(j).Item("CheminFichier")
                                        DataListeIntegrer.Rows(iRow).Cells("Comptable").Value = OlesumAnaData.Rows(j).Item("Comptable")
                                        DataListeIntegrer.Rows(iRow).Cells("Commercial").Value = OlesumAnaData.Rows(j).Item("Commercial")
                                        DataListeIntegrer.Rows(iRow).Cells("Chemin").Value = OlesumAnaData.Rows(j).Item("PathFormat")
                                        DataListeIntegrer.Rows(iRow).Cells("NomFormat").Value = OlesumAnaData.Rows(j).Item("NomFormat")
                                        DataListeIntegrer.Rows(iRow).Cells("TypeFormat").Value = Afficheauuser(OlesumAnaData.Rows(j).Item("TypeFormat"))
                                        DataListeIntegrer.Rows(iRow).Cells("Mode").Value = OlesumAnaData.Rows(j).Item("Mode")
                                        DataListeIntegrer.Rows(iRow).Cells("FeuilleExcel").Value = OlesumAnaData.Rows(j).Item("FeuilleExcel")
                                        DataListeIntegrer.Rows(iRow).Cells("Deplace").Value = OlesumAnaData.Rows(j).Item("Deplace")
                                        DataListeIntegrer.Rows(iRow).Cells("Dossier").Value = OlesumAnaData.Rows(j).Item("Dossier")
                                        DataListeIntegrer.Rows(iRow).Cells("Valider").Value = True
                                        DataListeIntegrer.Rows(iRow).Cells("Modification").Value = Strings.FormatDateTime(OlesumAnaData.Rows(j).Item("DateModif"), DateFormat.ShortDate)
                                        DataListeIntegrer.Rows(iRow).Cells("Creation").Value = Strings.FormatDateTime(OlesumAnaData.Rows(j).Item("DateCreation"), DateFormat.ShortDate)
                                        iRow = iRow + 1
                                    Next j
                                End If
                            End If
                        Else
                            If OlesumData.Rows(i).Item("TriModification") = True Then
                                Dim OleAdapsumAna As OleDbDataAdapter
                                Dim OlesumAnaDst As DataSet
                                Dim OlesumAnaData As DataTable
                                OleAdapsumAna = New OleDbDataAdapter("select * from TRI_IMPORTFICHIER Where Cible='" & OlesumData.Rows(i).Item("Cible") & "' And  PathFormat='" & Trim(OlesumData.Rows(i).Item("CheminFormat")) & "' And CheminImport='" & Trim(OlesumData.Rows(i).Item("CheminFilexport")) & "' Order by Fichier ASC,DateModif ASC", OleConnenection)
                                OlesumAnaDst = New DataSet
                                OleAdapsumAna.Fill(OlesumAnaDst)
                                OlesumAnaData = OlesumAnaDst.Tables(0)
                                If OlesumAnaData.Rows.Count <> 0 Then
                                    For j = 0 To OlesumAnaData.Rows.Count - 1
                                        DataListeIntegrer.RowCount = iRow + 1
                                        DataListeIntegrer.Rows(iRow).Cells("ID").Value = OlesumData.Rows(i).Item("IDDossier")
                                        DataListeIntegrer.Rows(iRow).Cells("Cible").Value = OlesumData.Rows(i).Item("Cible")
                                        DataListeIntegrer.Rows(iRow).Cells("FichierExport").Value = OlesumAnaData.Rows(j).Item("Fichier")
                                        DataListeIntegrer.Rows(iRow).Cells("CheminExport").Value = OlesumAnaData.Rows(j).Item("CheminFichier")
                                        DataListeIntegrer.Rows(iRow).Cells("Comptable").Value = OlesumAnaData.Rows(j).Item("Comptable")
                                        DataListeIntegrer.Rows(iRow).Cells("Commercial").Value = OlesumAnaData.Rows(j).Item("Commercial")
                                        DataListeIntegrer.Rows(iRow).Cells("Chemin").Value = OlesumAnaData.Rows(j).Item("PathFormat")
                                        DataListeIntegrer.Rows(iRow).Cells("NomFormat").Value = OlesumAnaData.Rows(j).Item("NomFormat")
                                        DataListeIntegrer.Rows(iRow).Cells("TypeFormat").Value = Afficheauuser(OlesumAnaData.Rows(j).Item("TypeFormat"))
                                        DataListeIntegrer.Rows(iRow).Cells("Mode").Value = OlesumAnaData.Rows(j).Item("Mode")
                                        DataListeIntegrer.Rows(iRow).Cells("FeuilleExcel").Value = OlesumAnaData.Rows(j).Item("FeuilleExcel")
                                        DataListeIntegrer.Rows(iRow).Cells("Deplace").Value = OlesumAnaData.Rows(j).Item("Deplace")
                                        DataListeIntegrer.Rows(iRow).Cells("Dossier").Value = OlesumAnaData.Rows(j).Item("Dossier")
                                        DataListeIntegrer.Rows(iRow).Cells("Valider").Value = True
                                        DataListeIntegrer.Rows(iRow).Cells("Modification").Value = Strings.FormatDateTime(OlesumAnaData.Rows(j).Item("DateModif"), DateFormat.ShortDate)
                                        DataListeIntegrer.Rows(iRow).Cells("Creation").Value = Strings.FormatDateTime(OlesumAnaData.Rows(j).Item("DateCreation"), DateFormat.ShortDate)
                                        iRow = iRow + 1
                                    Next j
                                End If
                            Else
                                Dim OleAdapsumAna As OleDbDataAdapter
                                Dim OlesumAnaDst As DataSet
                                Dim OlesumAnaData As DataTable
                                OleAdapsumAna = New OleDbDataAdapter("select * from TRI_IMPORTFICHIER Where Cible='" & OlesumData.Rows(i).Item("Cible") & "' And  PathFormat='" & Trim(OlesumData.Rows(i).Item("CheminFormat")) & "' And CheminImport='" & Trim(OlesumData.Rows(i).Item("CheminFilexport")) & "' Order by Fichier ASC", OleConnenection)
                                OlesumAnaDst = New DataSet
                                OleAdapsumAna.Fill(OlesumAnaDst)
                                OlesumAnaData = OlesumAnaDst.Tables(0)
                                If OlesumAnaData.Rows.Count <> 0 Then
                                    For j = 0 To OlesumAnaData.Rows.Count - 1
                                        DataListeIntegrer.RowCount = iRow + 1
                                        DataListeIntegrer.Rows(iRow).Cells("ID").Value = OlesumData.Rows(i).Item("IDDossier")
                                        DataListeIntegrer.Rows(iRow).Cells("Cible").Value = OlesumData.Rows(i).Item("Cible")
                                        DataListeIntegrer.Rows(iRow).Cells("FichierExport").Value = OlesumAnaData.Rows(j).Item("Fichier")
                                        DataListeIntegrer.Rows(iRow).Cells("CheminExport").Value = OlesumAnaData.Rows(j).Item("CheminFichier")
                                        DataListeIntegrer.Rows(iRow).Cells("Comptable").Value = OlesumAnaData.Rows(j).Item("Comptable")
                                        DataListeIntegrer.Rows(iRow).Cells("Commercial").Value = OlesumAnaData.Rows(j).Item("Commercial")
                                        DataListeIntegrer.Rows(iRow).Cells("Chemin").Value = OlesumAnaData.Rows(j).Item("PathFormat")
                                        DataListeIntegrer.Rows(iRow).Cells("NomFormat").Value = OlesumAnaData.Rows(j).Item("NomFormat")
                                        DataListeIntegrer.Rows(iRow).Cells("TypeFormat").Value = Afficheauuser(OlesumAnaData.Rows(j).Item("TypeFormat"))
                                        DataListeIntegrer.Rows(iRow).Cells("Mode").Value = OlesumAnaData.Rows(j).Item("Mode")
                                        DataListeIntegrer.Rows(iRow).Cells("FeuilleExcel").Value = OlesumAnaData.Rows(j).Item("FeuilleExcel")
                                        DataListeIntegrer.Rows(iRow).Cells("Deplace").Value = OlesumAnaData.Rows(j).Item("Deplace")
                                        DataListeIntegrer.Rows(iRow).Cells("Dossier").Value = OlesumAnaData.Rows(j).Item("Dossier")
                                        DataListeIntegrer.Rows(iRow).Cells("Valider").Value = True
                                        DataListeIntegrer.Rows(iRow).Cells("Modification").Value = Strings.FormatDateTime(OlesumAnaData.Rows(j).Item("DateModif"), DateFormat.ShortDate)
                                        DataListeIntegrer.Rows(iRow).Cells("Creation").Value = Strings.FormatDateTime(OlesumAnaData.Rows(j).Item("DateCreation"), DateFormat.ShortDate)
                                        iRow = iRow + 1
                                    Next j
                                End If
                            End If
                        End If
                    Else
                        If OlesumData.Rows(i).Item("TriCreation") = True Then
                            If OlesumData.Rows(i).Item("TriModification") = True Then
                                Dim OleAdapsumAna As OleDbDataAdapter
                                Dim OlesumAnaDst As DataSet
                                Dim OlesumAnaData As DataTable
                                OleAdapsumAna = New OleDbDataAdapter("select * from TRI_IMPORTFICHIER Where Cible='" & OlesumData.Rows(i).Item("Cible") & "' And  PathFormat='" & Trim(OlesumData.Rows(i).Item("CheminFormat")) & "' And CheminImport='" & Trim(OlesumData.Rows(i).Item("CheminFilexport")) & "' Order by DateCreation ASC,DateModif ASC", OleConnenection)
                                OlesumAnaDst = New DataSet
                                OleAdapsumAna.Fill(OlesumAnaDst)
                                OlesumAnaData = OlesumAnaDst.Tables(0)
                                If OlesumAnaData.Rows.Count <> 0 Then
                                    For j = 0 To OlesumAnaData.Rows.Count - 1
                                        DataListeIntegrer.RowCount = iRow + 1
                                        DataListeIntegrer.Rows(iRow).Cells("ID").Value = OlesumData.Rows(i).Item("IDDossier")
                                        DataListeIntegrer.Rows(iRow).Cells("Cible").Value = OlesumData.Rows(i).Item("Cible")
                                        DataListeIntegrer.Rows(iRow).Cells("FichierExport").Value = OlesumAnaData.Rows(j).Item("Fichier")
                                        DataListeIntegrer.Rows(iRow).Cells("CheminExport").Value = OlesumAnaData.Rows(j).Item("CheminFichier")
                                        DataListeIntegrer.Rows(iRow).Cells("Comptable").Value = OlesumAnaData.Rows(j).Item("Comptable")
                                        DataListeIntegrer.Rows(iRow).Cells("Commercial").Value = OlesumAnaData.Rows(j).Item("Commercial")
                                        DataListeIntegrer.Rows(iRow).Cells("Chemin").Value = OlesumAnaData.Rows(j).Item("PathFormat")
                                        DataListeIntegrer.Rows(iRow).Cells("NomFormat").Value = OlesumAnaData.Rows(j).Item("NomFormat")
                                        DataListeIntegrer.Rows(iRow).Cells("TypeFormat").Value = Afficheauuser(OlesumAnaData.Rows(j).Item("TypeFormat"))
                                        DataListeIntegrer.Rows(iRow).Cells("Mode").Value = OlesumAnaData.Rows(j).Item("Mode")
                                        DataListeIntegrer.Rows(iRow).Cells("FeuilleExcel").Value = OlesumAnaData.Rows(j).Item("FeuilleExcel")
                                        DataListeIntegrer.Rows(iRow).Cells("Deplace").Value = OlesumAnaData.Rows(j).Item("Deplace")
                                        DataListeIntegrer.Rows(iRow).Cells("Dossier").Value = OlesumAnaData.Rows(j).Item("Dossier")
                                        DataListeIntegrer.Rows(iRow).Cells("Valider").Value = True
                                        DataListeIntegrer.Rows(iRow).Cells("Modification").Value = Strings.FormatDateTime(OlesumAnaData.Rows(j).Item("DateModif"), DateFormat.ShortDate)
                                        DataListeIntegrer.Rows(iRow).Cells("Creation").Value = Strings.FormatDateTime(OlesumAnaData.Rows(j).Item("DateCreation"), DateFormat.ShortDate)
                                        iRow = iRow + 1
                                    Next j
                                End If
                            Else
                                Dim OleAdapsumAna As OleDbDataAdapter
                                Dim OlesumAnaDst As DataSet
                                Dim OlesumAnaData As DataTable
                                OleAdapsumAna = New OleDbDataAdapter("select * from TRI_IMPORTFICHIER Where Cible='" & OlesumData.Rows(i).Item("Cible") & "' And  PathFormat='" & Trim(OlesumData.Rows(i).Item("CheminFormat")) & "' And CheminImport='" & Trim(OlesumData.Rows(i).Item("CheminFilexport")) & "' Order by DateCreation ASC", OleConnenection)
                                OlesumAnaDst = New DataSet
                                OleAdapsumAna.Fill(OlesumAnaDst)
                                OlesumAnaData = OlesumAnaDst.Tables(0)
                                If OlesumAnaData.Rows.Count <> 0 Then
                                    For j = 0 To OlesumAnaData.Rows.Count - 1
                                        DataListeIntegrer.RowCount = iRow + 1
                                        DataListeIntegrer.Rows(iRow).Cells("ID").Value = OlesumData.Rows(i).Item("IDDossier")
                                        DataListeIntegrer.Rows(iRow).Cells("Cible").Value = OlesumData.Rows(i).Item("Cible")
                                        DataListeIntegrer.Rows(iRow).Cells("FichierExport").Value = OlesumAnaData.Rows(j).Item("Fichier")
                                        DataListeIntegrer.Rows(iRow).Cells("CheminExport").Value = OlesumAnaData.Rows(j).Item("CheminFichier")
                                        DataListeIntegrer.Rows(iRow).Cells("Comptable").Value = OlesumAnaData.Rows(j).Item("Comptable")
                                        DataListeIntegrer.Rows(iRow).Cells("Commercial").Value = OlesumAnaData.Rows(j).Item("Commercial")
                                        DataListeIntegrer.Rows(iRow).Cells("Chemin").Value = OlesumAnaData.Rows(j).Item("PathFormat")
                                        DataListeIntegrer.Rows(iRow).Cells("NomFormat").Value = OlesumAnaData.Rows(j).Item("NomFormat")
                                        DataListeIntegrer.Rows(iRow).Cells("TypeFormat").Value = Afficheauuser(OlesumAnaData.Rows(j).Item("TypeFormat"))
                                        DataListeIntegrer.Rows(iRow).Cells("Mode").Value = OlesumAnaData.Rows(j).Item("Mode")
                                        DataListeIntegrer.Rows(iRow).Cells("FeuilleExcel").Value = OlesumAnaData.Rows(j).Item("FeuilleExcel")
                                        DataListeIntegrer.Rows(iRow).Cells("Deplace").Value = OlesumAnaData.Rows(j).Item("Deplace")
                                        DataListeIntegrer.Rows(iRow).Cells("Dossier").Value = OlesumAnaData.Rows(j).Item("Dossier")
                                        DataListeIntegrer.Rows(iRow).Cells("Valider").Value = True
                                        DataListeIntegrer.Rows(iRow).Cells("Modification").Value = Strings.FormatDateTime(OlesumAnaData.Rows(j).Item("DateModif"), DateFormat.ShortDate)
                                        DataListeIntegrer.Rows(iRow).Cells("Creation").Value = Strings.FormatDateTime(OlesumAnaData.Rows(j).Item("DateCreation"), DateFormat.ShortDate)
                                        iRow = iRow + 1
                                    Next j
                                End If
                            End If
                        Else
                            If OlesumData.Rows(i).Item("TriModification") = True Then
                                Dim OleAdapsumAna As OleDbDataAdapter
                                Dim OlesumAnaDst As DataSet
                                Dim OlesumAnaData As DataTable
                                OleAdapsumAna = New OleDbDataAdapter("select * from TRI_IMPORTFICHIER Where Cible='" & OlesumData.Rows(i).Item("Cible") & "' And  PathFormat='" & Trim(OlesumData.Rows(i).Item("CheminFormat")) & "' And CheminImport='" & Trim(OlesumData.Rows(i).Item("CheminFilexport")) & "' Order by DateModif ASC", OleConnenection)
                                OlesumAnaDst = New DataSet
                                OleAdapsumAna.Fill(OlesumAnaDst)
                                OlesumAnaData = OlesumAnaDst.Tables(0)
                                If OlesumAnaData.Rows.Count <> 0 Then
                                    For j = 0 To OlesumAnaData.Rows.Count - 1
                                        DataListeIntegrer.RowCount = iRow + 1
                                        DataListeIntegrer.Rows(iRow).Cells("ID").Value = OlesumData.Rows(i).Item("IDDossier")
                                        DataListeIntegrer.Rows(iRow).Cells("Cible").Value = OlesumData.Rows(i).Item("Cible")
                                        DataListeIntegrer.Rows(iRow).Cells("FichierExport").Value = OlesumAnaData.Rows(j).Item("Fichier")
                                        DataListeIntegrer.Rows(iRow).Cells("CheminExport").Value = OlesumAnaData.Rows(j).Item("CheminFichier")
                                        DataListeIntegrer.Rows(iRow).Cells("Comptable").Value = OlesumAnaData.Rows(j).Item("Comptable")
                                        DataListeIntegrer.Rows(iRow).Cells("Commercial").Value = OlesumAnaData.Rows(j).Item("Commercial")
                                        DataListeIntegrer.Rows(iRow).Cells("Chemin").Value = OlesumAnaData.Rows(j).Item("PathFormat")
                                        DataListeIntegrer.Rows(iRow).Cells("NomFormat").Value = OlesumAnaData.Rows(j).Item("NomFormat")
                                        DataListeIntegrer.Rows(iRow).Cells("TypeFormat").Value = Afficheauuser(OlesumAnaData.Rows(j).Item("TypeFormat"))
                                        DataListeIntegrer.Rows(iRow).Cells("Mode").Value = OlesumAnaData.Rows(j).Item("Mode")
                                        DataListeIntegrer.Rows(iRow).Cells("FeuilleExcel").Value = OlesumAnaData.Rows(j).Item("FeuilleExcel")
                                        DataListeIntegrer.Rows(iRow).Cells("Deplace").Value = OlesumAnaData.Rows(j).Item("Deplace")
                                        DataListeIntegrer.Rows(iRow).Cells("Dossier").Value = OlesumAnaData.Rows(j).Item("Dossier")
                                        DataListeIntegrer.Rows(iRow).Cells("Valider").Value = True
                                        DataListeIntegrer.Rows(iRow).Cells("Modification").Value = Strings.FormatDateTime(OlesumAnaData.Rows(j).Item("DateModif"), DateFormat.ShortDate)
                                        DataListeIntegrer.Rows(iRow).Cells("Creation").Value = Strings.FormatDateTime(OlesumAnaData.Rows(j).Item("DateCreation"), DateFormat.ShortDate)
                                        iRow = iRow + 1
                                    Next j
                                End If
                            Else
                                Dim OleAdapsumAna As OleDbDataAdapter
                                Dim OlesumAnaDst As DataSet
                                Dim OlesumAnaData As DataTable
                                OleAdapsumAna = New OleDbDataAdapter("select * from TRI_IMPORTFICHIER Where Cible='" & OlesumData.Rows(i).Item("Cible") & "' And  PathFormat='" & Trim(OlesumData.Rows(i).Item("CheminFormat")) & "' And CheminImport='" & Trim(OlesumData.Rows(i).Item("CheminFilexport")) & "'", OleConnenection)
                                OlesumAnaDst = New DataSet
                                OleAdapsumAna.Fill(OlesumAnaDst)
                                OlesumAnaData = OlesumAnaDst.Tables(0)
                                If OlesumAnaData.Rows.Count <> 0 Then
                                    For j = 0 To OlesumAnaData.Rows.Count - 1
                                        DataListeIntegrer.RowCount = iRow + 1
                                        DataListeIntegrer.Rows(iRow).Cells("ID").Value = OlesumData.Rows(i).Item("IDDossier")
                                        DataListeIntegrer.Rows(iRow).Cells("Cible").Value = OlesumData.Rows(i).Item("Cible")
                                        DataListeIntegrer.Rows(iRow).Cells("FichierExport").Value = OlesumAnaData.Rows(j).Item("Fichier")
                                        DataListeIntegrer.Rows(iRow).Cells("CheminExport").Value = OlesumAnaData.Rows(j).Item("CheminFichier")
                                        DataListeIntegrer.Rows(iRow).Cells("Comptable").Value = OlesumAnaData.Rows(j).Item("Comptable")
                                        DataListeIntegrer.Rows(iRow).Cells("Commercial").Value = OlesumAnaData.Rows(j).Item("Commercial")
                                        DataListeIntegrer.Rows(iRow).Cells("Chemin").Value = OlesumAnaData.Rows(j).Item("PathFormat")
                                        DataListeIntegrer.Rows(iRow).Cells("NomFormat").Value = OlesumAnaData.Rows(j).Item("NomFormat")
                                        DataListeIntegrer.Rows(iRow).Cells("TypeFormat").Value = Afficheauuser(OlesumAnaData.Rows(j).Item("TypeFormat"))
                                        DataListeIntegrer.Rows(iRow).Cells("Mode").Value = OlesumAnaData.Rows(j).Item("Mode")
                                        DataListeIntegrer.Rows(iRow).Cells("FeuilleExcel").Value = OlesumAnaData.Rows(j).Item("FeuilleExcel")
                                        DataListeIntegrer.Rows(iRow).Cells("Deplace").Value = OlesumAnaData.Rows(j).Item("Deplace")
                                        DataListeIntegrer.Rows(iRow).Cells("Dossier").Value = OlesumAnaData.Rows(j).Item("Dossier")
                                        DataListeIntegrer.Rows(iRow).Cells("Valider").Value = True
                                        DataListeIntegrer.Rows(iRow).Cells("Modification").Value = Strings.FormatDateTime(OlesumAnaData.Rows(j).Item("DateModif"), DateFormat.ShortDate)
                                        DataListeIntegrer.Rows(iRow).Cells("Creation").Value = Strings.FormatDateTime(OlesumAnaData.Rows(j).Item("DateCreation"), DateFormat.ShortDate)
                                        iRow = iRow + 1
                                    Next j
                                End If
                            End If
                        End If
                    End If
                End If
            Next i
        End If
    End Sub
    Private Function AffichFormatintegration(ByRef PathFileFormat As String, ByRef Affichetype As String) As Boolean
        Dim NomColonne As String
        Dim NomEntete As String
        Dim PosLeft As Integer
        Dim posGauche As Integer
        Dim Libre As String
        Dim typedeformat As Object = Nothing
        Dim typeImport As Object = Nothing
        Dim ModeFormat As Object = Nothing
        Dim DateFormat As Object = Nothing
        Dim Defaut, Piece, LigneArticle As String
        Dim ValeurDefaut, SageFichier As String
        Dim PieceAuto, Punitaire As Object
        Datagridaffiche.Rows.Clear()
        Datagridaffiche.Columns.Clear()
        If File.Exists(PathFileFormat) = True Then
            Dim FileXml As New XmlTextReader(PathFileFormat)
            Try
                While (FileXml.Read())
                    If FileXml.LocalName = "ColUse" Then
                        NomColonne = FileXml.ReadString
                        FileXml.Read()

                        NomEntete = FileXml.ReadString
                        FileXml.Read()
                        PosLeft = FileXml.ReadString

                        If Trim(Affichetype) = "Excel" Then
                        Else
                            If Trim(Affichetype) = "Longueur Fixe" Then
                                FileXml.Read()
                                posGauche = FileXml.ReadString
                            End If
                        End If
                        FileXml.Read()
                        Libre = FileXml.ReadString

                        FileXml.Read()
                        SageFichier = FileXml.ReadString

                        FileXml.Read()
                        Piece = FileXml.ReadString

                        FileXml.Read()
                        LigneArticle = FileXml.ReadString

                        FileXml.Read()
                        Defaut = FileXml.ReadString

                        FileXml.Read()
                        ValeurDefaut = FileXml.ReadString

                        FileXml.Read()
                        DecFormat = FileXml.ReadString

                        FileXml.Read()
                        typedeformat = FileXml.ReadString

                        FileXml.Read()
                        ModeFormat = FileXml.ReadString

                        FileXml.Read()
                        DateFormat = FileXml.ReadString

                        FileXml.Read()
                        PieceAuto = FileXml.ReadString

                        FileXml.Read()
                        Punitaire = FileXml.ReadString

                        FileXml.Read()
                        typeImport = FileXml.ReadString

                        If Trim(Affichetype) = "Longueur Fixe" Then
                            If (NomColonne <> "" And NomEntete <> "") And (CInt(PosLeft) >= 0 And CInt(posGauche) >= 0) Then
                                CreateComboBoxColumn(Datagridaffiche, NomColonne & "-->(" & PosLeft & ")" & "<-->[" & posGauche & "]" & "->{" & Libre & "-" & SageFichier & "}", NomEntete)
                            End If
                        Else
                            If (NomColonne <> "" And NomEntete <> "") And (CInt(PosLeft) >= 0 And Trim(Libre) <> "") Then
                                CreateComboBoxColumn(Datagridaffiche, NomColonne & "-->(" & PosLeft & ")" & "<-->[" & Libre & "-" & SageFichier & "]", NomEntete)
                            End If
                        End If
                    End If
                End While
                FileXml.Close()
                AffichFormatintegration = True
            Catch ex As Exception
                AffichFormatintegration = False
            End Try
        End If
    End Function
    Private Sub Initialiser()
        ProgresMax = 0
        NbreLigne = 0
        Label8.Text = ""
        NbreTotal = 0
    End Sub
    Private Sub Lecture_Suivant_DuFichierExcel(ByVal sPathFilexporter As String, ByVal spathFileFormat As String, ByRef Formatdefichier As String, ByRef tablefeuille As String, ByRef sColumnsSepar As String)
        Dim OleAdaptater As OleDbDataAdapter
        Dim OleAfficheDataset As DataSet
        Dim Oledatable As DataTable
        Dim jColD As Integer
        Dim iColPosition As Integer
        Dim iColGauchetxt As String
        Dim iLine As Integer
        Dim aRows() As String
        Dim i As Integer, aCols() As String
        iLine = 0
        aRows = Nothing
        If Trim(Formatdefichier) = "Excel" Then
            Try
                If Trim(tablefeuille) <> "" Then
                    If OleExcelConnected.State = ConnectionState.Open Then
                        OleExcelConnected.Close()
                    End If
                    If LoginAuFichierExcel(Trim(sPathFilexporter)) = True Then
                        If AffichFormatintegration(spathFileFormat, Formatdefichier) = True Then
                            ProgressBar1.Value = ProgressBar1.Minimum
                            Datagridaffiche.Rows.Clear()
                            NbreTotal = DecFormat
                            OleAdaptater = New OleDbDataAdapter("select * from [" & tablefeuille & "$] ", OleExcelConnected)
                            OleAfficheDataset = New DataSet
                            OleAdaptater.Fill(OleAfficheDataset)
                            Oledatable = OleAfficheDataset.Tables(0)
                            If Oledatable.Rows.Count <> 0 Then
                                ProgresMax = Oledatable.Rows.Count - DecFormat
                                For i = DecFormat To Oledatable.Rows.Count - 1
                                    Datagridaffiche.RowCount = iLine + 1
                                    For jColD = 0 To Datagridaffiche.ColumnCount - 1
                                        iColPosition = CInt(Strings.Right(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1), Strings.Len(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1)) - InStr(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1), "(")))
                                        If iColPosition <= Oledatable.Columns.Count Then
                                            If iColPosition <> 0 Then
                                                If Convert.IsDBNull(Oledatable.Rows(i).Item(iColPosition - 1)) = False Then
                                                    Datagridaffiche.Item(jColD, iLine).Value = Trim(Oledatable.Rows(i).Item(iColPosition - 1))
                                                Else
                                                    Datagridaffiche.Item(jColD, iLine).Value = ""
                                                End If
                                            Else
                                                Datagridaffiche.Item(jColD, iLine).Value = LireFichierFormat(spathFileFormat, Datagridaffiche.Columns(jColD).Name, Formatdefichier)
                                            End If
                                        Else
                                            Datagridaffiche.Item(jColD, iLine).Value = ""
                                        End If
                                    Next jColD
                                    iLine = iLine + 1
                                    If i >= 500 Then
                                        Exit Sub
                                    End If
                                Next i
                            End If
                        Else
                            Label5.Text = "Impossible d' Integrer le Fichier " & sPathFilexporter & " car Erreur de lecture du Fichier Format " & spathFileFormat
                        End If
                    Else
                        Label5.Text = "Echec de Connexion au fichier Excel :" & sPathFilexporter & " : Echec de traitement"
                    End If
                Else
                    Label5.Text = "Aucune Feuille Excel paramétrée pour le fichier :" & Trim(sPathFilexporter) & " : Echec de traitement"
                End If
            Catch ex As Exception
            End Try
        Else
            If Trim(Formatdefichier) = "Délimité" Or Trim(Formatdefichier) = "Tabulation" Or Trim(Formatdefichier) = "Pipe" Then
                Try
                    If AffichFormatintegration(spathFileFormat, Formatdefichier) = True Then
                        aRows = GetArrayFile(sPathFilexporter, aRows)
                        ProgressBar1.Value = ProgressBar1.Minimum
                        Datagridaffiche.Rows.Clear()
                        ProgresMax = UBound(aRows) + 1 - DecFormat
                        For i = DecFormat To UBound(aRows)
                            aCols = Split(aRows(i), sColumnsSepar)
                            Datagridaffiche.RowCount = iLine + 1
                            For jColD = 0 To Datagridaffiche.ColumnCount - 1
                                iColPosition = CInt(Strings.Right(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1), Strings.Len(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1)) - InStr(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1), "(")))
                                If iColPosition <> 0 Then
                                    If iColPosition <= (UBound(aCols) + 1) Then
                                        Datagridaffiche.Item(jColD, iLine).Value = Trim(aCols(iColPosition - 1))
                                    Else
                                        Datagridaffiche.Item(jColD, iLine).Value = ""
                                    End If
                                Else
                                    Datagridaffiche.Item(jColD, iLine).Value = LireFichierFormat(spathFileFormat, Datagridaffiche.Columns(jColD).Name, Formatdefichier)
                                End If
                            Next jColD
                            iLine = iLine + 1
                            If i >= 500 Then
                                Exit Sub
                            End If
                        Next i
                    Else

                    End If
                Catch ex As Exception
                End Try
            Else
                If Trim(Formatdefichier) = "Longueur Fixe" Then
                    Try
                        If AffichFormatintegration(spathFileFormat, Formatdefichier) = True Then
                            aRows = GetArrayFile(sPathFilexporter, aRows)
                            NbreTotal = DecFormat
                            ProgressBar1.Value = ProgressBar1.Minimum
                            Datagridaffiche.Rows.Clear()
                            ProgresMax = UBound(aRows) + 1 - DecFormat
                            For i = DecFormat To UBound(aRows)
                                Datagridaffiche.RowCount = iLine + 1
                                For jColD = 0 To Datagridaffiche.ColumnCount - 1
                                    iColPosition = CInt(Strings.Left(Strings.Right(Datagridaffiche.Columns(jColD).HeaderText, Strings.Len(Datagridaffiche.Columns(jColD).HeaderText) - InStr(Datagridaffiche.Columns(jColD).HeaderText, "[")), InStr(Strings.Right(Datagridaffiche.Columns(jColD).HeaderText, Strings.Len(Datagridaffiche.Columns(jColD).HeaderText) - InStr(Datagridaffiche.Columns(jColD).HeaderText, "[")), "]") - 1))
                                    iColGauchetxt = CInt(Strings.Right(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1), Strings.Len(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1)) - InStr(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1), "(")))
                                    If iColPosition <> 0 Or iColGauchetxt <> 0 Then
                                        Datagridaffiche.Item(jColD, iLine).Value = Trim(Strings.Mid(aRows(i), iColPosition, iColGauchetxt))
                                    Else
                                        Datagridaffiche.Item(jColD, iLine).Value = LireFichierFormat(spathFileFormat, Datagridaffiche.Columns(jColD).Name, Formatdefichier)
                                    End If
                                Next jColD
                                iLine = iLine + 1
                                If i >= 100 Then
                                    Exit Sub
                                End If
                            Next i
                        Else
                            ErreurJrn.WriteLine("Impossible d' Integrer le Fichier " & sPathFilexporter & " car Erreur de lecture du Fichier Format " & spathFileFormat)
                        End If
                    Catch ex As Exception
                    End Try
                End If
            End If
        End If
    End Sub
    Private Function LireFichierFormat(ByRef ScheminFileFormat As String, ByRef Colname As String, ByRef Lireformatype As String) As Object
        Dim NomColonne As String
        Dim NomEntete As String
        Dim PosLeft As Integer
        Dim poslongueur, LigneArticle As String
        Dim Defaut, Piece, SageFichier As String
        Dim ValeurDefaut, Infolibre As String
        Dim typedeformat As Object = Nothing
        Dim typeImport As Object = Nothing
        Dim ModeFormat As Object = Nothing
        Dim DateFormat As Object = Nothing
        Dim PieceAuto, Punitaire As Object
        Try
            If Trim(ScheminFileFormat) <> "" Then
                If File.Exists(ScheminFileFormat) = True Then
                    Dim FileXml As New XmlTextReader(Trim(ScheminFileFormat))
                    While (FileXml.Read())
                        If FileXml.LocalName = "ColUse" Then
                            NomColonne = FileXml.ReadString

                            FileXml.Read()
                            NomEntete = FileXml.ReadString

                            FileXml.Read()
                            PosLeft = FileXml.ReadString

                            If Trim(Lireformatype) = "Excel" Then
                            Else
                                If Trim(Lireformatype) = "Longueur Fixe" Then
                                    FileXml.Read()
                                    poslongueur = FileXml.ReadString
                                End If
                            End If

                            FileXml.Read()
                            Infolibre = FileXml.ReadString

                            FileXml.Read()
                            SageFichier = FileXml.ReadString

                            FileXml.Read()
                            Piece = FileXml.ReadString

                            FileXml.Read()
                            LigneArticle = FileXml.ReadString

                            FileXml.Read()
                            Defaut = FileXml.ReadString

                            FileXml.Read()
                            ValeurDefaut = FileXml.ReadString

                            FileXml.Read()
                            DecFormat = FileXml.ReadString

                            FileXml.Read()
                            typedeformat = FileXml.ReadString

                            FileXml.Read()
                            ModeFormat = FileXml.ReadString

                            FileXml.Read()
                            DateFormat = FileXml.ReadString

                            FileXml.Read()
                            PieceAuto = FileXml.ReadString

                            FileXml.Read()
                            Punitaire = FileXml.ReadString

                            FileXml.Read()
                            typeImport = FileXml.ReadString

                            If Colname = NomEntete Then
                                Result = ValeurDefaut
                            End If
                        End If
                    End While
                    FileXml.Close()
                End If
            End If
        Catch ex As Exception
        End Try
        Return Result
    End Function
    Private Sub vidage()
        PieceArticle = Nothing
        LigneCodeArticle = Nothing
        EnteteDateDocument = Nothing
        EntetePieceInterne = Nothing
        EnteteReference = Nothing
        EnteteSoucheDocument = Nothing
        EnteteTyPeDocument = Nothing
        LigneDatedeFabrication = Nothing
        LigneDatedeLivraison = Nothing
        LigneDatedePeremption = Nothing
        LigneDesignationArticle = Nothing
        IDDepotEnteteOrigine = Nothing
        LigneNSerieLot = Nothing
        LignePoidsBrut = Nothing
        LignePoidsNet = Nothing
        LignePrixUnitaire = Nothing
        LigneQuantite = Nothing
        LigneReference = Nothing
        EntetePlanAnalytique = Nothing
        EnteteCodeAffaire = Nothing
        EnteteIntituleDepotDestination = Nothing
        IDDepotEnteteDestination = Nothing
        LigneIntituleDepot = Nothing
        IDDepotLigne = Nothing
        PuStock = 0
    End Sub
    Private Sub RecuperationEnregistrement(ByRef sPathFilexporter As String, ByRef spathFileFormat As String, ByRef typeFormat As String, ByRef Base_Excel As String, ByRef sColumnsSepar As String, ByRef ListePiece As List(Of String), ByRef PieceCreation As Object, ByRef PieceAuto As Object)
        Dim OleAdaptater As OleDbDataAdapter
        Dim OleAfficheDataset As DataSet
        Dim Oledatable As DataTable
        Dim jColD As Integer
        Dim aRows(), aCols(), Dataname(), FichierRecup As String
        Dim iColPosition, iColGauchetxt As Integer
        Dim i, j As Integer
        Dim Piecedocument As Object = Nothing
        aRows = Nothing
        Dataname = Split(sPathFilexporter, "\")
        FichierRecup = "Recup_" & Strings.Right(DateAndTime.Year(Now), 2) & "" & Format(DateAndTime.Month(Now), "00") & "" & Format(DateAndTime.Day(Now), "00") & "[" & "" & Format(DateAndTime.Hour(Now), "00") & "-" & Format(DateAndTime.Minute(Now), "00") & "-" & Format(DateAndTime.Second(Now), "00") & "]" & Dataname(UBound(Dataname))
        If ListePiece.Count <> 0 Then
            If Trim(typeFormat) = "Excel" Then
                'Dim Cnapplica As New Microsoft.Office.Interop.Excel.Application
                'Dim Cnbook As Microsoft.Office.Interop.Excel.Workbook
                'Dim Cnsheet As Microsoft.Office.Interop.Excel.Worksheet
                'Try
                '    If AffichFormatintegration(spathFileFormat, typeFormat) = True Then
                '        Datagridaffiche.Rows.Clear()
                '        OleAdaptater = New OleDbDataAdapter("select * from [" & Base_Excel & "$] ", OleExcelConnected)
                '        OleAfficheDataset = New DataSet
                '        OleAdaptater.Fill(OleAfficheDataset)
                '        Oledatable = OleAfficheDataset.Tables(0)
                '        If Oledatable.Rows.Count <> 0 Then
                '            Cnapplica = CreateObject("Excel.Application")
                '            Cnbook = Cnapplica.Workbooks.Add
                '            Cnsheet = Cnbook.Worksheets.Add
                '            For i = Cnbook.Sheets.Count To 1 Step -1
                '                If Cnbook.Sheets(i).name() = Base_Excel Then
                '                    Cnbook.Worksheets(i).Delete()
                '                End If
                '            Next i
                '            Cnsheet.Name = Base_Excel
                '            ProgressBar1.Value = ProgressBar1.Minimum
                '            ProgressBar1.Maximum = Oledatable.Rows.Count - DecFormat
                '            j = 0 + CInt(DecFormat)
                '            For i = DecFormat To Oledatable.Rows.Count - 1
                '                If Datagridaffiche.Columns.Contains(PieceCreation) = True Then
                '                    iColPosition = CInt(Strings.Right(Strings.Left(Datagridaffiche.Columns(PieceCreation).HeaderText, InStr(Datagridaffiche.Columns(PieceCreation).HeaderText, ")") - 1), Strings.Len(Strings.Left(Datagridaffiche.Columns(PieceCreation).HeaderText, InStr(Datagridaffiche.Columns(PieceCreation).HeaderText, ")") - 1)) - InStr(Strings.Left(Datagridaffiche.Columns(PieceCreation).HeaderText, InStr(Datagridaffiche.Columns(PieceCreation).HeaderText, ")") - 1), "(")))
                '                    If iColPosition <= Oledatable.Columns.Count Then
                '                        If iColPosition <> 0 Then
                '                            If Convert.IsDBNull(Oledatable.Rows(i).Item(iColPosition - 1)) = False Then
                '                                If Trim(PieceAuto) = "non" Then
                '                                    If Strings.Len(Trim(Oledatable.Rows(i).Item(iColPosition - 1))) <= 8 Then
                '                                        Piecedocument = Formatage_Chaine(Trim(Oledatable.Rows(i).Item(iColPosition - 1)))
                '                                    Else
                '                                        Piecedocument = Formatage_Chaine(Strings.Left(Trim(Oledatable.Rows(i).Item(iColPosition - 1)), 8))
                '                                    End If
                '                                Else
                '                                    Piecedocument = Trim(Oledatable.Rows(i).Item(iColPosition - 1))
                '                                End If
                '                                If ListePiece.Contains(Trim(Piecedocument)) = True Then
                '                                    j = j + 1
                '                                    For jColD = 0 To Oledatable.Columns.Count - 1
                '                                        If Convert.IsDBNull(Oledatable.Rows(i).Item(jColD)) = False Then
                '                                            Cnsheet.Cells(j, jColD + 1) = Oledatable.Rows(i).Item(jColD)
                '                                        End If
                '                                    Next jColD
                '                                End If
                '                            End If
                '                        Else
                '                            If Trim(PieceAuto) = "non" Then
                '                                If Strings.Len(Trim(LireFichierFormat(spathFileFormat, PieceCreation, typeFormat))) <= 8 Then
                '                                    Piecedocument = Formatage_Chaine(Trim(LireFichierFormat(spathFileFormat, PieceCreation, typeFormat)))
                '                                Else
                '                                    Piecedocument = Formatage_Chaine(Strings.Left(Trim(LireFichierFormat(spathFileFormat, PieceCreation, typeFormat)), 8))
                '                                End If
                '                            Else
                '                                Piecedocument = Trim(LireFichierFormat(spathFileFormat, PieceCreation, typeFormat))
                '                            End If
                '                            If ListePiece.Contains(Trim(Piecedocument)) = True Then
                '                                j = j + 1
                '                                For jColD = 0 To Oledatable.Columns.Count - 1
                '                                    If Convert.IsDBNull(Oledatable.Rows(i).Item(jColD)) = False Then
                '                                        Cnsheet.Cells(j, jColD + 1) = Oledatable.Rows(i).Item(jColD)
                '                                    End If
                '                                Next jColD
                '                            End If
                '                        End If
                '                    End If
                '                End If
                '                ProgressBar1.Value = ProgressBar1.Value + 1
                '            Next i
                '            Cnbook.SaveCopyAs(PathsFileRecuperer & "" & FichierRecup)
                '            Cnapplica.DefaultSaveFormat = Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel5
                '            Cnbook.Close(False) 'Ferme le classeur
                '            Cnapplica.Quit()
                '            Cnbook = Nothing
                '            Cnapplica = Nothing
                '        End If
                '    Else
                '        ErreurJrn.WriteLine("Impossible d'integrer le Fichier " & sPathFilexporter & " car Erreur de lecture du Fichier Format " & spathFileFormat)
                '    End If
                'Catch ex As Exception
                '    exceptionTrouve = True
                'End Try
            Else
                If Trim(typeFormat) = "Délimité" Or Trim(typeFormat) = "Tabulation" Or Trim(typeFormat) = "Pipe" Then
                    Try
                        Error_journal = File.AppendText(PathsFileRecuperer & "" & FichierRecup)
                        If AffichFormatintegration(spathFileFormat, typeFormat) = True Then
                            aRows = GetArrayFile(sPathFilexporter, aRows)
                            Datagridaffiche.Rows.Clear()
                            ProgressBar1.Value = ProgressBar1.Minimum
                            ProgressBar1.Maximum = UBound(aRows) + 1 - DecFormat
                            For i = DecFormat To UBound(aRows)
                                aCols = Split(aRows(i), sColumnsSepar)
                                If Datagridaffiche.Columns.Contains(PieceCreation) = True Then
                                    iColPosition = CInt(Strings.Right(Strings.Left(Datagridaffiche.Columns(PieceCreation).HeaderText, InStr(Datagridaffiche.Columns(PieceCreation).HeaderText, ")") - 1), Strings.Len(Strings.Left(Datagridaffiche.Columns(PieceCreation).HeaderText, InStr(Datagridaffiche.Columns(PieceCreation).HeaderText, ")") - 1)) - InStr(Strings.Left(Datagridaffiche.Columns(PieceCreation).HeaderText, InStr(Datagridaffiche.Columns(PieceCreation).HeaderText, ")") - 1), "(")))
                                    If iColPosition <> 0 Then
                                        If iColPosition <= (UBound(aCols) + 1) Then
                                            If Trim(PieceAuto) = "non" Then
                                                If Strings.Len(Trim(aCols(iColPosition - 1))) <= 8 Then
                                                    Piecedocument = Formatage_Chaine(Trim(aCols(iColPosition - 1)))
                                                Else
                                                    Piecedocument = Formatage_Chaine(Strings.Left(Trim(aCols(iColPosition - 1)), 8))
                                                End If
                                            Else
                                                Piecedocument = Trim(aCols(iColPosition - 1))
                                            End If
                                            If ListePiece.Contains(Trim(Piecedocument)) = True Then
                                                Error_journal.WriteLine(aRows(i))
                                            End If
                                        End If
                                    Else
                                        If Trim(PieceAuto) = "non" Then
                                            If Strings.Len(Trim(LireFichierFormat(spathFileFormat, PieceCreation, typeFormat))) <= 8 Then
                                                Piecedocument = Formatage_Chaine(Trim(LireFichierFormat(spathFileFormat, PieceCreation, typeFormat)))
                                            Else
                                                Piecedocument = Formatage_Chaine(Strings.Left(Trim(LireFichierFormat(spathFileFormat, PieceCreation, typeFormat)), 8))
                                            End If
                                        Else
                                            Piecedocument = Trim(LireFichierFormat(spathFileFormat, PieceCreation, typeFormat))
                                        End If
                                        If ListePiece.Contains(Trim(Piecedocument)) = True Then
                                            Error_journal.WriteLine(aRows(i))
                                        End If
                                    End If
                                End If
                                ProgressBar1.Value = ProgressBar1.Value + 1
                            Next i
                            Error_journal.Close()
                        Else
                            ErreurJrn.WriteLine("Impossible d' Integrer le Fichier " & sPathFilexporter & " car Erreur de lecture du Fichier Format " & spathFileFormat)
                        End If
                    Catch ex As Exception
                        exceptionTrouve = True
                    End Try
                Else
                    If Trim(typeFormat) = "Longueur Fixe" Then
                        Try
                            Error_journal = File.AppendText(PathsFileRecuperer & "" & FichierRecup)
                            If AffichFormatintegration(spathFileFormat, typeFormat) = True Then
                                aRows = GetArrayFile(sPathFilexporter, aRows)
                                Datagridaffiche.Rows.Clear()
                                ProgressBar1.Value = ProgressBar1.Minimum
                                ProgressBar1.Maximum = UBound(aRows) + 1 - DecFormat
                                For i = DecFormat To UBound(aRows)
                                    If Datagridaffiche.Columns.Contains(PieceCreation) = True Then
                                        iColPosition = CInt(Strings.Left(Strings.Right(Datagridaffiche.Columns(PieceCreation).HeaderText, Strings.Len(Datagridaffiche.Columns(PieceCreation).HeaderText) - InStr(Datagridaffiche.Columns(PieceCreation).HeaderText, "[")), InStr(Strings.Right(Datagridaffiche.Columns(PieceCreation).HeaderText, Strings.Len(Datagridaffiche.Columns(PieceCreation).HeaderText) - InStr(Datagridaffiche.Columns(PieceCreation).HeaderText, "[")), "]") - 1))
                                        iColGauchetxt = CInt(Strings.Right(Strings.Left(Datagridaffiche.Columns(PieceCreation).HeaderText, InStr(Datagridaffiche.Columns(PieceCreation).HeaderText, ")") - 1), Strings.Len(Strings.Left(Datagridaffiche.Columns(PieceCreation).HeaderText, InStr(Datagridaffiche.Columns(PieceCreation).HeaderText, ")") - 1)) - InStr(Strings.Left(Datagridaffiche.Columns(PieceCreation).HeaderText, InStr(Datagridaffiche.Columns(PieceCreation).HeaderText, ")") - 1), "(")))
                                        If iColPosition <> 0 Or iColGauchetxt <> 0 Then
                                            If Trim(PieceAuto) = "non" Then
                                                If Strings.Len(Trim(Trim(Strings.Mid(aRows(i), iColPosition, iColGauchetxt)))) <= 8 Then
                                                    Piecedocument = Formatage_Chaine(Trim(Trim(Strings.Mid(aRows(i), iColPosition, iColGauchetxt))))
                                                Else
                                                    Piecedocument = Formatage_Chaine(Strings.Left(Trim(Trim(Strings.Mid(aRows(i), iColPosition, iColGauchetxt))), 8))
                                                End If
                                            Else
                                                Piecedocument = Trim(Trim(Strings.Mid(aRows(i), iColPosition, iColGauchetxt)))
                                            End If
                                            If ListePiece.Contains(Trim(Piecedocument)) = True Then
                                                Error_journal.WriteLine(aRows(i))
                                            End If
                                        Else
                                            If Trim(PieceAuto) = "non" Then
                                                If Strings.Len(Trim(LireFichierFormat(spathFileFormat, PieceCreation, typeFormat))) <= 8 Then
                                                    Piecedocument = Formatage_Chaine(Trim(LireFichierFormat(spathFileFormat, PieceCreation, typeFormat)))
                                                Else
                                                    Piecedocument = Formatage_Chaine(Strings.Left(Trim(LireFichierFormat(spathFileFormat, PieceCreation, typeFormat)), 8))
                                                End If
                                            Else
                                                Piecedocument = Trim(LireFichierFormat(spathFileFormat, PieceCreation, typeFormat))
                                            End If
                                            If ListePiece.Contains(Trim(Piecedocument)) = True Then
                                                Error_journal.WriteLine(aRows(i))
                                            End If
                                        End If
                                    End If
                                    ProgressBar1.Value = ProgressBar1.Value + 1
                                Next i
                                Error_journal.Close()
                            Else
                                ErreurJrn.WriteLine("Impossible d'integrer le Fichier " & sPathFilexporter & " car Erreur de lecture du Fichier Format " & spathFileFormat)
                            End If
                        Catch ex As Exception
                            exceptionTrouve = True
                        End Try
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub Integration_Du_Fichier(ByVal sPathFilexporter As String, ByVal spathFileFormat As String, ByRef Formatype As String, ByRef Base_Excel As String, ByRef sColumnsSepar As String, ByRef FormatdeDatefich As String, ByRef DocumentPiece As Object, ByRef PieceAutomatique As Object, ByRef TypeImport As Object)
        Dim OleAdaptater As OleDbDataAdapter
        Dim OleAfficheDataset As DataSet
        Dim Oledatable As DataTable
        Dim m As Integer
        Dim jColD As Integer
        Dim iLine As Integer
        Dim aRows() As String
        Dim iColPosition, iColGauchetxt As Integer
        Dim i As Integer, aCols() As String
        Initialiser()
        iLine = 0
        aRows = Nothing
        Dim DossierAdap As OleDbDataAdapter
        Dim DossierDs As DataSet
        Dim DossierTab As DataTable
        Dim FormatQte As Integer = 0
        DossierAdap = New OleDbDataAdapter("select * from P_DOSSIERCIAL", OleSocieteConnect)
        DossierDs = New DataSet
        DossierAdap.Fill(DossierDs)
        DossierTab = DossierDs.Tables(0)
        If DossierTab.Rows.Count <> 0 Then
            If InStr(DossierTab.Rows(0).Item("D_FormatQte"), ",") <> 0 Then
                If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ",")), ".") <> 0 Then
                    FormatQte = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ","))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ",")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ","))), ".")))
                Else
                    FormatQte = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ",")))
                End If
            Else
                If InStr(DossierTab.Rows(0).Item("D_FormatQte"), ".") <> 0 Then
                    If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ".")), ",") <> 0 Then
                        FormatQte = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), "."))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ".")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), "."))), ",")))
                    Else
                        FormatQte = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ".")))
                    End If
                End If
            End If
        End If
        If Trim(Formatype) = "Excel" Then
            Try
                If AffichFormatintegration(spathFileFormat, Formatype) = True Then
                    ProgressBar1.Value = ProgressBar1.Minimum
                    Datagridaffiche.Rows.Clear()
                    NbreTotal = DecFormat
                    OleAdaptater = New OleDbDataAdapter("select * from [" & Base_Excel & "$] ", OleExcelConnected)
                    OleAfficheDataset = New DataSet
                    OleAdaptater.Fill(OleAfficheDataset)
                    Oledatable = OleAfficheDataset.Tables(0)
                    If Oledatable.Rows.Count <> 0 Then
                        ProgresMax = Oledatable.Rows.Count - DecFormat
                        m = 0
                        infoListe = New List(Of Integer)
                        infoLigne = New List(Of Integer)
                        For i = DecFormat To Oledatable.Rows.Count - 1
                            Datagridaffiche.RowCount = iLine + 1 - m
                            For jColD = 0 To Datagridaffiche.ColumnCount - 1
                                iColPosition = CInt(Strings.Right(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1), Strings.Len(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1)) - InStr(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1), "(")))
                                If iColPosition <= Oledatable.Columns.Count Then
                                    If iColPosition <> 0 Then
                                        If Convert.IsDBNull(Oledatable.Rows(i).Item(iColPosition - 1)) = False Then
                                            Datagridaffiche.Item(jColD, iLine - m).Value = Trim(Oledatable.Rows(i).Item(iColPosition - 1))
                                        Else
                                            Datagridaffiche.Item(jColD, iLine - m).Value = ""
                                        End If
                                    Else
                                        Datagridaffiche.Item(jColD, iLine - m).Value = LireFichierFormat(spathFileFormat, Datagridaffiche.Columns(jColD).Name, Formatype)
                                    End If
                                Else
                                    Datagridaffiche.Item(jColD, iLine - m).Value = ""
                                End If
                            Next jColD
                            iLine = iLine + 1
                            If i Mod 10 = 0 Then
                                Me.Refresh()
                                Integrer_Ecriture(FormatdeDatefich, DocumentPiece, PieceAutomatique, Trim(Formatype), LirePUDefaut(spathFileFormat, Trim(Formatype)), LireLigneArticle(spathFileFormat, Trim(Formatype)), FormatQte, TypeImport)
                                m = iLine
                            Else
                                If i = (Oledatable.Rows.Count - 1) Then
                                    Me.Refresh()
                                    Integrer_Ecriture(FormatdeDatefich, DocumentPiece, PieceAutomatique, Trim(Formatype), LirePUDefaut(spathFileFormat, Trim(Formatype)), LireLigneArticle(spathFileFormat, Trim(Formatype)), FormatQte, TypeImport)
                                    m = iLine
                                End If
                            End If
                        Next i
                    End If
                Else
                    ErreurJrn.WriteLine("Impossible d' Integrer le Fichier " & sPathFilexporter & " car Erreur de lecture du Fichier Format " & spathFileFormat)
                End If
            Catch ex As Exception
                exceptionTrouve = True
            End Try
        Else
            If Trim(Formatype) = "Délimité" Or Trim(Formatype) = "Tabulation" Or Trim(Formatype) = "Pipe" Then
                Try
                    If AffichFormatintegration(spathFileFormat, Formatype) = True Then
                        aRows = GetArrayFile(sPathFilexporter, aRows)
                        NbreTotal = DecFormat
                        ProgressBar1.Value = ProgressBar1.Minimum
                        Datagridaffiche.Rows.Clear()
                        ProgresMax = UBound(aRows) + 1 - DecFormat
                        m = 0
                        infoListe = New List(Of Integer)
                        infoLigne = New List(Of Integer)
                        For i = DecFormat To UBound(aRows)
                            aCols = Split(aRows(i), sColumnsSepar)
                            Datagridaffiche.RowCount = iLine + 1 - m
                            For jColD = 0 To Datagridaffiche.ColumnCount - 1
                                iColPosition = CInt(Strings.Right(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1), Strings.Len(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1)) - InStr(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1), "(")))
                                If iColPosition <> 0 Then
                                    If iColPosition <= (UBound(aCols) + 1) Then
                                        Datagridaffiche.Item(jColD, iLine - m).Value = Trim(aCols(iColPosition - 1))
                                    Else
                                        Datagridaffiche.Item(jColD, iLine - m).Value = ""
                                    End If
                                Else
                                    Datagridaffiche.Item(jColD, iLine - m).Value = LireFichierFormat(spathFileFormat, Datagridaffiche.Columns(jColD).Name, Formatype)
                                End If
                            Next jColD
                            iLine = iLine + 1
                            If i Mod 10 = 0 Then
                                Me.Refresh()
                                Integrer_Ecriture(FormatdeDatefich, DocumentPiece, PieceAutomatique, Trim(Formatype), LirePUDefaut(spathFileFormat, Trim(Formatype)), LireLigneArticle(spathFileFormat, Trim(Formatype)), FormatQte, TypeImport)
                                m = iLine
                            Else
                                If i = UBound(aRows) Then
                                    Me.Refresh()
                                    Integrer_Ecriture(FormatdeDatefich, DocumentPiece, PieceAutomatique, Trim(Formatype), LirePUDefaut(spathFileFormat, Trim(Formatype)), LireLigneArticle(spathFileFormat, Trim(Formatype)), FormatQte, TypeImport)
                                    m = iLine
                                End If
                            End If
                        Next i
                    Else
                        ErreurJrn.WriteLine("Impossible d' Integrer le Fichier " & sPathFilexporter & " car Erreur de lecture du Fichier Format " & spathFileFormat)
                    End If
                Catch ex As Exception
                    exceptionTrouve = True
                End Try
            Else
                If Trim(Formatype) = "Longueur Fixe" Then
                    Try
                        If AffichFormatintegration(spathFileFormat, Formatype) = True Then
                            aRows = GetArrayFile(sPathFilexporter, aRows)
                            NbreTotal = DecFormat
                            ProgressBar1.Value = ProgressBar1.Minimum
                            Datagridaffiche.Rows.Clear()
                            ProgresMax = UBound(aRows) + 1 - DecFormat
                            m = 0
                            infoListe = New List(Of Integer)
                            infoLigne = New List(Of Integer)
                            For i = DecFormat To UBound(aRows)
                                Datagridaffiche.RowCount = iLine + 1 - m
                                For jColD = 0 To Datagridaffiche.ColumnCount - 1
                                    iColPosition = CInt(Strings.Left(Strings.Right(Datagridaffiche.Columns(jColD).HeaderText, Strings.Len(Datagridaffiche.Columns(jColD).HeaderText) - InStr(Datagridaffiche.Columns(jColD).HeaderText, "[")), InStr(Strings.Right(Datagridaffiche.Columns(jColD).HeaderText, Strings.Len(Datagridaffiche.Columns(jColD).HeaderText) - InStr(Datagridaffiche.Columns(jColD).HeaderText, "[")), "]") - 1))
                                    iColGauchetxt = CInt(Strings.Right(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1), Strings.Len(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1)) - InStr(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1), "(")))
                                    If iColPosition <> 0 Or iColGauchetxt <> 0 Then
                                        Datagridaffiche.Item(jColD, iLine - m).Value = Trim(Strings.Mid(aRows(i), iColPosition, iColGauchetxt))
                                    Else
                                        Datagridaffiche.Item(jColD, iLine - m).Value = LireFichierFormat(spathFileFormat, Datagridaffiche.Columns(jColD).Name, Formatype)
                                    End If
                                Next jColD
                                iLine = iLine + 1
                                If i Mod 10 = 0 Then
                                    Me.Refresh()
                                    Integrer_Ecriture(FormatdeDatefich, DocumentPiece, PieceAutomatique, Trim(Formatype), LirePUDefaut(spathFileFormat, Trim(Formatype)), LireLigneArticle(spathFileFormat, Trim(Formatype)), FormatQte, TypeImport)
                                    m = iLine
                                Else
                                    If i = UBound(aRows) Then
                                        Me.Refresh()
                                        Integrer_Ecriture(FormatdeDatefich, DocumentPiece, PieceAutomatique, Trim(Formatype), LirePUDefaut(spathFileFormat, Trim(Formatype)), LireLigneArticle(spathFileFormat, Trim(Formatype)), FormatQte, TypeImport)
                                        m = iLine
                                    End If
                                End If
                            Next i
                        Else
                            ErreurJrn.WriteLine("Impossible d' Integrer le Fichier " & sPathFilexporter & " car Erreur de lecture du Fichier Format " & spathFileFormat)
                        End If
                    Catch ex As Exception
                        exceptionTrouve = True
                    End Try
                End If
            End If
        End If

    End Sub
    Private Sub Integrer_Ecriture(ByRef FormatfichierDate As String, ByRef DocumentCreationPiece As Object, ByRef PieceAutoma As Object, ByRef FormatIntegrer As Object, ByRef Punitaire As String, ByRef IdentifiantArticle As String, ByRef FormatQteU As Integer, ByRef TypeImport As Object)
        Me.Cursor = Cursors.WaitCursor
        BT_integrer.Enabled = False
        If Datagridaffiche.RowCount >= 0 Then
            ProgressBar1.Maximum = ProgresMax
            Try
                For numLigne = 0 To Datagridaffiche.RowCount - 1
                    vidage()
                    NbreTotal = NbreTotal + 1
                    Label5.Refresh()
                    Label5.Text = "Integration En Cours..."
                    For numColonne = 0 To Datagridaffiche.ColumnCount - 1
                        'Entête Document
                        If Trim(PieceAutoma) = "non" Then
                            If Datagridaffiche.Columns.Contains(DocumentCreationPiece) = True Then
                                If Strings.Len(Trim(Datagridaffiche.Rows(numLigne).Cells(DocumentCreationPiece).Value)) <= 8 Then
                                    EntetePieceInterne = Formatage_Chaine(Trim(Datagridaffiche.Rows(numLigne).Cells(DocumentCreationPiece).Value))
                                Else
                                    EntetePieceInterne = Formatage_Chaine(Strings.Left(Trim(Datagridaffiche.Rows(numLigne).Cells(DocumentCreationPiece).Value), 8))
                                End If
                            End If
                        Else
                            EntetePieceInterne = Trim(Datagridaffiche.Rows(numLigne).Cells(DocumentCreationPiece).Value)
                        End If
                        If Datagridaffiche.Columns.Contains(IdentifiantArticle) = True Then
                            PieceArticle = Trim(Datagridaffiche.Rows(numLigne).Cells(IdentifiantArticle).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "EnteteTyPeDocument" Then
                            EnteteTyPeDocument = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "EnteteCodeAffaire" Then
                            EnteteCodeAffaire = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "EntetePlanAnalytique" Then
                            EntetePlanAnalytique = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If

                        If Datagridaffiche.Columns(numColonne).Name = "EnteteDateDocument" Then
                            EnteteDateDocument = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "EnteteReference" Then
                            If Strings.Len(Trim(Datagridaffiche.Item(numColonne, numLigne).Value)) <= 17 Then
                                EnteteReference = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                            Else
                                EnteteReference = Strings.Left(Trim(Datagridaffiche.Item(numColonne, numLigne).Value), 17)
                            End If
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "EnteteSoucheDocument" Then
                            EnteteSoucheDocument = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "LigneDatedeFabrication" Then
                            LigneDatedeFabrication = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "LigneDatedeLivraison" Then
                            LigneDatedeLivraison = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "LigneDatedePeremption" Then
                            LigneDatedePeremption = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "LigneDesignationArticle" Then
                            If Strings.Len(Trim(Datagridaffiche.Item(numColonne, numLigne).Value)) <= 69 Then
                                LigneDesignationArticle = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                            Else
                                LigneDesignationArticle = Strings.Left(Trim(Datagridaffiche.Item(numColonne, numLigne).Value), 69)
                            End If
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "EnteteIntituleDepotOrigine" Then
                            EnteteIntituleDepotOrigine = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If

                        If Datagridaffiche.Columns(numColonne).Name = "EnteteIntituleDepotDestination" Then
                            EnteteIntituleDepotDestination = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "IDDepotLigne" Then
                            IDDepotLigne = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "LigneIntituleDepot" Then
                            LigneIntituleDepot = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "IDDepotEnteteOrigine" Then
                            IDDepotEnteteOrigine = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "IDDepotEnteteDestination" Then
                            IDDepotEnteteDestination = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "LigneNSerieLot" Then
                            If Strings.Len(Trim(Datagridaffiche.Item(numColonne, numLigne).Value)) <= 30 Then
                                LigneNSerieLot = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                            Else
                                LigneNSerieLot = Strings.Left(Trim(Datagridaffiche.Item(numColonne, numLigne).Value), 30)
                            End If
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "LignePoidsBrut" Then
                            LignePoidsBrut = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "LignePoidsNet" Then
                            LignePoidsNet = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "LignePrixUnitaire" Then
                            LignePrixUnitaire = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns.Contains("LigneQuantite") = True Then
                            LigneQuantite = Trim(Datagridaffiche.Rows(numLigne).Cells("LigneQuantite").Value)
                        End If

                        If Datagridaffiche.Columns(numColonne).Name = "LigneReference" Then
                            If Strings.Len(Trim(Datagridaffiche.Item(numColonne, numLigne).Value)) <= 17 Then
                                LigneReference = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                            Else
                                LigneReference = Strings.Left(Trim(Datagridaffiche.Item(numColonne, numLigne).Value), 17)
                            End If
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "LigneCodeArticle" Then
                            LigneCodeArticle = Formatage_Article(Trim(Datagridaffiche.Item(numColonne, numLigne).Value))
                        End If

                        'RECHERCHE DE L'INTITULE DE L'INFO LIBRE
                        If Trim(FormatIntegrer) = "Longueur Fixe" Then
                            Dim InfoTableau() As String = Split(Trim(Strings.Left(Strings.Right(Datagridaffiche.Columns(numColonne).HeaderText, Strings.Len(Datagridaffiche.Columns(numColonne).HeaderText) - InStr(Datagridaffiche.Columns(numColonne).HeaderText, "{")), Strings.Len(Strings.Right(Datagridaffiche.Columns(numColonne).HeaderText, Strings.Len(Datagridaffiche.Columns(numColonne).HeaderText) - InStr(Datagridaffiche.Columns(numColonne).HeaderText, "{"))) - 1)), "-")
                            If Trim(InfoTableau(0)) = "oui" Then
                                If Trim(InfoTableau(1)) = "F_DOCENTETE" Then
                                    infoListe.Add(numColonne)
                                End If
                                If Trim(InfoTableau(1)) = "F_DOCLIGNE" Then
                                    infoLigne.Add(numColonne)
                                End If
                            End If
                        Else
                            Dim InfoTableau() As String = Split(Trim(Strings.Left(Strings.Right(Datagridaffiche.Columns(numColonne).HeaderText, Strings.Len(Datagridaffiche.Columns(numColonne).HeaderText) - InStr(Datagridaffiche.Columns(numColonne).HeaderText, "[")), Strings.Len(Strings.Right(Datagridaffiche.Columns(numColonne).HeaderText, Strings.Len(Datagridaffiche.Columns(numColonne).HeaderText) - InStr(Datagridaffiche.Columns(numColonne).HeaderText, "["))) - 1)), "-")
                            If Trim(InfoTableau(0)) = "oui" Then
                                If Trim(InfoTableau(1)) = "F_DOCENTETE" Then
                                    infoListe.Add(numColonne)
                                End If
                                If Trim(InfoTableau(1)) = "F_DOCLIGNE" Then
                                    infoLigne.Add(numColonne)
                                End If
                            End If
                        End If
                    Next numColonne
                    'Creation Effective du Document Commercial
                    EnteteTyPeDocument = TranscodageTypedocument(EnteteTyPeDocument)
                    If CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQteU, DecimalNomb, DecimalMone)) <> 0 Then
                        If Trim(EnteteTyPeDocument) = "23" Then 'Transfert
                            CreationTouslesTiers(EnteteIntituleDepotOrigine, EntetePieceInterne, EnteteTyPeDocument, Document, infoListe, FormatfichierDate, DocumentCreationPiece, PieceAutoma, PieceArticle, Punitaire, IdentifiantArticle, TypeImport)
                        Else
                            exceptionTrouve = True
                            ErreurJrn.WriteLine("Le type de document ne correspond à aucune de ces valeurs (23:Transfert)")
                        End If
                    End If
                    NbreLigne = NbreLigne + 1
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label8.Text = NbreLigne & "/" & ProgresMax
                    Label8.Refresh()
                Next numLigne
            Catch ex As Exception
            End Try
        End If
        Datagridaffiche.Rows.Clear()
        Me.Cursor = Cursors.Default
        BT_integrer.Enabled = True
    End Sub
    Private Sub Creation_Entete_Document(ByRef typedoc As String, ByRef FormatDatefichier As String, ByRef CreationPieceDocument As Object, ByRef PieceInterne As Object, ByRef PieceAutomatique As Object)
        Dim OleRecherAdapter As OleDbDataAdapter
        Dim OleRecherDataset As DataSet
        Dim OleRechDatable As DataTable
        Dim fournisseurAdap As OleDbDataAdapter
        Dim fournisseurDs As DataSet
        Dim fournisseurTab As DataTable
        Dim statistAdap As OleDbDataAdapter
        Dim statistDs As DataSet
        Dim statistTab As DataTable
        Dim OleAdaptaterDelete As OleDbDataAdapter
        Dim OleDeleteDataset As DataSet
        Dim OledatableDelete As DataTable
        Dim DossierAdap As OleDbDataAdapter
        Dim DossierDs As DataSet
        Dim DossierTab As DataTable
        Dim FormatQte As Integer = 0
        Dim FormatMnt As Integer = 0
        DossierAdap = New OleDbDataAdapter("select * from P_DOSSIERCIAL", OleSocieteConnect)
        DossierDs = New DataSet
        DossierAdap.Fill(DossierDs)
        DossierTab = DossierDs.Tables(0)
        If DossierTab.Rows.Count <> 0 Then
            If InStr(DossierTab.Rows(0).Item("D_FormatQte"), ",") <> 0 Then
                If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ",")), ".") <> 0 Then
                    FormatQte = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ","))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ",")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ","))), ".")))
                Else
                    FormatQte = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ",")))
                End If
            Else
                If InStr(DossierTab.Rows(0).Item("D_FormatQte"), ".") <> 0 Then
                    If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ".")), ",") <> 0 Then
                        FormatQte = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), "."))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ".")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), "."))), ",")))
                    Else
                        FormatQte = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ".")))
                    End If
                End If
            End If

            If InStr(DossierTab.Rows(0).Item("D_FormatPrix"), ",") <> 0 Then
                If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ",")), ".") <> 0 Then
                    FormatMnt = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ","))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ",")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ","))), ".")))
                Else
                    FormatMnt = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ",")))
                End If
            Else
                If InStr(DossierTab.Rows(0).Item("D_FormatPrix"), ".") <> 0 Then
                    If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ".")), ",") <> 0 Then
                        FormatMnt = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), "."))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ".")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), "."))), ",")))
                    Else
                        FormatMnt = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ".")))
                    End If
                End If
            End If
        End If
        With Document
            If Datagridaffiche.Columns.Contains("EntetePlanAnalytique") = True Then
                If Datagridaffiche.Columns.Contains("EnteteCodeAffaire") = True Then
                    If BaseCpta.FactoryAnalytique.ExistIntitule(Trim(EntetePlanAnalytique)) = True Then
                        PlanAna = BaseCpta.FactoryAnalytique.ReadIntitule(Trim(EntetePlanAnalytique))
                        If BaseCpta.FactoryCompteA.ExistNumero(PlanAna, Trim(EnteteCodeAffaire)) = True Then
                            .CompteA = BaseCpta.FactoryCompteA.ReadNumero(PlanAna, Trim(EnteteCodeAffaire))
                        Else
                            statistAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Section Analytique>' and Valeurlue ='" & Join(Split(Trim(EnteteCodeAffaire), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                            statistDs = New DataSet
                            statistAdap.Fill(statistDs)
                            statistTab = statistDs.Tables(0)
                            If statistTab.Rows.Count <> 0 Then
                                If BaseCpta.FactoryCompteA.ExistNumero(PlanAna, Trim(statistTab.Rows(0).Item("Correspond"))) = True Then
                                    .CompteA = BaseCpta.FactoryCompteA.ReadNumero(PlanAna, Trim(statistTab.Rows(0).Item("Correspond")))
                                End If
                            End If
                        End If
                    Else
                        fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Plan Analytique>' and Valeurlue ='" & Join(Split(Trim(EntetePlanAnalytique), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                        fournisseurDs = New DataSet
                        fournisseurAdap.Fill(fournisseurDs)
                        fournisseurTab = fournisseurDs.Tables(0)
                        If fournisseurTab.Rows.Count > 0 Then
                            If BaseCpta.FactoryAnalytique.ExistIntitule(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                                PlanAna = BaseCpta.FactoryAnalytique.ReadIntitule(Trim(fournisseurTab.Rows(0).Item("Correspond")))
                                If BaseCpta.FactoryCompteA.ExistNumero(PlanAna, Trim(EnteteCodeAffaire)) = True Then
                                    .CompteA = BaseCpta.FactoryCompteA.ReadNumero(PlanAna, Trim(EnteteCodeAffaire))
                                Else
                                    statistAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Section Analytique>' and Valeurlue ='" & Join(Split(Trim(EnteteCodeAffaire), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                    statistDs = New DataSet
                                    statistAdap.Fill(statistDs)
                                    statistTab = statistDs.Tables(0)
                                    If statistTab.Rows.Count <> 0 Then
                                        If BaseCpta.FactoryCompteA.ExistNumero(PlanAna, Trim(statistTab.Rows(0).Item("Correspond"))) = True Then
                                            .CompteA = BaseCpta.FactoryCompteA.ReadNumero(PlanAna, Trim(statistTab.Rows(0).Item("Correspond")))
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            If Datagridaffiche.Columns.Contains("IDDepotEnteteOrigine") = True Then
                If IsNumeric(Trim(IDDepotEnteteOrigine)) = True Then
                    statistAdap = New OleDbDataAdapter("select * from F_DEPOT where DE_No ='" & CInt(Trim(IDDepotEnteteOrigine)) & "'", OleSocieteConnect)
                    statistDs = New DataSet
                    statistAdap.Fill(statistDs)
                    statistTab = statistDs.Tables(0)
                    If statistTab.Rows.Count <> 0 Then
                        .DepotOrigine = BaseCial.FactoryDepot.ReadIntitule(Trim(statistTab.Rows(0).Item("DE_Intitule")))
                    Else
                        fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Depot Stockage>' and Valeurlue ='" & Join(Split(Trim(IDDepotEnteteOrigine), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                        fournisseurDs = New DataSet
                        fournisseurAdap.Fill(fournisseurDs)
                        fournisseurTab = fournisseurDs.Tables(0)
                        If fournisseurTab.Rows.Count > 0 Then
                            If IsNumeric(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                                statistAdap = New OleDbDataAdapter("select * from F_DEPOT where DE_No ='" & CInt(Trim(fournisseurTab.Rows(0).Item("Correspond"))) & "'", OleSocieteConnect)
                                statistDs = New DataSet
                                statistAdap.Fill(statistDs)
                                statistTab = statistDs.Tables(0)
                                If statistTab.Rows.Count <> 0 Then
                                    .DepotOrigine = BaseCial.FactoryDepot.ReadIntitule(Trim(statistTab.Rows(0).Item("DE_Intitule")))
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            ''''''
            If Datagridaffiche.Columns.Contains("IDDepotEnteteDestination") = True Then
                If IsNumeric(Trim(IDDepotEnteteDestination)) = True Then
                    statistAdap = New OleDbDataAdapter("select * from F_DEPOT where DE_No ='" & CInt(Trim(IDDepotEnteteDestination)) & "'", OleSocieteConnect)
                    statistDs = New DataSet
                    statistAdap.Fill(statistDs)
                    statistTab = statistDs.Tables(0)
                    If statistTab.Rows.Count <> 0 Then
                        .DepotDestination = BaseCial.FactoryDepot.ReadIntitule(Trim(statistTab.Rows(0).Item("DE_Intitule")))
                    Else
                        fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Depot Stockage>' and Valeurlue ='" & Join(Split(Trim(IDDepotEnteteDestination), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                        fournisseurDs = New DataSet
                        fournisseurAdap.Fill(fournisseurDs)
                        fournisseurTab = fournisseurDs.Tables(0)
                        If fournisseurTab.Rows.Count > 0 Then
                            If IsNumeric(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                                statistAdap = New OleDbDataAdapter("select * from F_DEPOT where DE_No ='" & CInt(Trim(fournisseurTab.Rows(0).Item("Correspond"))) & "'", OleSocieteConnect)
                                statistDs = New DataSet
                                statistAdap.Fill(statistDs)
                                statistTab = statistDs.Tables(0)
                                If statistTab.Rows.Count <> 0 Then
                                    .DepotDestination = BaseCial.FactoryDepot.ReadIntitule(Trim(statistTab.Rows(0).Item("DE_Intitule")))
                                End If
                            End If
                        End If
                    End If
                End If
            End If

            If Datagridaffiche.Columns.Contains("EnteteIntituleDepotOrigine") = True Then
                If BaseCial.FactoryDepot.ExistIntitule(Trim(EnteteIntituleDepotOrigine)) = True Then
                    .DepotOrigine = BaseCial.FactoryDepot.ReadIntitule(Trim(EnteteIntituleDepotOrigine))
                Else
                    fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Depot Stockage>' and Valeurlue ='" & Join(Split(Trim(EnteteIntituleDepotOrigine), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                    fournisseurDs = New DataSet
                    fournisseurAdap.Fill(fournisseurDs)
                    fournisseurTab = fournisseurDs.Tables(0)
                    If fournisseurTab.Rows.Count > 0 Then
                        If BaseCial.FactoryDepot.ExistIntitule(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                            .DepotOrigine = BaseCial.FactoryDepot.ReadIntitule(Trim(fournisseurTab.Rows(0).Item("Correspond")))
                        End If
                    End If
                End If
            End If

            ''''''
            If Datagridaffiche.Columns.Contains("EnteteIntituleDepotDestination") = True Then
                If BaseCial.FactoryDepot.ExistIntitule(Trim(EnteteIntituleDepotDestination)) = True Then
                    .DepotDestination = BaseCial.FactoryDepot.ReadIntitule(Trim(EnteteIntituleDepotDestination))
                Else
                    fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Depot Stockage>' and Valeurlue ='" & Join(Split(Trim(EnteteIntituleDepotDestination), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                    fournisseurDs = New DataSet
                    fournisseurAdap.Fill(fournisseurDs)
                    fournisseurTab = fournisseurDs.Tables(0)
                    If fournisseurTab.Rows.Count > 0 Then
                        If BaseCial.FactoryDepot.ExistIntitule(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                            .DepotDestination = BaseCial.FactoryDepot.ReadIntitule(Trim(fournisseurTab.Rows(0).Item("Correspond")))
                        End If
                    End If
                End If
            End If

            If Datagridaffiche.Columns.Contains("EnteteDateDocument") = True Then
                If Trim(EnteteDateDocument) <> "" Then
                    .DO_Date = RenvoieDateValide(Trim(EnteteDateDocument), FormatDatefichier)
                End If
            End If
            If Trim(PieceAutomatique) = "non" Then
                .DO_Piece = PieceInterne
            Else
                If Datagridaffiche.Columns.Contains("EnteteSoucheDocument") = True Then
                    If Trim(EnteteSoucheDocument) <> "" Then
                        If EstNumeric(Trim(EnteteSoucheDocument), DecimalNomb, DecimalMone) = True Then

                        Else
                            If BaseCial.FactorySoucheStock.ExistIntitule(Trim(EnteteSoucheDocument)) = True Then
                                If BaseCial.FactorySoucheStock.ReadIntitule(Trim(EnteteSoucheDocument)).IsValide = True Then
                                    If typedoc = "23" Then
                                        .Souche = BaseCial.FactorySoucheStock.ReadIntitule(Trim(EnteteSoucheDocument))
                                        .DO_Piece = BaseCial.FactorySoucheStock.ReadIntitule(Trim(EnteteSoucheDocument)).ReadCurrentPiece(DocumentType.DocumentTypeStockVirement, DocumentProvenanceType.DocProvenanceNormale)
                                    End If
                                End If
                            Else
                                fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Intitule Souche>' and Valeurlue ='" & Join(Split(Trim(EnteteSoucheDocument), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                fournisseurDs = New DataSet
                                fournisseurAdap.Fill(fournisseurDs)
                                fournisseurTab = fournisseurDs.Tables(0)
                                If fournisseurTab.Rows.Count <> 0 Then
                                    If BaseCial.FactorySoucheVente.ReadIntitule(Trim(fournisseurTab.Rows(0).Item("Correspond"))).IsValide = True Then
                                        If typedoc = "23" Then
                                            .Souche = BaseCial.FactorySoucheStock.ReadIntitule(Trim(fournisseurTab.Rows(0).Item("Correspond")))
                                            .DO_Piece = BaseCial.FactorySoucheStock.ReadIntitule(Trim(fournisseurTab.Rows(0).Item("Correspond"))).ReadCurrentPiece(DocumentType.DocumentTypeStockVirement, DocumentProvenanceType.DocProvenanceNormale)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            If Datagridaffiche.Columns.Contains("EnteteReference") = True Then
                .DO_Ref = EnteteReference
            End If
            .Write()
            ErreurJrn.WriteLine("-----------------------------------------------------------------------------------------------------")
            ErreurJrn.WriteLine("")

            If typedoc = "23" Then
                ErreurJrn.WriteLine("Mouvement de Transfert N° : " & Trim(Document.DO_Piece) & " Créé Pour la pièce N° :" & Trim(EntetePieceInterne))
            End If
            'Traitement des Infos Libres
            Try
                If infoListe.Count > 0 Then
                    While infoListe.Count <> 0
                        OleAdaptaterDelete = New OleDbDataAdapter("select * From WIT_COL where Libelle='" & Trim(Datagridaffiche.Columns(infoListe.Item(0)).Name) & "' And Libre=True", OleConnenection)
                        OleDeleteDataset = New DataSet
                        OleAdaptaterDelete.Fill(OleDeleteDataset)
                        OledatableDelete = OleDeleteDataset.Tables(0)
                        If OledatableDelete.Rows.Count <> 0 Then
                            'L'info Libre Parametrée par l'utilisateur existe dans Sage
                            If Document.InfoLibre.Count <> 0 Then
                                If IsNothing(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) = False Then
                                    If Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) <> "" Then
                                        statistAdap = New OleDbDataAdapter("select * from cbSysLibre where CB_File='F_DOCENTETE' and CB_Name ='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "'", OleSocieteConnect)
                                        statistDs = New DataSet
                                        statistAdap.Fill(statistDs)
                                        statistTab = statistDs.Tables(0)
                                        If statistTab.Rows.Count <> 0 Then
                                            'Texte
                                            If statistTab.Rows(0).Item("CB_Type") = 9 Then
                                                OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                                OleRecherDataset = New DataSet
                                                OleRecherAdapter.Fill(OleRecherDataset)
                                                OleRechDatable = OleRecherDataset.Tables(0)
                                                If OleRechDatable.Rows.Count <> 0 Then
                                                    If statistTab.Rows(0).Item("CB_Len") >= Strings.Len(Trim(OleRechDatable.Rows(0).Item("Correspond"))) Then
                                                        Document.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = Trim(OleRechDatable.Rows(0).Item("Correspond"))
                                                        Document.Write()
                                                    End If
                                                Else
                                                    If statistTab.Rows(0).Item("CB_Len") >= Strings.Len(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value)) Then
                                                        Document.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value)
                                                        Document.Write()
                                                    End If
                                                End If
                                            End If
                                            'Table
                                            If statistTab.Rows(0).Item("CB_Type") = 22 Then
                                                OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                                OleRecherDataset = New DataSet
                                                OleRecherAdapter.Fill(OleRecherDataset)
                                                OleRechDatable = OleRecherDataset.Tables(0)
                                                If OleRechDatable.Rows.Count <> 0 Then
                                                    If statistTab.Rows(0).Item("CB_Len") > Strings.Len(Trim(OleRechDatable.Rows(0).Item("Correspond"))) Then
                                                        Document.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = Trim(OleRechDatable.Rows(0).Item("Correspond"))
                                                        Document.Write()
                                                    End If
                                                Else
                                                    If statistTab.Rows(0).Item("CB_Len") > Strings.Len(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value)) Then
                                                        Document.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value)
                                                        Document.Write()
                                                    End If
                                                End If
                                            End If
                                            'Montant
                                            If statistTab.Rows(0).Item("CB_Type") = 20 Then
                                                If Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) <> "" Then
                                                    OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                                    OleRecherDataset = New DataSet
                                                    OleRecherAdapter.Fill(OleRecherDataset)
                                                    OleRechDatable = OleRecherDataset.Tables(0)
                                                    If OleRechDatable.Rows.Count <> 0 Then
                                                        If EstNumeric(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone) = True Then
                                                            Document.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone))
                                                            Document.Write()
                                                        End If
                                                    Else
                                                        If EstNumeric(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), DecimalNomb, DecimalMone) = True Then
                                                            Document.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), DecimalNomb, DecimalMone))
                                                            Document.Write()
                                                        End If
                                                    End If
                                                End If
                                            End If
                                            'Valeur
                                            If statistTab.Rows(0).Item("CB_Type") = 7 Then
                                                If Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) <> "" Then
                                                    OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                                    OleRecherDataset = New DataSet
                                                    OleRecherAdapter.Fill(OleRecherDataset)
                                                    OleRechDatable = OleRecherDataset.Tables(0)
                                                    If OleRechDatable.Rows.Count <> 0 Then
                                                        If EstNumeric(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone) = True Then
                                                            Document.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone))
                                                            Document.Write()
                                                        End If
                                                    Else
                                                        If EstNumeric(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), DecimalNomb, DecimalMone) = True Then
                                                            Document.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), DecimalNomb, DecimalMone))
                                                            Document.Write()
                                                        End If
                                                    End If
                                                End If
                                            End If

                                            'Date Court
                                            If statistTab.Rows(0).Item("CB_Type") = 3 Then
                                                If Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) <> "" Then
                                                    OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                                    OleRecherDataset = New DataSet
                                                    OleRecherAdapter.Fill(OleRecherDataset)
                                                    OleRechDatable = OleRecherDataset.Tables(0)
                                                    If OleRechDatable.Rows.Count <> 0 Then
                                                        If Verificatdate(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier, Datagridaffiche.Columns(infoListe.Item(0)).Name) = True Then
                                                            Document.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = RenvoieDateValide(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier)
                                                            Document.Write()
                                                        End If
                                                    Else
                                                        If Verificatdate(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), FormatDatefichier, Datagridaffiche.Columns(infoListe.Item(0)).Name) = True Then
                                                            Document.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = RenvoieDateValide(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), FormatDatefichier)
                                                            Document.Write()
                                                        End If
                                                    End If
                                                End If
                                            End If
                                            'Date Longue
                                            If statistTab.Rows(0).Item("CB_Type") = 14 Then
                                                If Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) <> "" Then
                                                    OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                                    OleRecherDataset = New DataSet
                                                    OleRecherAdapter.Fill(OleRecherDataset)
                                                    OleRechDatable = OleRecherDataset.Tables(0)
                                                    If OleRechDatable.Rows.Count <> 0 Then
                                                        If Verificatdate(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier, Datagridaffiche.Columns(infoListe.Item(0)).Name) = True Then
                                                            Document.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = RenvoieDateValide(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier)
                                                            Document.Write()
                                                        End If
                                                    Else
                                                        If Verificatdate(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), FormatDatefichier, Datagridaffiche.Columns(infoListe.Item(0)).Name) = True Then
                                                            Document.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = RenvoieDateValide(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), FormatDatefichier)
                                                            Document.Write()
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Else
                                    'nothing
                                End If
                            End If
                        End If
                        'L'info Libre Parametrée par l'utilisateur n'existe pas dans Sage
                        infoListe.RemoveAt(0)
                    End While
                End If
            Catch ex As Exception
                exceptionTrouve = True
                If typedoc = "23" Then
                    ErreurJrn.WriteLine("Mouvement de Transfert N° : " & Trim(Document.DO_Piece) & " Erreur de Modification de L'information Libre , vérifiez la longueur de la chaine avec celle paramétrée/ ou la cohérence des données " & ex.Message)
                End If
            End Try
        End With
    End Sub
    Private Sub RenvoiErr_Infolibre(ByRef DocPiece As Object, ByRef typDocumen As Object, ByRef TypeChamp As String, ByRef Nomchamp As String, ByRef valeurChamp As Object, ByRef Messagesystemm As String)
        If typDocumen = "23" Then
            ErreurJrn.WriteLine("Mouvement de Transfert N° : " & Trim(DocPiece) & " Impossible de traiter l'information libre de type " & TypeChamp & " :" & Nomchamp & "  De  valeur entrée '" & Trim(valeurChamp) & " dans Sage. Message système " & Messagesystemm)
        End If
    End Sub
    Private Sub Creation_Ligne_Article(ByRef FormatDatefichier As String, ByRef PieceArticle As String, ByRef Punitaire As String, ByRef IdentifiantArticle As String, ByRef TypeImport As Object)
        Dim fournisseurAdap As OleDbDataAdapter
        Dim fournisseurDs As DataSet
        Dim fournisseurTab As DataTable
        Dim statistAdap As OleDbDataAdapter
        Dim statistDs As DataSet
        Dim statistTab As DataTable
        Dim OleAdaptaterDelete As OleDbDataAdapter
        Dim OleDeleteDataset As DataSet
        Dim OledatableDelete As DataTable
        Dim OleRecherAdapter As OleDbDataAdapter
        Dim OleRecherDataset As DataSet
        Dim OleRechDatable As DataTable
        Dim DossierAdap As OleDbDataAdapter
        Dim DossierDs As DataSet
        Dim DossierTab As DataTable
        Dim FormatQte As Integer = 0
        Dim FormatMnt As Integer = 0
        Dim ExisteQuantite As Boolean = True
        DossierAdap = New OleDbDataAdapter("select * from P_DOSSIERCIAL", OleSocieteConnect)
        DossierDs = New DataSet
        DossierAdap.Fill(DossierDs)
        DossierTab = DossierDs.Tables(0)
        If DossierTab.Rows.Count <> 0 Then
            If InStr(DossierTab.Rows(0).Item("D_FormatQte"), ",") <> 0 Then
                If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ",")), ".") <> 0 Then
                    FormatQte = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ","))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ",")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ","))), ".")))
                Else
                    FormatQte = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ",")))
                End If
            Else
                If InStr(DossierTab.Rows(0).Item("D_FormatQte"), ".") <> 0 Then
                    If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ".")), ",") <> 0 Then
                        FormatQte = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), "."))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ".")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), "."))), ",")))
                    Else
                        FormatQte = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ".")))
                    End If
                End If
            End If

            If InStr(DossierTab.Rows(0).Item("D_FormatPrix"), ",") <> 0 Then
                If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ",")), ".") <> 0 Then
                    FormatMnt = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ","))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ",")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ","))), ".")))
                Else
                    FormatMnt = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ",")))
                End If
            Else
                If InStr(DossierTab.Rows(0).Item("D_FormatPrix"), ".") <> 0 Then
                    If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ".")), ",") <> 0 Then
                        FormatMnt = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), "."))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ".")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), "."))), ",")))
                    Else
                        FormatMnt = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ".")))
                    End If
                End If
            End If
        End If
        Try
            LigneDocument = Document.FactoryDocumentLigne.Create
            With LigneDocument
                If Datagridaffiche.Columns.Contains("LigneDesignationArticle") = True Then
                    .DL_Design = LigneDesignationArticle
                End If
                If Datagridaffiche.Columns.Contains("LignePoidsNet") = True Then
                    If EstNumeric(Trim(LignePoidsNet), DecimalNomb, DecimalMone) = True Then
                        .DL_PoidsNet = CDbl(RenvoiMontant(Trim(LignePoidsNet), FormatQte, DecimalNomb, DecimalMone))
                    End If
                End If
                .Valorisee = True
                If Datagridaffiche.Columns.Contains("LignePoidsBrut") = True Then
                    If EstNumeric(Trim(LignePoidsBrut), DecimalNomb, DecimalMone) = True Then
                        .DL_PoidsBrut = CDbl(RenvoiMontant(Trim(LignePoidsBrut), FormatQte, DecimalNomb, DecimalMone))
                    End If
                End If
                If Datagridaffiche.Columns.Contains("LigneReference") = True Then
                    .DO_Ref = Trim(LigneReference)
                End If
                If Datagridaffiche.Columns.Contains("LigneDatedeFabrication") = True Then
                    If Trim(LigneDatedeFabrication) <> "" Then
                        .LS_Fabrication = RenvoieDateValide(Trim(LigneDatedeFabrication), FormatDatefichier)
                    End If
                End If
                If Datagridaffiche.Columns.Contains("LigneDatedePeremption") = True Then
                    If Trim(LigneDatedePeremption) <> "" Then
                        .LS_Peremption = RenvoieDateValide(Trim(LigneDatedePeremption), FormatDatefichier)
                    End If
                End If

                If Datagridaffiche.Columns.Contains("LigneNSerieLot") = True Then
                    .LS_NoSerie = LigneNSerieLot
                End If
                If Datagridaffiche.Columns.Contains("LigneCodeArticle") = True Then
                    If BaseCial.FactoryArticle.ExistReference(Trim(LigneCodeArticle)) = True Then
                        If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                            .SetDefaultArticle(BaseCial.FactoryArticle.ReadReference(Trim(LigneCodeArticle)), CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)))
                        End If
                    Else
                        fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<ArticleStock>' and Valeurlue ='" & Join(Split(Trim(LigneCodeArticle), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                        fournisseurDs = New DataSet
                        fournisseurAdap.Fill(fournisseurDs)
                        fournisseurTab = fournisseurDs.Tables(0)
                        If fournisseurTab.Rows.Count <> 0 Then
                            If BaseCial.FactoryArticle.ExistReference(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                                If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                                    .SetDefaultArticle(BaseCial.FactoryArticle.ReadReference(Trim(fournisseurTab.Rows(0).Item("Correspond"))), CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)))
                                End If
                            End If
                        End If
                    End If
                End If
                If Datagridaffiche.Columns.Contains("IDDepotEnteteOrigine") = True Then
                    If Datagridaffiche.Columns.Contains("LigneCodeArticle") = True Then
                        If BaseCial.FactoryArticle.ExistReference(Trim(LigneCodeArticle)) = True Then
                            If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                                DossierAdap = New OleDbDataAdapter("select * from F_ARTSTOCK where DE_No ='" & CInt(Trim(IDDepotEnteteOrigine)) & "' And AR_Ref='" & Trim(LigneCodeArticle) & "'", OleSocieteConnect)
                                DossierDs = New DataSet
                                DossierAdap.Fill(DossierDs)
                                DossierTab = DossierDs.Tables(0)
                                If DossierTab.Rows.Count <> 0 Then
                                    If DossierTab.Rows(0).Item("AS_QteSto") < CDbl(Trim(LigneQuantite)) Then
                                        ExisteQuantite = False
                                    End If
                                Else
                                    ExisteQuantite = False
                                End If
                            End If
                        Else
                            fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<ArticleStock>' and Valeurlue ='" & Join(Split(Trim(LigneCodeArticle), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                            fournisseurDs = New DataSet
                            fournisseurAdap.Fill(fournisseurDs)
                            fournisseurTab = fournisseurDs.Tables(0)
                            If fournisseurTab.Rows.Count <> 0 Then
                                If BaseCial.FactoryArticle.ExistReference(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                                    If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                                        DossierAdap = New OleDbDataAdapter("select * from F_ARTSTOCK where DE_No ='" & CInt(Trim(IDDepotEnteteOrigine)) & "' And AR_Ref='" & Trim(fournisseurTab.Rows(0).Item("Correspond")) & "'", OleSocieteConnect)
                                        DossierDs = New DataSet
                                        DossierAdap.Fill(DossierDs)
                                        DossierTab = DossierDs.Tables(0)
                                        If DossierTab.Rows.Count <> 0 Then
                                            If DossierTab.Rows(0).Item("AS_QteSto") < CDbl(Trim(LigneQuantite)) Then
                                                ExisteQuantite = False
                                            End If
                                        Else
                                            ExisteQuantite = False
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                If Datagridaffiche.Columns.Contains("LigneCodeArticle") = True Then
                    If BaseCial.FactoryArticle.ExistReference(Trim(LigneCodeArticle)) = True Then
                        If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                            If Datagridaffiche.Columns.Contains("EnteteIntituleDepotOrigine") = True Then
                                If BaseCial.FactoryDepot.ExistIntitule(Trim(EnteteIntituleDepotOrigine)) = True Then
                                    DossierAdap = New OleDbDataAdapter("select * from F_ARTSTOCK where DE_No =(Select DE_No from F_DEPOT WHERE DE_Intitule='" & Trim(EnteteIntituleDepotOrigine) & "') And AR_Ref='" & Trim(LigneCodeArticle) & "'", OleSocieteConnect)
                                    DossierDs = New DataSet
                                    DossierAdap.Fill(DossierDs)
                                    DossierTab = DossierDs.Tables(0)
                                    If DossierTab.Rows.Count <> 0 Then
                                        If DossierTab.Rows(0).Item("AS_QteSto") < CDbl(Trim(LigneQuantite)) Then
                                            ExisteQuantite = False
                                        End If
                                    Else
                                        ExisteQuantite = False
                                    End If
                                Else
                                    fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Depot Stockage>' and Valeurlue ='" & Join(Split(Trim(EnteteIntituleDepotOrigine), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                    fournisseurDs = New DataSet
                                    fournisseurAdap.Fill(fournisseurDs)
                                    fournisseurTab = fournisseurDs.Tables(0)
                                    If fournisseurTab.Rows.Count > 0 Then
                                        If BaseCial.FactoryDepot.ExistIntitule(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                                            DossierAdap = New OleDbDataAdapter("select * from F_ARTSTOCK where DE_No =(Select DE_No from F_DEPOT WHERE DE_Intitule='" & Trim(fournisseurTab.Rows(0).Item("Correspond")) & "') And AR_Ref='" & Trim(LigneCodeArticle) & "'", OleSocieteConnect)
                                            DossierDs = New DataSet
                                            DossierAdap.Fill(DossierDs)
                                            DossierTab = DossierDs.Tables(0)
                                            If DossierTab.Rows.Count <> 0 Then
                                                If DossierTab.Rows(0).Item("AS_QteSto") < CDbl(Trim(LigneQuantite)) Then
                                                    ExisteQuantite = False
                                                End If
                                            Else
                                                ExisteQuantite = False
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<ArticleStock>' and Valeurlue ='" & Join(Split(Trim(LigneCodeArticle), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                        fournisseurDs = New DataSet
                        fournisseurAdap.Fill(fournisseurDs)
                        fournisseurTab = fournisseurDs.Tables(0)
                        If fournisseurTab.Rows.Count <> 0 Then
                            If BaseCial.FactoryArticle.ExistReference(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                                If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                                    If Datagridaffiche.Columns.Contains("EnteteIntituleDepotOrigine") = True Then
                                        If BaseCial.FactoryDepot.ExistIntitule(Trim(EnteteIntituleDepotOrigine)) = True Then
                                            DossierAdap = New OleDbDataAdapter("select * from F_ARTSTOCK where DE_No =(Select DE_No from F_DEPOT WHERE DE_Intitule='" & Trim(EnteteIntituleDepotOrigine) & "') And AR_Ref='" & Trim(fournisseurTab.Rows(0).Item("Correspond")) & "'", OleSocieteConnect)
                                            DossierDs = New DataSet
                                            DossierAdap.Fill(DossierDs)
                                            DossierTab = DossierDs.Tables(0)
                                            If DossierTab.Rows.Count <> 0 Then
                                                If DossierTab.Rows(0).Item("AS_QteSto") < CDbl(Trim(LigneQuantite)) Then
                                                    ExisteQuantite = False
                                                End If
                                            Else
                                                ExisteQuantite = False
                                            End If
                                        Else
                                            Dim DepotAdap As OleDbDataAdapter
                                            Dim DepotDs As DataSet
                                            Dim DepotTab As DataTable
                                            DepotAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Depot Stockage>' and Valeurlue ='" & Join(Split(Trim(EnteteIntituleDepotOrigine), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                            DepotDs = New DataSet
                                            DepotAdap.Fill(DepotDs)
                                            DepotTab = DepotDs.Tables(0)
                                            If DepotTab.Rows.Count > 0 Then
                                                If BaseCial.FactoryDepot.ExistIntitule(Trim(DepotTab.Rows(0).Item("Correspond"))) = True Then
                                                    DossierAdap = New OleDbDataAdapter("select * from F_ARTSTOCK where DE_No =(Select DE_No from F_DEPOT WHERE DE_Intitule='" & Trim(DepotTab.Rows(0).Item("Correspond")) & "') And AR_Ref='" & Trim(fournisseurTab.Rows(0).Item("Correspond")) & "'", OleSocieteConnect)
                                                    DossierDs = New DataSet
                                                    DossierAdap.Fill(DossierDs)
                                                    DossierTab = DossierDs.Tables(0)
                                                    If DossierTab.Rows.Count <> 0 Then
                                                        If DossierTab.Rows(0).Item("AS_QteSto") < CDbl(Trim(LigneQuantite)) Then
                                                            ExisteQuantite = False
                                                        End If
                                                    Else
                                                        ExisteQuantite = False
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                If Datagridaffiche.Columns.Contains("IDDepotLigne") = True Then
                    If IsNumeric(Trim(IDDepotLigne)) = True Then
                        statistAdap = New OleDbDataAdapter("select * from F_DEPOT where DE_No ='" & CInt(Trim(IDDepotLigne)) & "'", OleSocieteConnect)
                        statistDs = New DataSet
                        statistAdap.Fill(statistDs)
                        statistTab = statistDs.Tables(0)
                        If statistTab.Rows.Count <> 0 Then
                            .Depot = BaseCial.FactoryDepot.ReadIntitule(Trim(statistTab.Rows(0).Item("DE_Intitule")))
                        Else
                            fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Depot Stockage>' and Valeurlue ='" & Join(Split(Trim(IDDepotLigne), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                            fournisseurDs = New DataSet
                            fournisseurAdap.Fill(fournisseurDs)
                            fournisseurTab = fournisseurDs.Tables(0)
                            If fournisseurTab.Rows.Count > 0 Then
                                If IsNumeric(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                                    statistAdap = New OleDbDataAdapter("select * from F_DEPOT where DE_No ='" & CInt(Trim(fournisseurTab.Rows(0).Item("Correspond"))) & "'", OleSocieteConnect)
                                    statistDs = New DataSet
                                    statistAdap.Fill(statistDs)
                                    statistTab = statistDs.Tables(0)
                                    If statistTab.Rows.Count <> 0 Then
                                        .Depot = BaseCial.FactoryDepot.ReadIntitule(Trim(statistTab.Rows(0).Item("DE_Intitule")))
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

                If Datagridaffiche.Columns.Contains("LigneIntituleDepot") = True Then
                    If BaseCial.FactoryDepot.ExistIntitule(Trim(LigneIntituleDepot)) = True Then
                        .Depot = BaseCial.FactoryDepot.ReadIntitule(Trim(LigneIntituleDepot))
                    Else
                        fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Depot Stockage>' and Valeurlue ='" & Join(Split(Trim(LigneIntituleDepot), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                        fournisseurDs = New DataSet
                        fournisseurAdap.Fill(fournisseurDs)
                        fournisseurTab = fournisseurDs.Tables(0)
                        If fournisseurTab.Rows.Count > 0 Then
                            If BaseCial.FactoryDepot.ExistIntitule(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                                .Depot = BaseCial.FactoryDepot.ReadIntitule(Trim(fournisseurTab.Rows(0).Item("Correspond")))
                            End If
                        End If
                    End If
                End If

                If Datagridaffiche.Columns.Contains("LignePrixUnitaire") = True Then
                    If EstNumeric(Trim(LignePrixUnitaire), DecimalNomb, DecimalMone) = True Then
                        .DL_PrixUnitaire = CDbl(RenvoiMontant(Trim(LignePrixUnitaire), FormatMnt, DecimalNomb, DecimalMone))
                    End If
                End If
                If ExisteQuantite = True Then
                    If Punitaire = "oui" Then
                        .WriteDefault()
                    Else
                        .Write()
                    End If
                    If Datagridaffiche.Columns.Contains("LignePoidsNet") = True Then
                        If EstNumeric(Trim(LignePoidsNet), DecimalNomb, DecimalMone) = True Then
                            .DL_PoidsNet = CDbl(RenvoiMontant(Trim(LignePoidsNet), FormatQte, DecimalNomb, DecimalMone))
                            .Write()
                        End If
                    End If
                    If Datagridaffiche.Columns.Contains("LignePoidsBrut") = True Then
                        If EstNumeric(Trim(LignePoidsBrut), DecimalNomb, DecimalMone) = True Then
                            .DL_PoidsBrut = CDbl(RenvoiMontant(Trim(LignePoidsBrut), FormatQte, DecimalNomb, DecimalMone))
                            .Write()
                        End If
                    End If
                    If IsNothing(LigneDocument.Article) = False Then
                        ErreurJrn.WriteLine("Code Article : " & Trim(LigneDocument.Article.AR_Ref) & " Créé Pour la pièce du fichier N° :" & Trim(EntetePieceInterne))
                    Else
                        ErreurJrn.WriteLine("Code Article : " & Trim("Vide") & " Créé Pour la pièce du fichier N° :" & Trim(EntetePieceInterne))
                    End If
                    'Traitement des Infos Libres
                    Try
                        If infoLigne.Count > 0 Then
                            While infoLigne.Count <> 0
                                OleAdaptaterDelete = New OleDbDataAdapter("select * From WIT_COL where Libelle='" & Trim(Datagridaffiche.Columns(infoLigne.Item(0)).Name) & "' And InfoLigne=True", OleConnenection)
                                OleDeleteDataset = New DataSet
                                OleAdaptaterDelete.Fill(OleDeleteDataset)
                                OledatableDelete = OleDeleteDataset.Tables(0)
                                If OledatableDelete.Rows.Count <> 0 Then
                                    'L'info Libre Parametrée par l'utilisateur existe dans Sage
                                    If LigneDocument.InfoLibre.Count <> 0 Then
                                        If IsNothing(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) = False Then
                                            If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                                statistAdap = New OleDbDataAdapter("select * from cbSysLibre where CB_File='F_DOCLIGNE' and CB_Name ='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "'", OleSocieteConnect)
                                                statistDs = New DataSet
                                                statistAdap.Fill(statistDs)
                                                statistTab = statistDs.Tables(0)
                                                If statistTab.Rows.Count <> 0 Then
                                                    'Texte
                                                    If statistTab.Rows(0).Item("CB_Type") = 9 Then
                                                        OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                                        OleRecherDataset = New DataSet
                                                        OleRecherAdapter.Fill(OleRecherDataset)
                                                        OleRechDatable = OleRecherDataset.Tables(0)
                                                        If OleRechDatable.Rows.Count <> 0 Then
                                                            If statistTab.Rows(0).Item("CB_Len") >= Strings.Len(Trim(OleRechDatable.Rows(0).Item("Correspond"))) Then
                                                                LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = Trim(OleRechDatable.Rows(0).Item("Correspond"))
                                                                LigneDocument.Write()
                                                            End If
                                                        Else
                                                            If statistTab.Rows(0).Item("CB_Len") >= Strings.Len(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value)) Then
                                                                LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value)
                                                                LigneDocument.Write()
                                                            End If
                                                        End If
                                                    End If
                                                    'Table
                                                    If statistTab.Rows(0).Item("CB_Type") = 22 Then
                                                        OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                                        OleRecherDataset = New DataSet
                                                        OleRecherAdapter.Fill(OleRecherDataset)
                                                        OleRechDatable = OleRecherDataset.Tables(0)
                                                        If OleRechDatable.Rows.Count <> 0 Then
                                                            If statistTab.Rows(0).Item("CB_Len") > Strings.Len(Trim(OleRechDatable.Rows(0).Item("Correspond"))) Then
                                                                LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = Trim(OleRechDatable.Rows(0).Item("Correspond"))
                                                                LigneDocument.Write()
                                                            End If
                                                        Else
                                                            If statistTab.Rows(0).Item("CB_Len") > Strings.Len(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value)) Then
                                                                LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value)
                                                                LigneDocument.Write()
                                                            End If
                                                        End If
                                                    End If
                                                    'Montant
                                                    If statistTab.Rows(0).Item("CB_Type") = 20 Then
                                                        If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                                            OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                                            OleRecherDataset = New DataSet
                                                            OleRecherAdapter.Fill(OleRecherDataset)
                                                            OleRechDatable = OleRecherDataset.Tables(0)
                                                            If OleRechDatable.Rows.Count <> 0 Then
                                                                If EstNumeric(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone) = True Then
                                                                    LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone))
                                                                    LigneDocument.Write()
                                                                End If
                                                            Else
                                                                If EstNumeric(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), DecimalNomb, DecimalMone) = True Then
                                                                    LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), DecimalNomb, DecimalMone))
                                                                    LigneDocument.Write()
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                    'Valeur
                                                    If statistTab.Rows(0).Item("CB_Type") = 7 Then
                                                        If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                                            OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                                            OleRecherDataset = New DataSet
                                                            OleRecherAdapter.Fill(OleRecherDataset)
                                                            OleRechDatable = OleRecherDataset.Tables(0)
                                                            If OleRechDatable.Rows.Count <> 0 Then
                                                                If EstNumeric(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone) = True Then
                                                                    LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone))
                                                                    LigneDocument.Write()
                                                                End If
                                                            Else
                                                                If EstNumeric(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), DecimalNomb, DecimalMone) = True Then
                                                                    LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), DecimalNomb, DecimalMone))
                                                                    LigneDocument.Write()
                                                                End If
                                                            End If
                                                        End If
                                                    End If

                                                    'Date Court
                                                    If statistTab.Rows(0).Item("CB_Type") = 3 Then
                                                        If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                                            OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                                            OleRecherDataset = New DataSet
                                                            OleRecherAdapter.Fill(OleRecherDataset)
                                                            OleRechDatable = OleRecherDataset.Tables(0)
                                                            If OleRechDatable.Rows.Count <> 0 Then
                                                                If Verificatdate(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then
                                                                    LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = RenvoieDateValide(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier)
                                                                    LigneDocument.Write()
                                                                End If
                                                            Else
                                                                If Verificatdate(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then
                                                                    LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = RenvoieDateValide(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), FormatDatefichier)
                                                                    LigneDocument.Write()
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                    'Date Longue
                                                    If statistTab.Rows(0).Item("CB_Type") = 14 Then
                                                        If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                                            OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                                            OleRecherDataset = New DataSet
                                                            OleRecherAdapter.Fill(OleRecherDataset)
                                                            OleRechDatable = OleRecherDataset.Tables(0)
                                                            If OleRechDatable.Rows.Count <> 0 Then
                                                                If Verificatdate(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then
                                                                    LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = RenvoieDateValide(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier)
                                                                    LigneDocument.Write()
                                                                End If
                                                            Else
                                                                If Verificatdate(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then
                                                                    LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = RenvoieDateValide(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), FormatDatefichier)
                                                                    LigneDocument.Write()
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Else
                                                    If IsNothing(LigneDocument.Article) = False Then
                                                        ErreurJrn.WriteLine("Code Article : " & Trim(LigneDocument.Article.AR_Ref) & " Impossible de traiter l'information libre :" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "  Il est inexistant dans la table CbsysLibre de Sage")
                                                    Else
                                                        ErreurJrn.WriteLine("Code Article : " & Trim("Vide") & " Impossible de traiter l'information libre :" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "  Il est inexistant dans la table CbsysLibre de Sage")
                                                    End If
                                                End If
                                            End If
                                        Else
                                            'nothing
                                        End If
                                    End If
                                End If
                                'L'info Libre Parametrée par l'utilisateur n'existe pas dans Sage
                                infoLigne.RemoveAt(0)
                            End While
                        End If
                    Catch ex As Exception
                        exceptionTrouve = True
                        If IsNothing(LigneDocument.Article) = False Then
                            ErreurJrn.WriteLine("Code Article : " & Trim(LigneDocument.Article.AR_Ref) & " Erreur de Création de L'information Libre , vérifiez la longueur de la chaine avec celle paramétrée/ ou la cohérence des données " & ex.Message)
                        Else
                            ErreurJrn.WriteLine("Code Article : " & Trim("Vide") & " Erreur de Création de L'information Libre , vérifiez la longueur de la chaine avec celle paramétrée/ ou la cohérence des données " & ex.Message)
                        End If
                    End Try
                Else
                    ErreurJrn.WriteLine("Le Stock dans le dépôt d'origine   ne permet pas de transférer l'article : " & LigneCodeArticle & " ,  N° :" & Trim(EntetePieceInterne))
                End If
            End With
        Catch ex As Exception
            exceptionTrouve = True
            ErreurJrn.WriteLine("Code Article : " & Trim(LigneCodeArticle) & " N°Pièce : " & EntetePieceInterne & " Erreur système de Création de l'article : " & ex.Message)
            ListePiece.Add(EntetePieceInterne)
        End Try
    End Sub
    Private Sub Creation_Ligne_ArticleSaisieD(ByRef FormatDatefichier As String, ByRef PieceArticle As String, ByRef Punitaire As String, ByRef IdentifiantArticle As String, ByRef TypeImport As Object)
        Dim fournisseurAdap As OleDbDataAdapter
        Dim fournisseurDs As DataSet
        Dim fournisseurTab As DataTable
        Dim statistAdap As OleDbDataAdapter
        Dim statistDs As DataSet
        Dim statistTab As DataTable
        Dim OleAdaptaterDelete As OleDbDataAdapter
        Dim OleDeleteDataset As DataSet
        Dim OledatableDelete As DataTable
        Dim OleRecherAdapter As OleDbDataAdapter
        Dim OleRecherDataset As DataSet
        Dim OleRechDatable As DataTable
        Dim DossierAdap As OleDbDataAdapter
        Dim DossierDs As DataSet
        Dim DossierTab As DataTable
        Dim FormatQte As Integer = 0
        Dim FormatMnt As Integer = 0
        DossierAdap = New OleDbDataAdapter("select * from P_DOSSIERCIAL", OleSocieteConnect)
        DossierDs = New DataSet
        DossierAdap.Fill(DossierDs)
        DossierTab = DossierDs.Tables(0)
        If DossierTab.Rows.Count <> 0 Then
            If InStr(DossierTab.Rows(0).Item("D_FormatQte"), ",") <> 0 Then
                If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ",")), ".") <> 0 Then
                    FormatQte = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ","))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ",")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ","))), ".")))
                Else
                    FormatQte = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ",")))
                End If
            Else
                If InStr(DossierTab.Rows(0).Item("D_FormatQte"), ".") <> 0 Then
                    If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ".")), ",") <> 0 Then
                        FormatQte = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), "."))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ".")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), "."))), ",")))
                    Else
                        FormatQte = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ".")))
                    End If
                End If
            End If

            If InStr(DossierTab.Rows(0).Item("D_FormatPrix"), ",") <> 0 Then
                If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ",")), ".") <> 0 Then
                    FormatMnt = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ","))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ",")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ","))), ".")))
                Else
                    FormatMnt = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ",")))
                End If
            Else
                If InStr(DossierTab.Rows(0).Item("D_FormatPrix"), ".") <> 0 Then
                    If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ".")), ",") <> 0 Then
                        FormatMnt = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), "."))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ".")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), "."))), ",")))
                    Else
                        FormatMnt = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ".")))
                    End If
                End If
            End If
        End If
        Try
            LigneDocument = Document.FactoryDocumentLigne.Create
            With LigneDocument
                If Datagridaffiche.Columns.Contains("LigneDesignationArticle") = True Then
                    .DL_Design = LigneDesignationArticle
                End If
                If Datagridaffiche.Columns.Contains("LignePoidsNet") = True Then
                    If EstNumeric(Trim(LignePoidsNet), DecimalNomb, DecimalMone) = True Then
                        .DL_PoidsNet = CDbl(RenvoiMontant(Trim(LignePoidsNet), FormatQte, DecimalNomb, DecimalMone))
                    End If
                End If
                If Datagridaffiche.Columns.Contains("LignePoidsBrut") = True Then
                    If EstNumeric(Trim(LignePoidsBrut), DecimalNomb, DecimalMone) = True Then
                        .DL_PoidsBrut = CDbl(RenvoiMontant(Trim(LignePoidsBrut), FormatQte, DecimalNomb, DecimalMone))
                    End If
                End If
                If Datagridaffiche.Columns.Contains("LigneReference") = True Then
                    .DO_Ref = Trim(LigneReference)
                End If
                If Datagridaffiche.Columns.Contains("LigneDatedeFabrication") = True Then
                    If Trim(LigneDatedeFabrication) <> "" Then
                        .LS_Fabrication = RenvoieDateValide(Trim(LigneDatedeFabrication), FormatDatefichier)
                    End If
                End If
                '.Valorisee = True
                If Datagridaffiche.Columns.Contains("LigneDatedePeremption") = True Then
                    If Trim(LigneDatedePeremption) <> "" Then
                        .LS_Peremption = RenvoieDateValide(Trim(LigneDatedePeremption), FormatDatefichier)
                    End If
                End If
                .DL_PrixUnitaire = PuStock
                If Datagridaffiche.Columns.Contains("LigneNSerieLot") = True Then
                    .LS_NoSerie = LigneNSerieLot
                End If
                If Datagridaffiche.Columns.Contains("LigneCodeArticle") = True Then
                    If BaseCial.FactoryArticle.ExistReference(Trim(LigneCodeArticle)) = True Then
                        If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                            .SetDefaultArticle(BaseCial.FactoryArticle.ReadReference(Trim(LigneCodeArticle)), CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)))
                        End If
                    Else
                        fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<ArticleStock>' and Valeurlue ='" & Join(Split(Trim(LigneCodeArticle), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                        fournisseurDs = New DataSet
                        fournisseurAdap.Fill(fournisseurDs)
                        fournisseurTab = fournisseurDs.Tables(0)
                        If fournisseurTab.Rows.Count <> 0 Then
                            If BaseCial.FactoryArticle.ExistReference(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                                If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                                    .SetDefaultArticle(BaseCial.FactoryArticle.ReadReference(Trim(fournisseurTab.Rows(0).Item("Correspond"))), CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)))
                                End If
                            End If
                        End If
                    End If
                End If
                If Datagridaffiche.Columns.Contains("IDDepotEnteteDestination") = True Then
                    If IsNumeric(Trim(IDDepotEnteteDestination)) = True Then
                        statistAdap = New OleDbDataAdapter("select * from F_DEPOT where DE_No ='" & CInt(Trim(IDDepotEnteteDestination)) & "'", OleSocieteConnect)
                        statistDs = New DataSet
                        statistAdap.Fill(statistDs)
                        statistTab = statistDs.Tables(0)
                        If statistTab.Rows.Count <> 0 Then
                            .Depot = BaseCial.FactoryDepot.ReadIntitule(Trim(statistTab.Rows(0).Item("DE_Intitule")))
                        Else
                            fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Depot Stockage>' and Valeurlue ='" & Join(Split(Trim(IDDepotEnteteDestination), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                            fournisseurDs = New DataSet
                            fournisseurAdap.Fill(fournisseurDs)
                            fournisseurTab = fournisseurDs.Tables(0)
                            If fournisseurTab.Rows.Count > 0 Then
                                If IsNumeric(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                                    statistAdap = New OleDbDataAdapter("select * from F_DEPOT where DE_No ='" & CInt(Trim(fournisseurTab.Rows(0).Item("Correspond"))) & "'", OleSocieteConnect)
                                    statistDs = New DataSet
                                    statistAdap.Fill(statistDs)
                                    statistTab = statistDs.Tables(0)
                                    If statistTab.Rows.Count <> 0 Then
                                        .Depot = BaseCial.FactoryDepot.ReadIntitule(Trim(statistTab.Rows(0).Item("DE_Intitule")))
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

                If Datagridaffiche.Columns.Contains("EnteteIntituleDepotDestination") = True Then
                    If BaseCial.FactoryDepot.ExistIntitule(Trim(EnteteIntituleDepotDestination)) = True Then
                        .Depot = BaseCial.FactoryDepot.ReadIntitule(Trim(EnteteIntituleDepotDestination))
                    Else
                        fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Depot Stockage>' and Valeurlue ='" & Join(Split(Trim(EnteteIntituleDepotDestination), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                        fournisseurDs = New DataSet
                        fournisseurAdap.Fill(fournisseurDs)
                        fournisseurTab = fournisseurDs.Tables(0)
                        If fournisseurTab.Rows.Count > 0 Then
                            If BaseCial.FactoryDepot.ExistIntitule(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                                .Depot = BaseCial.FactoryDepot.ReadIntitule(Trim(fournisseurTab.Rows(0).Item("Correspond")))
                            End If
                        End If
                    End If
                End If
                If Datagridaffiche.Columns.Contains("LignePrixUnitaire") = True Then
                    If EstNumeric(Trim(LignePrixUnitaire), DecimalNomb, DecimalMone) = True Then
                        .DL_PrixUnitaire = CDbl(RenvoiMontant(Trim(LignePrixUnitaire), FormatMnt, DecimalNomb, DecimalMone))
                    End If
                End If
                If Punitaire = "oui" Then
                    .WriteDefault()
                Else
                    .Write()
                End If
                If Datagridaffiche.Columns.Contains("LignePoidsNet") = True Then
                    If EstNumeric(Trim(LignePoidsNet), DecimalNomb, DecimalMone) = True Then
                        .DL_PoidsNet = CDbl(RenvoiMontant(Trim(LignePoidsNet), FormatQte, DecimalNomb, DecimalMone))
                        .Write()
                    End If
                End If
                If Datagridaffiche.Columns.Contains("LignePoidsBrut") = True Then
                    If EstNumeric(Trim(LignePoidsBrut), DecimalNomb, DecimalMone) = True Then
                        .DL_PoidsBrut = CDbl(RenvoiMontant(Trim(LignePoidsBrut), FormatQte, DecimalNomb, DecimalMone))
                        .Write()
                    End If
                End If
                If IsNothing(LigneDocument.Article) = False Then
                    ErreurJrn.WriteLine("Code Article : " & Trim(LigneDocument.Article.AR_Ref) & " Créé Pour la pièce du fichier N° :" & Trim(EntetePieceInterne))
                Else
                    ErreurJrn.WriteLine("Code Article : " & Trim("Vide") & " Créé Pour la pièce du fichier N° :" & Trim(EntetePieceInterne))
                End If
                'Traitement des Infos Libres
                Try
                    If infoLigne.Count > 0 Then
                        While infoLigne.Count <> 0
                            OleAdaptaterDelete = New OleDbDataAdapter("select * From WIT_COL where Libelle='" & Trim(Datagridaffiche.Columns(infoLigne.Item(0)).Name) & "' And InfoLigne=True", OleConnenection)
                            OleDeleteDataset = New DataSet
                            OleAdaptaterDelete.Fill(OleDeleteDataset)
                            OledatableDelete = OleDeleteDataset.Tables(0)
                            If OledatableDelete.Rows.Count <> 0 Then
                                'L'info Libre Parametrée par l'utilisateur existe dans Sage
                                If LigneDocument.InfoLibre.Count <> 0 Then
                                    If IsNothing(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) = False Then
                                        If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                            statistAdap = New OleDbDataAdapter("select * from cbSysLibre where CB_File='F_DOCLIGNE' and CB_Name ='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "'", OleSocieteConnect)
                                            statistDs = New DataSet
                                            statistAdap.Fill(statistDs)
                                            statistTab = statistDs.Tables(0)
                                            If statistTab.Rows.Count <> 0 Then
                                                'Texte
                                                If statistTab.Rows(0).Item("CB_Type") = 9 Then
                                                    OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                                    OleRecherDataset = New DataSet
                                                    OleRecherAdapter.Fill(OleRecherDataset)
                                                    OleRechDatable = OleRecherDataset.Tables(0)
                                                    If OleRechDatable.Rows.Count <> 0 Then
                                                        If statistTab.Rows(0).Item("CB_Len") >= Strings.Len(Trim(OleRechDatable.Rows(0).Item("Correspond"))) Then
                                                            LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = Trim(OleRechDatable.Rows(0).Item("Correspond"))
                                                            LigneDocument.Write()
                                                        End If
                                                    Else
                                                        If statistTab.Rows(0).Item("CB_Len") >= Strings.Len(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value)) Then
                                                            LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value)
                                                            LigneDocument.Write()
                                                        End If
                                                    End If
                                                End If
                                                'Table
                                                If statistTab.Rows(0).Item("CB_Type") = 22 Then
                                                    OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                                    OleRecherDataset = New DataSet
                                                    OleRecherAdapter.Fill(OleRecherDataset)
                                                    OleRechDatable = OleRecherDataset.Tables(0)
                                                    If OleRechDatable.Rows.Count <> 0 Then
                                                        If statistTab.Rows(0).Item("CB_Len") > Strings.Len(Trim(OleRechDatable.Rows(0).Item("Correspond"))) Then
                                                            LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = Trim(OleRechDatable.Rows(0).Item("Correspond"))
                                                            LigneDocument.Write()
                                                        End If
                                                    Else
                                                        If statistTab.Rows(0).Item("CB_Len") > Strings.Len(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value)) Then
                                                            LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value)
                                                            LigneDocument.Write()
                                                        End If
                                                    End If
                                                End If
                                                'Montant
                                                If statistTab.Rows(0).Item("CB_Type") = 20 Then
                                                    If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                                        OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                                        OleRecherDataset = New DataSet
                                                        OleRecherAdapter.Fill(OleRecherDataset)
                                                        OleRechDatable = OleRecherDataset.Tables(0)
                                                        If OleRechDatable.Rows.Count <> 0 Then
                                                            If EstNumeric(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone) = True Then
                                                                LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone))
                                                                LigneDocument.Write()
                                                            End If
                                                        Else
                                                            If EstNumeric(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), DecimalNomb, DecimalMone) = True Then
                                                                LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), DecimalNomb, DecimalMone))
                                                                LigneDocument.Write()
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                                'Valeur
                                                If statistTab.Rows(0).Item("CB_Type") = 7 Then
                                                    If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                                        OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                                        OleRecherDataset = New DataSet
                                                        OleRecherAdapter.Fill(OleRecherDataset)
                                                        OleRechDatable = OleRecherDataset.Tables(0)
                                                        If OleRechDatable.Rows.Count <> 0 Then
                                                            If EstNumeric(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone) = True Then
                                                                LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone))
                                                                LigneDocument.Write()
                                                            End If
                                                        Else
                                                            If EstNumeric(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), DecimalNomb, DecimalMone) = True Then
                                                                LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), DecimalNomb, DecimalMone))
                                                                LigneDocument.Write()
                                                            End If
                                                        End If
                                                    End If
                                                End If

                                                'Date Court
                                                If statistTab.Rows(0).Item("CB_Type") = 3 Then
                                                    If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                                        OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                                        OleRecherDataset = New DataSet
                                                        OleRecherAdapter.Fill(OleRecherDataset)
                                                        OleRechDatable = OleRecherDataset.Tables(0)
                                                        If OleRechDatable.Rows.Count <> 0 Then
                                                            If Verificatdate(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then
                                                                LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = RenvoieDateValide(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier)
                                                                LigneDocument.Write()
                                                            End If
                                                        Else
                                                            If Verificatdate(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then
                                                                LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = RenvoieDateValide(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), FormatDatefichier)
                                                                LigneDocument.Write()
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                                'Date Longue
                                                If statistTab.Rows(0).Item("CB_Type") = 14 Then
                                                    If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                                        OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                                        OleRecherDataset = New DataSet
                                                        OleRecherAdapter.Fill(OleRecherDataset)
                                                        OleRechDatable = OleRecherDataset.Tables(0)
                                                        If OleRechDatable.Rows.Count <> 0 Then
                                                            If Verificatdate(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then
                                                                LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = RenvoieDateValide(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier)
                                                                LigneDocument.Write()
                                                            End If
                                                        Else
                                                            If Verificatdate(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then
                                                                LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = RenvoieDateValide(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), FormatDatefichier)
                                                                LigneDocument.Write()
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If IsNothing(LigneDocument.Article) = False Then
                                                    ErreurJrn.WriteLine("Code Article : " & Trim(LigneDocument.Article.AR_Ref) & " Impossible de traiter l'information libre :" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "  Il est inexistant dans la table CbsysLibre de Sage")
                                                Else
                                                    ErreurJrn.WriteLine("Code Article : " & Trim("Vide") & " Impossible de traiter l'information libre :" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "  Il est inexistant dans la table CbsysLibre de Sage")
                                                End If
                                            End If
                                        End If
                                    Else
                                        'nothing
                                    End If
                                End If
                            End If
                            'L'info Libre Parametrée par l'utilisateur n'existe pas dans Sage
                            infoLigne.RemoveAt(0)
                        End While
                    End If
                Catch ex As Exception
                    exceptionTrouve = True
                    If IsNothing(LigneDocument.Article) = False Then
                        ErreurJrn.WriteLine("Code Article : " & Trim(LigneDocument.Article.AR_Ref) & " Erreur de Création de L'information Libre , vérifiez la longueur de la chaine avec celle paramétrée/ ou la cohérence des données " & ex.Message)
                    Else
                        ErreurJrn.WriteLine("Code Article : " & Trim("Vide") & " Erreur de Création de L'information Libre , vérifiez la longueur de la chaine avec celle paramétrée/ ou la cohérence des données " & ex.Message)
                    End If
                End Try

            End With
        Catch ex As Exception
            exceptionTrouve = True
            ErreurJrn.WriteLine("Code Article : " & Trim(LigneCodeArticle) & " N°Pièce : " & EntetePieceInterne & " Erreur système de Création de l'article : " & ex.Message)
            ListePiece.Add(EntetePieceInterne)
        End Try
    End Sub
    Private Function Creation_Ligne_ArticleSaisieO(ByRef FormatDatefichier As String, ByRef PieceArticle As String, ByRef Punitaire As String, ByRef IdentifiantArticle As String, ByRef TypeImport As Object) As Boolean
        Dim fournisseurAdap As OleDbDataAdapter
        Dim fournisseurDs As DataSet
        Dim fournisseurTab As DataTable
        Dim statistAdap As OleDbDataAdapter
        Dim statistDs As DataSet
        Dim statistTab As DataTable
        Dim OleAdaptaterDelete As OleDbDataAdapter
        Dim OleDeleteDataset As DataSet
        Dim OledatableDelete As DataTable
        Dim OleRecherAdapter As OleDbDataAdapter
        Dim OleRecherDataset As DataSet
        Dim OleRechDatable As DataTable
        Dim DossierAdap As OleDbDataAdapter
        Dim DossierDs As DataSet
        Dim DossierTab As DataTable
        Dim FormatQte As Integer = 0
        Dim FormatMnt As Integer = 0
        Dim ExisteQuantite As Boolean = True
        DossierAdap = New OleDbDataAdapter("select * from P_DOSSIERCIAL", OleSocieteConnect)
        DossierDs = New DataSet
        DossierAdap.Fill(DossierDs)
        DossierTab = DossierDs.Tables(0)
        If DossierTab.Rows.Count <> 0 Then
            If InStr(DossierTab.Rows(0).Item("D_FormatQte"), ",") <> 0 Then
                If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ",")), ".") <> 0 Then
                    FormatQte = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ","))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ",")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ","))), ".")))
                Else
                    FormatQte = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ",")))
                End If
            Else
                If InStr(DossierTab.Rows(0).Item("D_FormatQte"), ".") <> 0 Then
                    If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ".")), ",") <> 0 Then
                        FormatQte = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), "."))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ".")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), "."))), ",")))
                    Else
                        FormatQte = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ".")))
                    End If
                End If
            End If

            If InStr(DossierTab.Rows(0).Item("D_FormatPrix"), ",") <> 0 Then
                If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ",")), ".") <> 0 Then
                    FormatMnt = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ","))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ",")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ","))), ".")))
                Else
                    FormatMnt = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ",")))
                End If
            Else
                If InStr(DossierTab.Rows(0).Item("D_FormatPrix"), ".") <> 0 Then
                    If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ".")), ",") <> 0 Then
                        FormatMnt = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), "."))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ".")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), "."))), ",")))
                    Else
                        FormatMnt = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ".")))
                    End If
                End If
            End If
        End If
        Try

            LigneDocument = Document.FactoryDocumentLigne.Create
            With LigneDocument
                If Datagridaffiche.Columns.Contains("LigneDesignationArticle") = True Then
                    .DL_Design = LigneDesignationArticle
                End If
                If Datagridaffiche.Columns.Contains("LignePoidsNet") = True Then
                    If EstNumeric(Trim(LignePoidsNet), DecimalNomb, DecimalMone) = True Then
                        .DL_PoidsNet = CDbl(RenvoiMontant(Trim(LignePoidsNet), FormatQte, DecimalNomb, DecimalMone))
                    End If
                End If
                If Datagridaffiche.Columns.Contains("LignePoidsBrut") = True Then
                    If EstNumeric(Trim(LignePoidsBrut), DecimalNomb, DecimalMone) = True Then
                        .DL_PoidsBrut = CDbl(RenvoiMontant(Trim(LignePoidsBrut), FormatQte, DecimalNomb, DecimalMone))
                    End If
                End If
                If Datagridaffiche.Columns.Contains("LigneReference") = True Then
                    .DO_Ref = Trim(LigneReference)
                End If
                If Datagridaffiche.Columns.Contains("LigneDatedeFabrication") = True Then
                    If Trim(LigneDatedeFabrication) <> "" Then
                        .LS_Fabrication = RenvoieDateValide(Trim(LigneDatedeFabrication), FormatDatefichier)
                    End If
                End If
                If Datagridaffiche.Columns.Contains("LigneDatedePeremption") = True Then
                    If Trim(LigneDatedePeremption) <> "" Then
                        .LS_Peremption = RenvoieDateValide(Trim(LigneDatedePeremption), FormatDatefichier)
                    End If
                End If
                .Valorisee = True
                If Datagridaffiche.Columns.Contains("LignePrixUnitaire") = True Then
                    If EstNumeric(Trim(LignePrixUnitaire), DecimalNomb, DecimalMone) = True Then
                        .DL_PrixUnitaire = CDbl(RenvoiMontant(Trim(LignePrixUnitaire), FormatMnt, DecimalNomb, DecimalMone))
                    End If
                End If
                If Datagridaffiche.Columns.Contains("LigneNSerieLot") = True Then
                    .LS_NoSerie = LigneNSerieLot
                End If
                If Datagridaffiche.Columns.Contains("LigneCodeArticle") = True Then
                    If BaseCial.FactoryArticle.ExistReference(Trim(LigneCodeArticle)) = True Then
                        If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                            .SetDefaultArticle(BaseCial.FactoryArticle.ReadReference(Trim(LigneCodeArticle)), CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)))
                        End If
                    Else
                        fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<ArticleStock>' and Valeurlue ='" & Join(Split(Trim(LigneCodeArticle), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                        fournisseurDs = New DataSet
                        fournisseurAdap.Fill(fournisseurDs)
                        fournisseurTab = fournisseurDs.Tables(0)
                        If fournisseurTab.Rows.Count <> 0 Then
                            If BaseCial.FactoryArticle.ExistReference(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                                If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                                    .SetDefaultArticle(BaseCial.FactoryArticle.ReadReference(Trim(fournisseurTab.Rows(0).Item("Correspond"))), CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)))
                                End If
                            End If
                        End If
                    End If
                End If
                If Datagridaffiche.Columns.Contains("IDDepotEnteteOrigine") = True Then
                    If Datagridaffiche.Columns.Contains("LigneCodeArticle") = True Then
                        If BaseCial.FactoryArticle.ExistReference(Trim(LigneCodeArticle)) = True Then
                            If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                                DossierAdap = New OleDbDataAdapter("select * from F_ARTSTOCK where DE_No ='" & CInt(Trim(IDDepotEnteteOrigine)) & "' And AR_Ref='" & Trim(LigneCodeArticle) & "'", OleSocieteConnect)
                                DossierDs = New DataSet
                                DossierAdap.Fill(DossierDs)
                                DossierTab = DossierDs.Tables(0)
                                If DossierTab.Rows.Count <> 0 Then
                                    If DossierTab.Rows(0).Item("AS_QteSto") < CDbl(Trim(LigneQuantite)) Then
                                        ExisteQuantite = False
                                    End If
                                Else
                                    ExisteQuantite = False
                                End If
                            End If
                        Else
                            fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<ArticleStock>' and Valeurlue ='" & Join(Split(Trim(LigneCodeArticle), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                            fournisseurDs = New DataSet
                            fournisseurAdap.Fill(fournisseurDs)
                            fournisseurTab = fournisseurDs.Tables(0)
                            If fournisseurTab.Rows.Count <> 0 Then
                                If BaseCial.FactoryArticle.ExistReference(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                                    If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                                        DossierAdap = New OleDbDataAdapter("select * from F_ARTSTOCK where DE_No ='" & CInt(Trim(IDDepotEnteteOrigine)) & "' And AR_Ref='" & Trim(fournisseurTab.Rows(0).Item("Correspond")) & "'", OleSocieteConnect)
                                        DossierDs = New DataSet
                                        DossierAdap.Fill(DossierDs)
                                        DossierTab = DossierDs.Tables(0)
                                        If DossierTab.Rows.Count <> 0 Then
                                            If DossierTab.Rows(0).Item("AS_QteSto") < CDbl(Trim(LigneQuantite)) Then
                                                ExisteQuantite = False
                                            End If
                                        Else
                                            ExisteQuantite = False
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                If Datagridaffiche.Columns.Contains("LigneCodeArticle") = True Then
                    If BaseCial.FactoryArticle.ExistReference(Trim(LigneCodeArticle)) = True Then
                        If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                            If Datagridaffiche.Columns.Contains("EnteteIntituleDepotOrigine") = True Then
                                If BaseCial.FactoryDepot.ExistIntitule(Trim(EnteteIntituleDepotOrigine)) = True Then
                                    DossierAdap = New OleDbDataAdapter("select * from F_ARTSTOCK where DE_No =(Select DE_No from F_DEPOT WHERE DE_Intitule='" & Trim(EnteteIntituleDepotOrigine) & "') And AR_Ref='" & Trim(LigneCodeArticle) & "'", OleSocieteConnect)
                                    DossierDs = New DataSet
                                    DossierAdap.Fill(DossierDs)
                                    DossierTab = DossierDs.Tables(0)
                                    If DossierTab.Rows.Count <> 0 Then
                                        If DossierTab.Rows(0).Item("AS_QteSto") < CDbl(Trim(LigneQuantite)) Then
                                            ExisteQuantite = False
                                        End If
                                    Else
                                        ExisteQuantite = False
                                    End If
                                Else
                                    fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Depot Stockage>' and Valeurlue ='" & Join(Split(Trim(EnteteIntituleDepotOrigine), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                    fournisseurDs = New DataSet
                                    fournisseurAdap.Fill(fournisseurDs)
                                    fournisseurTab = fournisseurDs.Tables(0)
                                    If fournisseurTab.Rows.Count > 0 Then
                                        If BaseCial.FactoryDepot.ExistIntitule(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                                            DossierAdap = New OleDbDataAdapter("select * from F_ARTSTOCK where DE_No =(Select DE_No from F_DEPOT WHERE DE_Intitule='" & Trim(fournisseurTab.Rows(0).Item("Correspond")) & "') And AR_Ref='" & Trim(LigneCodeArticle) & "'", OleSocieteConnect)
                                            DossierDs = New DataSet
                                            DossierAdap.Fill(DossierDs)
                                            DossierTab = DossierDs.Tables(0)
                                            If DossierTab.Rows.Count <> 0 Then
                                                If DossierTab.Rows(0).Item("AS_QteSto") < CDbl(Trim(LigneQuantite)) Then
                                                    ExisteQuantite = False
                                                End If
                                            Else
                                                ExisteQuantite = False
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<ArticleStock>' and Valeurlue ='" & Join(Split(Trim(LigneCodeArticle), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                        fournisseurDs = New DataSet
                        fournisseurAdap.Fill(fournisseurDs)
                        fournisseurTab = fournisseurDs.Tables(0)
                        If fournisseurTab.Rows.Count <> 0 Then
                            If BaseCial.FactoryArticle.ExistReference(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                                If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                                    If Datagridaffiche.Columns.Contains("EnteteIntituleDepotOrigine") = True Then
                                        If BaseCial.FactoryDepot.ExistIntitule(Trim(EnteteIntituleDepotOrigine)) = True Then
                                            DossierAdap = New OleDbDataAdapter("select * from F_ARTSTOCK where DE_No =(Select DE_No from F_DEPOT WHERE DE_Intitule='" & Trim(EnteteIntituleDepotOrigine) & "') And AR_Ref='" & Trim(fournisseurTab.Rows(0).Item("Correspond")) & "'", OleSocieteConnect)
                                            DossierDs = New DataSet
                                            DossierAdap.Fill(DossierDs)
                                            DossierTab = DossierDs.Tables(0)
                                            If DossierTab.Rows.Count <> 0 Then
                                                If DossierTab.Rows(0).Item("AS_QteSto") < CDbl(Trim(LigneQuantite)) Then
                                                    ExisteQuantite = False
                                                End If
                                            Else
                                                ExisteQuantite = False
                                            End If
                                        Else
                                            OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Depot Stockage>' and Valeurlue ='" & Join(Split(Trim(EnteteIntituleDepotOrigine), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                            OleRecherDataset = New DataSet
                                            OleRecherAdapter.Fill(OleRecherDataset)
                                            OleRechDatable = OleRecherDataset.Tables(0)
                                            If OleRechDatable.Rows.Count > 0 Then
                                                If BaseCial.FactoryDepot.ExistIntitule(Trim(OleRechDatable.Rows(0).Item("Correspond"))) = True Then
                                                    DossierAdap = New OleDbDataAdapter("select * from F_ARTSTOCK where DE_No =(Select DE_No from F_DEPOT WHERE DE_Intitule='" & Trim(OleRechDatable.Rows(0).Item("Correspond")) & "') And AR_Ref='" & Trim(fournisseurTab.Rows(0).Item("Correspond")) & "'", OleSocieteConnect)
                                                    DossierDs = New DataSet
                                                    DossierAdap.Fill(DossierDs)
                                                    DossierTab = DossierDs.Tables(0)
                                                    If DossierTab.Rows.Count <> 0 Then
                                                        If DossierTab.Rows(0).Item("AS_QteSto") < CDbl(Trim(LigneQuantite)) Then
                                                            ExisteQuantite = False
                                                        End If
                                                    Else
                                                        ExisteQuantite = False
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                If Datagridaffiche.Columns.Contains("IDDepotEnteteOrigine") = True Then
                    If IsNumeric(Trim(IDDepotEnteteOrigine)) = True Then
                        statistAdap = New OleDbDataAdapter("select * from F_DEPOT where DE_No ='" & CInt(Trim(IDDepotEnteteOrigine)) & "'", OleSocieteConnect)
                        statistDs = New DataSet
                        statistAdap.Fill(statistDs)
                        statistTab = statistDs.Tables(0)
                        If statistTab.Rows.Count <> 0 Then
                            .Depot = BaseCial.FactoryDepot.ReadIntitule(Trim(statistTab.Rows(0).Item("DE_Intitule")))
                        Else
                            fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Depot Stockage>' and Valeurlue ='" & Join(Split(Trim(IDDepotEnteteOrigine), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                            fournisseurDs = New DataSet
                            fournisseurAdap.Fill(fournisseurDs)
                            fournisseurTab = fournisseurDs.Tables(0)
                            If fournisseurTab.Rows.Count > 0 Then
                                If IsNumeric(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                                    statistAdap = New OleDbDataAdapter("select * from F_DEPOT where DE_No ='" & CInt(Trim(fournisseurTab.Rows(0).Item("Correspond"))) & "'", OleSocieteConnect)
                                    statistDs = New DataSet
                                    statistAdap.Fill(statistDs)
                                    statistTab = statistDs.Tables(0)
                                    If statistTab.Rows.Count <> 0 Then
                                        .Depot = BaseCial.FactoryDepot.ReadIntitule(Trim(statistTab.Rows(0).Item("DE_Intitule")))
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

                If Datagridaffiche.Columns.Contains("EnteteIntituleDepotOrigine") = True Then
                    If BaseCial.FactoryDepot.ExistIntitule(Trim(EnteteIntituleDepotOrigine)) = True Then
                        .Depot = BaseCial.FactoryDepot.ReadIntitule(Trim(EnteteIntituleDepotOrigine))
                    Else
                        fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Depot Stockage>' and Valeurlue ='" & Join(Split(Trim(EnteteIntituleDepotOrigine), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                        fournisseurDs = New DataSet
                        fournisseurAdap.Fill(fournisseurDs)
                        fournisseurTab = fournisseurDs.Tables(0)
                        If fournisseurTab.Rows.Count > 0 Then
                            If BaseCial.FactoryDepot.ExistIntitule(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                                .Depot = BaseCial.FactoryDepot.ReadIntitule(Trim(fournisseurTab.Rows(0).Item("Correspond")))
                            End If
                        End If
                    End If
                End If
                If ExisteQuantite = True Then
                    If Punitaire = "oui" Then
                        .WriteDefault()
                    Else
                        .Write()
                    End If
                    If Datagridaffiche.Columns.Contains("LignePoidsNet") = True Then
                        If EstNumeric(Trim(LignePoidsNet), DecimalNomb, DecimalMone) = True Then
                            .DL_PoidsNet = CDbl(RenvoiMontant(Trim(LignePoidsNet), FormatQte, DecimalNomb, DecimalMone))
                            .Write()
                        End If
                    End If
                    If Datagridaffiche.Columns.Contains("LignePoidsBrut") = True Then
                        If EstNumeric(Trim(LignePoidsBrut), DecimalNomb, DecimalMone) = True Then
                            .DL_PoidsBrut = CDbl(RenvoiMontant(Trim(LignePoidsBrut), FormatQte, DecimalNomb, DecimalMone))
                            .Write()
                        End If
                    End If
                    PuStock = .DL_PrixUnitaire
                    If IsNothing(LigneDocument.Article) = False Then
                        ErreurJrn.WriteLine("Code Article : " & Trim(LigneDocument.Article.AR_Ref) & " Créé Pour la pièce du fichier N° :" & Trim(EntetePieceInterne))
                    Else
                        ErreurJrn.WriteLine("Code Article : " & Trim("Vide") & " Créé Pour la pièce du fichier N° :" & Trim(EntetePieceInterne))
                    End If
                    'Traitement des Infos Libres
                    Try
                        If infoLigne.Count > 0 Then
                            While infoLigne.Count <> 0
                                OleAdaptaterDelete = New OleDbDataAdapter("select * From WIT_COL where Libelle='" & Trim(Datagridaffiche.Columns(infoLigne.Item(0)).Name) & "' And InfoLigne=True", OleConnenection)
                                OleDeleteDataset = New DataSet
                                OleAdaptaterDelete.Fill(OleDeleteDataset)
                                OledatableDelete = OleDeleteDataset.Tables(0)
                                If OledatableDelete.Rows.Count <> 0 Then
                                    'L'info Libre Parametrée par l'utilisateur existe dans Sage
                                    If LigneDocument.InfoLibre.Count <> 0 Then
                                        If IsNothing(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) = False Then
                                            If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                                statistAdap = New OleDbDataAdapter("select * from cbSysLibre where CB_File='F_DOCLIGNE' and CB_Name ='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "'", OleSocieteConnect)
                                                statistDs = New DataSet
                                                statistAdap.Fill(statistDs)
                                                statistTab = statistDs.Tables(0)
                                                If statistTab.Rows.Count <> 0 Then
                                                    'Texte
                                                    If statistTab.Rows(0).Item("CB_Type") = 9 Then
                                                        OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                                        OleRecherDataset = New DataSet
                                                        OleRecherAdapter.Fill(OleRecherDataset)
                                                        OleRechDatable = OleRecherDataset.Tables(0)
                                                        If OleRechDatable.Rows.Count <> 0 Then
                                                            If statistTab.Rows(0).Item("CB_Len") >= Strings.Len(Trim(OleRechDatable.Rows(0).Item("Correspond"))) Then
                                                                LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = Trim(OleRechDatable.Rows(0).Item("Correspond"))
                                                                LigneDocument.Write()
                                                            End If
                                                        Else
                                                            If statistTab.Rows(0).Item("CB_Len") >= Strings.Len(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value)) Then
                                                                LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value)
                                                                LigneDocument.Write()
                                                            End If
                                                        End If
                                                    End If
                                                    'Table
                                                    If statistTab.Rows(0).Item("CB_Type") = 22 Then
                                                        OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                                        OleRecherDataset = New DataSet
                                                        OleRecherAdapter.Fill(OleRecherDataset)
                                                        OleRechDatable = OleRecherDataset.Tables(0)
                                                        If OleRechDatable.Rows.Count <> 0 Then
                                                            If statistTab.Rows(0).Item("CB_Len") > Strings.Len(Trim(OleRechDatable.Rows(0).Item("Correspond"))) Then
                                                                LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = Trim(OleRechDatable.Rows(0).Item("Correspond"))
                                                                LigneDocument.Write()
                                                            End If
                                                        Else
                                                            If statistTab.Rows(0).Item("CB_Len") > Strings.Len(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value)) Then
                                                                LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value)
                                                                LigneDocument.Write()
                                                            End If
                                                        End If
                                                    End If
                                                    'Montant
                                                    If statistTab.Rows(0).Item("CB_Type") = 20 Then
                                                        If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                                            OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                                            OleRecherDataset = New DataSet
                                                            OleRecherAdapter.Fill(OleRecherDataset)
                                                            OleRechDatable = OleRecherDataset.Tables(0)
                                                            If OleRechDatable.Rows.Count <> 0 Then
                                                                If EstNumeric(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone) = True Then
                                                                    LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone))
                                                                    LigneDocument.Write()
                                                                End If
                                                            Else
                                                                If EstNumeric(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), DecimalNomb, DecimalMone) = True Then
                                                                    LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), DecimalNomb, DecimalMone))
                                                                    LigneDocument.Write()
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                    'Valeur
                                                    If statistTab.Rows(0).Item("CB_Type") = 7 Then
                                                        If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                                            OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                                            OleRecherDataset = New DataSet
                                                            OleRecherAdapter.Fill(OleRecherDataset)
                                                            OleRechDatable = OleRecherDataset.Tables(0)
                                                            If OleRechDatable.Rows.Count <> 0 Then
                                                                If EstNumeric(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone) = True Then
                                                                    LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone))
                                                                    LigneDocument.Write()
                                                                End If
                                                            Else
                                                                If EstNumeric(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), DecimalNomb, DecimalMone) = True Then
                                                                    LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), DecimalNomb, DecimalMone))
                                                                    LigneDocument.Write()
                                                                End If
                                                            End If
                                                        End If
                                                    End If

                                                    'Date Court
                                                    If statistTab.Rows(0).Item("CB_Type") = 3 Then
                                                        If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                                            OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                                            OleRecherDataset = New DataSet
                                                            OleRecherAdapter.Fill(OleRecherDataset)
                                                            OleRechDatable = OleRecherDataset.Tables(0)
                                                            If OleRechDatable.Rows.Count <> 0 Then
                                                                If Verificatdate(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then
                                                                    LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = RenvoieDateValide(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier)
                                                                    LigneDocument.Write()
                                                                End If
                                                            Else
                                                                If Verificatdate(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then
                                                                    LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = RenvoieDateValide(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), FormatDatefichier)
                                                                    LigneDocument.Write()
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                    'Date Longue
                                                    If statistTab.Rows(0).Item("CB_Type") = 14 Then
                                                        If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                                            OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                                            OleRecherDataset = New DataSet
                                                            OleRecherAdapter.Fill(OleRecherDataset)
                                                            OleRechDatable = OleRecherDataset.Tables(0)
                                                            If OleRechDatable.Rows.Count <> 0 Then
                                                                If Verificatdate(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then
                                                                    LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = RenvoieDateValide(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier)
                                                                    LigneDocument.Write()
                                                                End If
                                                            Else
                                                                If Verificatdate(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then
                                                                    LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = RenvoieDateValide(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), FormatDatefichier)
                                                                    LigneDocument.Write()
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Else
                                                    If IsNothing(LigneDocument.Article) = False Then
                                                        ErreurJrn.WriteLine("Code Article : " & Trim(LigneDocument.Article.AR_Ref) & " Impossible de traiter l'information libre :" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "  Il est inexistant dans la table CbsysLibre de Sage")
                                                    Else
                                                        ErreurJrn.WriteLine("Code Article : " & Trim("Vide") & " Impossible de traiter l'information libre :" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "  Il est inexistant dans la table CbsysLibre de Sage")
                                                    End If
                                                End If
                                            End If
                                        Else
                                            'nothing
                                        End If
                                    End If
                                End If
                                'L'info Libre Parametrée par l'utilisateur n'existe pas dans Sage
                                infoLigne.RemoveAt(0)
                            End While
                        End If
                    Catch ex As Exception
                        exceptionTrouve = True
                        If IsNothing(LigneDocument.Article) = False Then
                            ErreurJrn.WriteLine("Code Article : " & Trim(LigneDocument.Article.AR_Ref) & " Erreur de Création de L'information Libre , vérifiez la longueur de la chaine avec celle paramétrée/ ou la cohérence des données " & ex.Message)
                        Else
                            ErreurJrn.WriteLine("Code Article : " & Trim("Vide") & " Erreur de Création de L'information Libre , vérifiez la longueur de la chaine avec celle paramétrée/ ou la cohérence des données " & ex.Message)
                        End If
                    End Try
                Else
                    ErreurJrn.WriteLine("Le Stock dans le dépôt d'origine   ne permet pas de transférer l'article : " & LigneCodeArticle & " ,  N° :" & Trim(EntetePieceInterne))
                End If
            End With
        Catch ex1 As Exception
            exceptionTrouve = True
            ErreurJrn.WriteLine("Code Article : " & Trim(LigneCodeArticle) & " N°Pièce : " & EntetePieceInterne & " Erreur système de Création de l'article : " & ex1.Message)
            ListePiece.Add(EntetePieceInterne)
        End Try
        Creation_Ligne_ArticleSaisieO = ExisteQuantite
    End Function
    Private Sub CreationTouslesTiers(ByRef EnteteIntituleDepotOrigine As String, ByRef EntetePieceInterne As String, ByRef EnteteTyPeDocument As String, ByRef Document As IBODocumentStock3, ByRef infoListe As List(Of Integer), ByRef FormatDatefichier As String, ByRef DocumentPieceCreation As Object, ByRef PieceAutomat As Object, ByRef PieceArticle As String, ByRef Punitaire As String, ByRef IdentifiantArticle As String, ByRef TypeImport As Object)
        If Trim(EntetePieceInterne) = Trim(EntetePiecePrecedent) Then
            If IsNothing(Document) = True Then
                If Trim(EnteteTyPeDocument) = "23" Then
                    Try
                        Document = BaseCial.FactoryDocumentStock.CreateType(DocumentType.DocumentTypeStockVirement)
                        Creation_Entete_Document(EnteteTyPeDocument, FormatDatefichier, DocumentPieceCreation, EntetePieceInterne, PieceAutomat)
                    Catch ex As Exception
                        exceptionTrouve = True
                        ErreurJrn.WriteLine("Erreur de Création Entête du Mouvement de Transfert N°Pièce Fchier : " & EntetePieceInterne & " Erreur système : " & ex.Message)
                        ListePiece.Add(EntetePieceInterne)
                    End Try
                End If
                If IsNothing(Document) = False Then
                    'Création Ligne du document
                    If TypeImport = "Import" Then
                        Creation_Ligne_Article(FormatDatefichier, PieceArticle, Punitaire, IdentifiantArticle, TypeImport)
                    Else
                        If Creation_Ligne_ArticleSaisieO(FormatDatefichier, PieceArticle, Punitaire, IdentifiantArticle, TypeImport) = True Then
                            Creation_Ligne_ArticleSaisieD(FormatDatefichier, PieceArticle, Punitaire, IdentifiantArticle, TypeImport)
                        End If
                    End If
                End If

            Else
                'Création Ligne Document piece precedente = piece en cours et document existe
                If TypeImport = "Import" Then
                    Creation_Ligne_Article(FormatDatefichier, PieceArticle, Punitaire, IdentifiantArticle, TypeImport)
                Else
                    If Creation_Ligne_ArticleSaisieO(FormatDatefichier, PieceArticle, Punitaire, IdentifiantArticle, TypeImport) = True Then
                        Creation_Ligne_ArticleSaisieD(FormatDatefichier, PieceArticle, Punitaire, IdentifiantArticle, TypeImport)
                    End If
                End If
            End If
        Else
            ' piece precedent <> piece en cours
            Document = Nothing
            If IsNothing(Document) = True Then
                If Trim(EnteteTyPeDocument) = "23" Then
                    Try
                        Document = BaseCial.FactoryDocumentStock.CreateType(DocumentType.DocumentTypeStockVirement)
                        Creation_Entete_Document(EnteteTyPeDocument, FormatDatefichier, DocumentPieceCreation, EntetePieceInterne, PieceAutomat)
                    Catch ex As Exception
                        exceptionTrouve = True
                        ErreurJrn.WriteLine("Erreur de Création Entête du Mouvement de Transfert N°Pièce Fchier : " & EntetePieceInterne & " Erreur système : " & ex.Message)
                        ListePiece.Add(EntetePieceInterne)
                    End Try
                End If
                If IsNothing(Document) = False Then
                    'Création Ligne Document
                    If TypeImport = "Import" Then
                        Creation_Ligne_Article(FormatDatefichier, PieceArticle, Punitaire, IdentifiantArticle, TypeImport)
                    Else
                        If Creation_Ligne_ArticleSaisieO(FormatDatefichier, PieceArticle, Punitaire, IdentifiantArticle, TypeImport) = True Then
                            Creation_Ligne_ArticleSaisieD(FormatDatefichier, PieceArticle, Punitaire, IdentifiantArticle, TypeImport)
                        End If
                    End If
                End If
            End If
        End If
        EntetePiecePrecedent = EntetePieceInterne
    End Sub
    Private Function Verification_Integration_Fichier(ByVal sPathFilexporter As String, ByVal spathFileFormat As String, ByRef Formatype As String, ByRef Base_Excel As String, ByRef sColumnsSepar As String, ByRef FormatdeDatefich As String, ByRef PieceCreation As Object, ByRef PieceAutomatique As Object, ByRef TypeImport As Object) As Boolean
        Dim OleAdaptater As OleDbDataAdapter
        Dim OleAfficheDataset As DataSet
        Dim Oledatable As DataTable
        Dim m As Integer
        Dim jColD As Integer
        Dim iLine As Integer
        Dim aRows() As String
        Dim iColPosition, iColGauchetxt As Integer
        Dim i As Integer, aCols() As String
        Initialiser()
        iLine = 0
        aRows = Nothing
        ListeStock = New List(Of String)
        If Trim(Formatype) = "Excel" Then
            Try
                If AffichFormatintegration(spathFileFormat, Formatype) = True Then
                    ProgressBar1.Value = ProgressBar1.Minimum
                    Datagridaffiche.Rows.Clear()
                    NbreTotal = DecFormat
                    OleAdaptater = New OleDbDataAdapter("select * from [" & Base_Excel & "$] ", OleExcelConnected)
                    OleAfficheDataset = New DataSet
                    OleAdaptater.Fill(OleAfficheDataset)
                    Oledatable = OleAfficheDataset.Tables(0)
                    If Oledatable.Rows.Count <> 0 Then
                        ProgresMax = Oledatable.Rows.Count - DecFormat
                        m = 0
                        infoListe = New List(Of Integer)
                        infoLigne = New List(Of Integer)
                        For i = DecFormat To Oledatable.Rows.Count - 1
                            Datagridaffiche.RowCount = iLine + 1 - m
                            For jColD = 0 To Datagridaffiche.ColumnCount - 1
                                iColPosition = CInt(Strings.Right(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1), Strings.Len(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1)) - InStr(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1), "(")))
                                If iColPosition <= Oledatable.Columns.Count Then
                                    If iColPosition <> 0 Then
                                        If Convert.IsDBNull(Oledatable.Rows(i).Item(iColPosition - 1)) = False Then
                                            Datagridaffiche.Item(jColD, iLine - m).Value = Trim(Oledatable.Rows(i).Item(iColPosition - 1))
                                        Else
                                            Datagridaffiche.Item(jColD, iLine - m).Value = ""
                                        End If
                                    Else
                                        Datagridaffiche.Item(jColD, iLine - m).Value = LireFichierFormat(spathFileFormat, Datagridaffiche.Columns(jColD).Name, Formatype)
                                    End If
                                Else
                                    Datagridaffiche.Item(jColD, iLine - m).Value = ""
                                End If
                            Next jColD
                            iLine = iLine + 1
                            If i Mod 10 = 0 Then
                                Me.Refresh()
                                Verification_Integrer_Ecriture(FormatdeDatefich, PieceCreation, PieceAutomatique, Trim(Formatype), LirePUDefaut(spathFileFormat, Trim(Formatype)), LireLigneArticle(spathFileFormat, Trim(Formatype)), TypeImport)
                                m = iLine
                            Else
                                If i = (Oledatable.Rows.Count - 1) Then
                                    Me.Refresh()
                                    Verification_Integrer_Ecriture(FormatdeDatefich, PieceCreation, PieceAutomatique, Trim(Formatype), LirePUDefaut(spathFileFormat, Trim(Formatype)), LireLigneArticle(spathFileFormat, Trim(Formatype)), TypeImport)
                                    m = iLine
                                End If
                            End If
                        Next i
                    End If
                Else
                    ErreurJrn.WriteLine("Impossible d' Integrer le Fichier " & sPathFilexporter & " car Erreur de lecture du Fichier Format " & spathFileFormat)
                End If
            Catch ex As Exception
                exceptionTrouve = True
            End Try
        Else
            If Trim(Formatype) = "Délimité" Or Trim(Formatype) = "Tabulation" Or Trim(Formatype) = "Pipe" Then
                Try
                    If AffichFormatintegration(spathFileFormat, Formatype) = True Then
                        aRows = GetArrayFile(sPathFilexporter, aRows)
                        NbreTotal = DecFormat
                        ProgressBar1.Value = ProgressBar1.Minimum
                        Datagridaffiche.Rows.Clear()
                        ProgresMax = UBound(aRows) + 1 - DecFormat
                        m = 0
                        infoListe = New List(Of Integer)
                        infoLigne = New List(Of Integer)
                        For i = DecFormat To UBound(aRows)
                            aCols = Split(aRows(i), sColumnsSepar)
                            Datagridaffiche.RowCount = iLine + 1 - m
                            For jColD = 0 To Datagridaffiche.ColumnCount - 1
                                iColPosition = CInt(Strings.Right(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1), Strings.Len(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1)) - InStr(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1), "(")))
                                If iColPosition <> 0 Then
                                    If iColPosition <= (UBound(aCols) + 1) Then
                                        Datagridaffiche.Item(jColD, iLine - m).Value = Trim(aCols(iColPosition - 1))
                                    Else
                                        Datagridaffiche.Item(jColD, iLine - m).Value = ""
                                    End If
                                Else
                                    Datagridaffiche.Item(jColD, iLine - m).Value = LireFichierFormat(spathFileFormat, Datagridaffiche.Columns(jColD).Name, Formatype)
                                End If
                            Next jColD
                            iLine = iLine + 1
                            If i Mod 10 = 0 Then
                                Me.Refresh()
                                Verification_Integrer_Ecriture(FormatdeDatefich, PieceCreation, PieceAutomatique, Trim(Formatype), LirePUDefaut(spathFileFormat, Trim(Formatype)), LireLigneArticle(spathFileFormat, Trim(Formatype)), TypeImport)
                                m = iLine
                            Else
                                If i = UBound(aRows) Then
                                    Me.Refresh()
                                    Verification_Integrer_Ecriture(FormatdeDatefich, PieceCreation, PieceAutomatique, Trim(Formatype), LirePUDefaut(spathFileFormat, Trim(Formatype)), LireLigneArticle(spathFileFormat, Trim(Formatype)), TypeImport)
                                    m = iLine
                                End If
                            End If
                        Next i
                    Else
                        ErreurJrn.WriteLine("Impossible d' Integrer le Fichier " & sPathFilexporter & " car Erreur de lecture du Fichier Format " & spathFileFormat)
                    End If
                Catch ex As Exception
                    exceptionTrouve = True
                End Try
            Else
                If Trim(Formatype) = "Longueur Fixe" Then
                    Try
                        If AffichFormatintegration(spathFileFormat, Formatype) = True Then
                            aRows = GetArrayFile(sPathFilexporter, aRows)
                            NbreTotal = DecFormat
                            ProgressBar1.Value = ProgressBar1.Minimum
                            Datagridaffiche.Rows.Clear()
                            ProgresMax = UBound(aRows) + 1 - DecFormat
                            m = 0
                            infoListe = New List(Of Integer)
                            infoLigne = New List(Of Integer)
                            For i = DecFormat To UBound(aRows)
                                Datagridaffiche.RowCount = iLine + 1 - m
                                For jColD = 0 To Datagridaffiche.ColumnCount - 1
                                    iColPosition = CInt(Strings.Left(Strings.Right(Datagridaffiche.Columns(jColD).HeaderText, Strings.Len(Datagridaffiche.Columns(jColD).HeaderText) - InStr(Datagridaffiche.Columns(jColD).HeaderText, "[")), InStr(Strings.Right(Datagridaffiche.Columns(jColD).HeaderText, Strings.Len(Datagridaffiche.Columns(jColD).HeaderText) - InStr(Datagridaffiche.Columns(jColD).HeaderText, "[")), "]") - 1))
                                    iColGauchetxt = CInt(Strings.Right(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1), Strings.Len(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1)) - InStr(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1), "(")))
                                    If iColPosition <> 0 Or iColGauchetxt <> 0 Then
                                        Datagridaffiche.Item(jColD, iLine - m).Value = Trim(Strings.Mid(aRows(i), iColPosition, iColGauchetxt))
                                    Else
                                        Datagridaffiche.Item(jColD, iLine - m).Value = LireFichierFormat(spathFileFormat, Datagridaffiche.Columns(jColD).Name, Formatype)
                                    End If
                                Next jColD
                                iLine = iLine + 1
                                If i Mod 10 = 0 Then
                                    Me.Refresh()
                                    Verification_Integrer_Ecriture(FormatdeDatefich, PieceCreation, PieceAutomatique, Trim(Formatype), LirePUDefaut(spathFileFormat, Trim(Formatype)), LireLigneArticle(spathFileFormat, Trim(Formatype)), TypeImport)
                                    m = iLine
                                Else
                                    If i = UBound(aRows) Then
                                        Me.Refresh()
                                        Verification_Integrer_Ecriture(FormatdeDatefich, PieceCreation, PieceAutomatique, Trim(Formatype), LirePUDefaut(spathFileFormat, Trim(Formatype)), LireLigneArticle(spathFileFormat, Trim(Formatype)), TypeImport)
                                        m = iLine
                                    End If
                                End If
                            Next i
                        Else
                            ErreurJrn.WriteLine("Impossible d' Integrer le Fichier " & sPathFilexporter & " car Erreur de lecture du Fichier Format " & spathFileFormat)
                        End If
                    Catch ex As Exception
                        exceptionTrouve = True
                    End Try
                End If
            End If
        End If
        Verification_Integration_Fichier = ExisteLecture
    End Function
    Private Function TranscodageTypedocument(ByRef vTypedocument As String) As String
        Dim fournisseurAdap As OleDbDataAdapter
        Dim fournisseurDs As DataSet
        Dim fournisseurTab As DataTable
        Try
            fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Typedocument>' and Valeurlue ='" & Join(Split(Trim(vTypedocument), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
            fournisseurDs = New DataSet
            fournisseurAdap.Fill(fournisseurDs)
            fournisseurTab = fournisseurDs.Tables(0)
            If fournisseurTab.Rows.Count <> 0 Then
                Return fournisseurTab.Rows(0).Item("Correspond")
            Else
                Return vTypedocument
            End If
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Private Sub Verification_Integrer_Ecriture(ByRef FormatfichierDate As String, ByRef PieceCreationDocument As Object, ByRef PieceAutoma As Object, ByRef FormatIntegrer As Object, ByRef Punitaire As String, ByRef IdentifiantArticle As String, ByRef TypeImport As Object)
        Me.Cursor = Cursors.WaitCursor
        BT_integrer.Enabled = False
        If Datagridaffiche.RowCount >= 0 Then
            ProgressBar1.Maximum = ProgresMax
            Try
                For numLigne = 0 To Datagridaffiche.RowCount - 1
                    vidage()
                    NbreTotal = NbreTotal + 1
                    Label5.Refresh()
                    Label5.Text = "Vérification des Integrations!"
                    For numColonne = 0 To Datagridaffiche.ColumnCount - 1
                        'Entête Document
                        If Trim(PieceAutoma) = "non" Then
                            If Datagridaffiche.Columns.Contains(PieceCreationDocument) = True Then
                                If Strings.Len(Trim(Datagridaffiche.Rows(numLigne).Cells(PieceCreationDocument).Value)) <= 8 Then
                                    EntetePieceInterne = Formatage_Chaine(Trim(Datagridaffiche.Rows(numLigne).Cells(PieceCreationDocument).Value))
                                Else
                                    EntetePieceInterne = Formatage_Chaine(Strings.Left(Trim(Datagridaffiche.Rows(numLigne).Cells(PieceCreationDocument).Value), 8))
                                    ErreurJrn.WriteLine("N°Pièce du Fichier :" & Trim(Datagridaffiche.Rows(numLigne).Cells(PieceCreationDocument).Value) & "  a été tronquée")
                                End If
                            End If
                        Else
                            EntetePieceInterne = Trim(Datagridaffiche.Rows(numLigne).Cells(PieceCreationDocument).Value)
                        End If

                        If Datagridaffiche.Columns.Contains(IdentifiantArticle) = True Then
                            PieceArticle = Trim(Datagridaffiche.Rows(numLigne).Cells(IdentifiantArticle).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "EnteteTyPeDocument" Then
                            EnteteTyPeDocument = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "EnteteCodeAffaire" Then
                            EnteteCodeAffaire = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "EnteteDateDocument" Then
                            EnteteDateDocument = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "EntetePlanAnalytique" Then
                            EntetePlanAnalytique = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "EnteteReference" Then
                            If Strings.Len(Trim(Datagridaffiche.Item(numColonne, numLigne).Value)) <= 17 Then
                                EnteteReference = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                            Else
                                EnteteReference = Strings.Left(Trim(Datagridaffiche.Item(numColonne, numLigne).Value), 17)
                                ErreurJrn.WriteLine("La Référence en Entête :" & EnteteReference & " de la Pièce : " & Trim(EntetePieceInterne) & " a été tronqué")
                            End If
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "EnteteSoucheDocument" Then
                            EnteteSoucheDocument = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "LigneDatedeFabrication" Then
                            LigneDatedeFabrication = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "LigneDatedePeremption" Then
                            LigneDatedePeremption = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "LigneDesignationArticle" Then
                            If Strings.Len(Trim(Datagridaffiche.Item(numColonne, numLigne).Value)) <= 69 Then
                                LigneDesignationArticle = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                            Else
                                LigneDesignationArticle = Strings.Left(Trim(Datagridaffiche.Item(numColonne, numLigne).Value), 69)
                                ErreurJrn.WriteLine("La Désignation Article :" & LigneDesignationArticle & " de la Pièce : " & Trim(EntetePieceInterne) & " a été tronqué")
                            End If
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "IDDepotEnteteOrigine" Then
                            IDDepotEnteteOrigine = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "IDDepotEnteteDestination" Then
                            IDDepotEnteteDestination = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "EnteteIntituleDepotOrigine" Then
                            EnteteIntituleDepotOrigine = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "EnteteIntituleDepotDestination" Then
                            EnteteIntituleDepotDestination = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If

                        If Datagridaffiche.Columns(numColonne).Name = "IDDepotLigne" Then
                            IDDepotLigne = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "LigneIntituleDepot" Then
                            LigneIntituleDepot = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "LigneNSerieLot" Then
                            If Strings.Len(Trim(Datagridaffiche.Item(numColonne, numLigne).Value)) <= 30 Then
                                LigneNSerieLot = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                            Else
                                LigneNSerieLot = Strings.Left(Trim(Datagridaffiche.Item(numColonne, numLigne).Value), 30)
                                ErreurJrn.WriteLine("Le N°SerieLot :" & LigneNSerieLot & " de la Pièce : " & Trim(EntetePieceInterne) & " a été tronqué")
                            End If
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "LignePoidsBrut" Then
                            LignePoidsBrut = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "LignePoidsNet" Then
                            LignePoidsNet = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "LignePrixUnitaire" Then
                            LignePrixUnitaire = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If

                        If Datagridaffiche.Columns.Contains("LigneQuantite") = True Then
                            LigneQuantite = Trim(Datagridaffiche.Rows(numLigne).Cells("LigneQuantite").Value)
                        End If

                        If Datagridaffiche.Columns(numColonne).Name = "LigneReference" Then
                            If Strings.Len(Trim(Datagridaffiche.Item(numColonne, numLigne).Value)) <= 17 Then
                                LigneReference = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                            Else
                                LigneReference = Strings.Left(Trim(Datagridaffiche.Item(numColonne, numLigne).Value), 17)
                                ErreurJrn.WriteLine("La Référence en Ligne :" & LigneReference & " de la Pièce : " & Trim(EntetePieceInterne) & " a été tronqué")
                            End If
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "LigneCodeArticle" Then
                            LigneCodeArticle = Formatage_Article(Trim(Datagridaffiche.Item(numColonne, numLigne).Value))
                        End If
                        'RECHERCHE DE L'INTITULE DE L'INFO LIBRE
                        If Trim(FormatIntegrer) = "Longueur Fixe" Then
                            Dim InfoTableau() As String = Split(Trim(Strings.Left(Strings.Right(Datagridaffiche.Columns(numColonne).HeaderText, Strings.Len(Datagridaffiche.Columns(numColonne).HeaderText) - InStr(Datagridaffiche.Columns(numColonne).HeaderText, "{")), Strings.Len(Strings.Right(Datagridaffiche.Columns(numColonne).HeaderText, Strings.Len(Datagridaffiche.Columns(numColonne).HeaderText) - InStr(Datagridaffiche.Columns(numColonne).HeaderText, "{"))) - 1)), "-")
                            If Trim(InfoTableau(0)) = "oui" Then
                                If Trim(InfoTableau(1)) = "F_DOCENTETE" Then
                                    infoListe.Add(numColonne)
                                End If
                                If Trim(InfoTableau(1)) = "F_DOCLIGNE" Then
                                    infoLigne.Add(numColonne)
                                End If
                            End If
                        Else
                            Dim InfoTableau() As String = Split(Trim(Strings.Left(Strings.Right(Datagridaffiche.Columns(numColonne).HeaderText, Strings.Len(Datagridaffiche.Columns(numColonne).HeaderText) - InStr(Datagridaffiche.Columns(numColonne).HeaderText, "[")), Strings.Len(Strings.Right(Datagridaffiche.Columns(numColonne).HeaderText, Strings.Len(Datagridaffiche.Columns(numColonne).HeaderText) - InStr(Datagridaffiche.Columns(numColonne).HeaderText, "["))) - 1)), "-")
                            If Trim(InfoTableau(0)) = "oui" Then
                                If Trim(InfoTableau(1)) = "F_DOCENTETE" Then
                                    infoListe.Add(numColonne)
                                End If
                                If Trim(InfoTableau(1)) = "F_DOCLIGNE" Then
                                    infoLigne.Add(numColonne)
                                End If
                            End If
                        End If
                    Next numColonne
                    'Creation Effective du Document Commercial
                    EnteteTyPeDocument = TranscodageTypedocument(EnteteTyPeDocument)
                    If Trim(EnteteTyPeDocument) = "23" Then 'Transfert
                        Verification_Parametrage(EnteteIntituleDepotOrigine, EntetePieceInterne, EnteteTyPeDocument, Document, infoListe, FormatfichierDate, PieceCreationDocument, PieceAutoma, PieceArticle, Punitaire, IdentifiantArticle, TypeImport)
                    Else
                        ExisteLecture = False
                        exceptionTrouve = True
                        ErreurJrn.WriteLine("Le type de document ne correspond à aucune de ces valeurs (23:Transfert)")
                    End If
                    NbreLigne = NbreLigne + 1
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label8.Text = NbreLigne & "/" & ProgresMax
                    Label8.Refresh()
                Next numLigne
            Catch ex As Exception
                ExisteLecture = False
                exceptionTrouve = True
                ErreurJrn.WriteLine("Une erreur s'est produit au moment de la lecture du fichier  : " & Trim(EnteteIntituleDepotOrigine))
            End Try
        End If
        Datagridaffiche.Rows.Clear()
        Me.Cursor = Cursors.Default
        BT_integrer.Enabled = True
    End Sub
    Private Sub Verification_Parametrage(ByRef EnteteIntituleDepotOrigine As String, ByRef EntetePieceInterne As String, ByRef EnteteTyPeDocument As String, ByRef Document As IBODocumentStock3, ByRef infoListe As List(Of Integer), ByRef FormatDatefichier As String, ByRef PieceCreationDocument As Object, ByRef PieceAutomtique As Object, ByRef PieceArticle As String, ByRef Punitaire As String, ByRef IdentifiantArticle As String, ByRef TypeImport As Object)
        Dim fournisseurAdap As OleDbDataAdapter
        Dim fournisseurDs As DataSet
        Dim fournisseurTab As DataTable
        Dim statistAdap As OleDbDataAdapter
        Dim statistDs As DataSet
        Dim statistTab As DataTable
        Dim OleAdaptaterDelete As OleDbDataAdapter
        Dim OleDeleteDataset As DataSet
        Dim OledatableDelete As DataTable
        Dim OleRecherAdapter As OleDbDataAdapter
        Dim OleRecherDataset As DataSet
        Dim OleRechDatable As DataTable
        Dim DossierAdap As OleDbDataAdapter
        Dim DossierDs As DataSet
        Dim DossierTab As DataTable
        Dim FormatQte As Integer = 0
        Dim FormatMnt As Integer = 0
        DossierAdap = New OleDbDataAdapter("select * from P_DOSSIERCIAL", OleSocieteConnect)
        DossierDs = New DataSet
        DossierAdap.Fill(DossierDs)
        DossierTab = DossierDs.Tables(0)
        If DossierTab.Rows.Count <> 0 Then
            If InStr(DossierTab.Rows(0).Item("D_FormatQte"), ",") <> 0 Then
                If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ",")), ".") <> 0 Then
                    FormatQte = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ","))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ",")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ","))), ".")))
                Else
                    FormatQte = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ",")))
                End If
            Else
                If InStr(DossierTab.Rows(0).Item("D_FormatQte"), ".") <> 0 Then
                    If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ".")), ",") <> 0 Then
                        FormatQte = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), "."))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ".")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), "."))), ",")))
                    Else
                        FormatQte = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ".")))
                    End If
                End If
            End If

            If InStr(DossierTab.Rows(0).Item("D_FormatPrix"), ",") <> 0 Then
                If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ",")), ".") <> 0 Then
                    FormatMnt = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ","))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ",")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ","))), ".")))
                Else
                    FormatMnt = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ",")))
                End If
            Else
                If InStr(DossierTab.Rows(0).Item("D_FormatPrix"), ".") <> 0 Then
                    If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ".")), ",") <> 0 Then
                        FormatMnt = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), "."))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ".")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), "."))), ",")))
                    Else
                        FormatMnt = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ".")))
                    End If
                End If
            End If
        End If
        If Trim(PieceAutomtique) = "non" Then

            If EnteteTyPeDocument = "23" Then
                If Datagridaffiche.Columns.Contains("LigneQuantite") = True Then
                    If Trim(LigneQuantite) <> "" Then
                        If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                            If CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)) < 0 Then
                                ExisteLecture = False
                                ErreurJrn.WriteLine("< La Quantité : " & Trim(LigneQuantite) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " ne doit pas être négative >")
                            End If
                        End If
                    End If
                End If

                If BaseCial.FactoryDocumentStock.ExistPiece(DocumentType.DocumentTypeStockVirement, EntetePieceInterne) = True Then
                    ErreurJrn.WriteLine("Mouvement de Transfert N° : " & EntetePieceInterne & " Existe déja ")
                    ExisteLecture = False
                End If
            Else

            End If
        End If
        If Datagridaffiche.Columns.Contains("EnteteTyPeDocument") = True Then
            If Trim(EnteteTyPeDocument) <> "" Then
                If Trim(EnteteTyPeDocument) = "23" Then
                    If Datagridaffiche.Columns.Contains("LigneQuantite") = True Then
                        If Trim(LigneQuantite) <> "" Then
                            If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                                If CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)) < 0 Then
                                    ExisteLecture = False
                                    ErreurJrn.WriteLine("< La Quantité : " & Trim(LigneQuantite) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " ne doit pas être négative >")
                                End If
                            End If
                        End If
                    End If

                Else
                    ErreurJrn.WriteLine("Le statut du document " & EnteteTyPeDocument & " dois être égal à 23:Transfert : " & EntetePieceInterne & " le statut par défaut va être utilisé")
                End If
            End If
        End If

        If Trim(EntetePieceInterne) <> "" Then
        Else
            ErreurJrn.WriteLine("Le N°Pièce du fichier : " & EntetePieceInterne & " ne doit pas être vide ")
            ExisteLecture = False
        End If
        If Datagridaffiche.Columns.Contains("EntetePlanAnalytique") = True Then
            If Datagridaffiche.Columns.Contains("EnteteCodeAffaire") = True Then
                If Trim(EntetePlanAnalytique) <> "" Then
                    If Trim(EnteteCodeAffaire) <> "" Then
                        If BaseCpta.FactoryAnalytique.ExistIntitule(Trim(EntetePlanAnalytique)) = True Then
                            PlanAna = BaseCpta.FactoryAnalytique.ReadIntitule(Trim(EntetePlanAnalytique))
                            If BaseCpta.FactoryCompteA.ExistNumero(PlanAna, Trim(EnteteCodeAffaire)) = True Then
                            Else
                                statistAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Section Analytique>' and Valeurlue ='" & Join(Split(Trim(EnteteCodeAffaire), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                statistDs = New DataSet
                                statistAdap.Fill(statistDs)
                                statistTab = statistDs.Tables(0)
                                If statistTab.Rows.Count <> 0 Then
                                    If BaseCpta.FactoryCompteA.ExistNumero(PlanAna, Trim(statistTab.Rows(0).Item("Correspond"))) = True Then

                                    Else
                                        ExisteLecture = False
                                        ErreurJrn.WriteLine("< Le Code de Section analytique : " & Trim(statistTab.Rows(0).Item("Correspond")) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'existe pas dans Sage >")
                                    End If
                                Else
                                    ExisteLecture = False
                                    ErreurJrn.WriteLine("< Le Code de Section analytique : " & Trim(EnteteCodeAffaire) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'existe pas dans la table de paramétrage et dans Sage >")
                                End If
                            End If
                        Else
                            fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Plan Analytique>' and Valeurlue ='" & Join(Split(Trim(EntetePlanAnalytique), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                            fournisseurDs = New DataSet
                            fournisseurAdap.Fill(fournisseurDs)
                            fournisseurTab = fournisseurDs.Tables(0)
                            If fournisseurTab.Rows.Count > 0 Then
                                If BaseCpta.FactoryAnalytique.ExistIntitule(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                                    PlanAna = BaseCpta.FactoryAnalytique.ReadIntitule(Trim(fournisseurTab.Rows(0).Item("Correspond")))
                                    If BaseCpta.FactoryCompteA.ExistNumero(PlanAna, Trim(EnteteCodeAffaire)) = True Then
                                    Else
                                        statistAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Section Analytique>' and Valeurlue ='" & Join(Split(Trim(EnteteCodeAffaire), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                        statistDs = New DataSet
                                        statistAdap.Fill(statistDs)
                                        statistTab = statistDs.Tables(0)
                                        If statistTab.Rows.Count <> 0 Then
                                            If BaseCpta.FactoryCompteA.ExistNumero(PlanAna, Trim(statistTab.Rows(0).Item("Correspond"))) = True Then

                                            Else
                                                ExisteLecture = False
                                                ErreurJrn.WriteLine("< Le Code de Section analytique : " & Trim(statistTab.Rows(0).Item("Correspond")) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'existe pas dans Sage >")
                                            End If
                                        Else
                                            ExisteLecture = False
                                            ErreurJrn.WriteLine("< Le Code de Section analytique : " & Trim(EnteteCodeAffaire) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'existe pas dans la table de paramétrage et dans Sage >")
                                        End If
                                    End If
                                Else
                                    ExisteLecture = False
                                    ErreurJrn.WriteLine("< Le Code du plan analytique : " & Trim(fournisseurTab.Rows(0).Item("Correspond")) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'existe pas dans Sage >")
                                End If
                            Else
                                ExisteLecture = False
                                ErreurJrn.WriteLine("< Le Code du plan analytique : " & Trim(EntetePlanAnalytique) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'existe pas dans la table de paramétrage et dans Sage >")
                            End If
                        End If
                    End If
                End If
            End If
        End If
        If Datagridaffiche.Columns.Contains("IDDepotEnteteOrigine") = True Then
            If Datagridaffiche.Columns.Contains("LigneCodeArticle") = True Then
                If BaseCial.FactoryArticle.ExistReference(Trim(LigneCodeArticle)) = True Then
                    If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                        DossierAdap = New OleDbDataAdapter("select * from F_ARTSTOCK where DE_No ='" & CInt(Trim(IDDepotEnteteOrigine)) & "' And AR_Ref='" & Trim(LigneCodeArticle) & "'", OleSocieteConnect)
                        DossierDs = New DataSet
                        DossierAdap.Fill(DossierDs)
                        DossierTab = DossierDs.Tables(0)
                        If DossierTab.Rows.Count <> 0 Then
                            If Convert.IsDBNull(DossierTab.Rows(0).Item("AS_QteSto")) = False Then
                                If Convert.IsDBNull(DossierTab.Rows(0).Item("AS_QtePrepa")) = False Then
                                    If (DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) < CDbl(Trim(LigneQuantite)) Then
                                        ExisteLecture = False
                                        ErreurJrn.WriteLine("Le Stock : " & (DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) & " dans le dépôt d'origine : " & Trim(IDDepotEnteteOrigine) & "   ne permet pas de transférer l'article : " & LigneCodeArticle & " de Quantité : " & CDbl(Trim(LigneQuantite)) & " ,  N° :" & Trim(EntetePieceInterne))
                                    Else
                                        If ListeStock.Count <> 0 Then
                                            Dim ExisteArtStock As Boolean = False
                                            For i As Integer = 0 To ListeStock.Count - 1
                                                Dim StockListe() As String = Split(ListeStock.Item(i), ControlChars.Tab)
                                                If StockListe(0) = CInt(Trim(IDDepotEnteteOrigine)) And Trim(StockListe(1)) = Trim(LigneCodeArticle) Then
                                                    ExisteArtStock = True
                                                    If StockListe(2) < CDbl(Trim(LigneQuantite)) Then
                                                        ExisteLecture = False
                                                        ErreurJrn.WriteLine("Le Stock Calculé " & StockListe(2) & " dans le dépôt d'origine : " & StockListe(0) & "   ne permet pas de transférer l'article : " & LigneCodeArticle & " de Quantité : " & CDbl(Trim(LigneQuantite)) & " ,  N° :" & Trim(EntetePieceInterne))
                                                    Else
                                                        ListeStock.RemoveAt(i)
                                                        ListeStock.Add(CInt(Trim(IDDepotEnteteOrigine)) & ControlChars.Tab & Trim(LigneCodeArticle) & ControlChars.Tab & (StockListe(2) - CDbl(Trim(LigneQuantite))))
                                                    End If
                                                End If
                                            Next i
                                            If ExisteArtStock = False Then
                                                ListeStock.Add(CInt(Trim(IDDepotEnteteOrigine)) & ControlChars.Tab & Trim(LigneCodeArticle) & ControlChars.Tab & ((DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) - CDbl(Trim(LigneQuantite))))
                                            End If
                                        Else
                                            ListeStock.Add(CInt(Trim(IDDepotEnteteOrigine)) & ControlChars.Tab & Trim(LigneCodeArticle) & ControlChars.Tab & ((DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) - CDbl(Trim(LigneQuantite))))
                                        End If
                                    End If
                                Else
                                    If DossierTab.Rows(0).Item("AS_QteSto") < CDbl(Trim(LigneQuantite)) Then
                                        ExisteLecture = False
                                        ErreurJrn.WriteLine("Le Stock : " & DossierTab.Rows(0).Item("AS_QteSto") & " dans le dépôt d'origine : " & Trim(IDDepotEnteteOrigine) & "   ne permet pas de transférer l'article : " & LigneCodeArticle & " de Quantité : " & CDbl(Trim(LigneQuantite)) & " ,  N° :" & Trim(EntetePieceInterne))
                                    Else
                                        If ListeStock.Count <> 0 Then
                                            Dim ExisteArtStock As Boolean = False
                                            For i As Integer = 0 To ListeStock.Count - 1
                                                Dim StockListe() As String = Split(ListeStock.Item(i), ControlChars.Tab)
                                                If StockListe(0) = CInt(Trim(IDDepotEnteteOrigine)) And Trim(StockListe(1)) = Trim(LigneCodeArticle) Then
                                                    ExisteArtStock = True
                                                    If StockListe(2) < CDbl(Trim(LigneQuantite)) Then
                                                        ExisteLecture = False
                                                        ErreurJrn.WriteLine("Le Stock Calculé " & StockListe(2) & " dans le dépôt d'origine : " & StockListe(0) & "   ne permet pas de transférer l'article : " & LigneCodeArticle & " de Quantité : " & CDbl(Trim(LigneQuantite)) & " ,   N° :" & Trim(EntetePieceInterne))
                                                    Else
                                                        ListeStock.RemoveAt(i)
                                                        ListeStock.Add(CInt(Trim(IDDepotEnteteOrigine)) & ControlChars.Tab & Trim(LigneCodeArticle) & ControlChars.Tab & (StockListe(2) - CDbl(Trim(LigneQuantite))))
                                                    End If
                                                End If
                                            Next i
                                            If ExisteArtStock = False Then
                                                ListeStock.Add(CInt(Trim(IDDepotEnteteOrigine)) & ControlChars.Tab & Trim(LigneCodeArticle) & ControlChars.Tab & (DossierTab.Rows(0).Item("AS_QteSto") - CDbl(Trim(LigneQuantite))))
                                            End If
                                        Else
                                            ListeStock.Add(CInt(Trim(IDDepotEnteteOrigine)) & ControlChars.Tab & Trim(LigneCodeArticle) & ControlChars.Tab & (DossierTab.Rows(0).Item("AS_QteSto") - CDbl(Trim(LigneQuantite))))
                                        End If
                                    End If
                                End If
                            Else
                                ExisteLecture = False
                                ErreurJrn.WriteLine("Le Stock dans le dépôt d'origine : " & Trim(IDDepotEnteteOrigine) & "  est NULL et ne permet pas de transférer l'article : " & LigneCodeArticle & " ,  N° :" & Trim(EntetePieceInterne))
                            End If
                        Else
                            ExisteLecture = False
                            ErreurJrn.WriteLine("Le Stock dans le dépôt d'origine : " & Trim(IDDepotEnteteOrigine) & "   ne permet pas de transférer l'article : " & LigneCodeArticle & " ,  N° :" & Trim(EntetePieceInterne))
                        End If
                    End If
                Else
                    fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<ArticleStock>' and Valeurlue ='" & Join(Split(Trim(LigneCodeArticle), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                    fournisseurDs = New DataSet
                    fournisseurAdap.Fill(fournisseurDs)
                    fournisseurTab = fournisseurDs.Tables(0)
                    If fournisseurTab.Rows.Count <> 0 Then
                        If BaseCial.FactoryArticle.ExistReference(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                            If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                                DossierAdap = New OleDbDataAdapter("select * from F_ARTSTOCK where DE_No ='" & CInt(Trim(IDDepotEnteteOrigine)) & "' And AR_Ref='" & Trim(fournisseurTab.Rows(0).Item("Correspond")) & "'", OleSocieteConnect)
                                DossierDs = New DataSet
                                DossierAdap.Fill(DossierDs)
                                DossierTab = DossierDs.Tables(0)
                                If DossierTab.Rows.Count <> 0 Then
                                    If Convert.IsDBNull(DossierTab.Rows(0).Item("AS_QteSto")) = False Then
                                        If Convert.IsDBNull(DossierTab.Rows(0).Item("AS_QtePrepa")) = False Then
                                            If (DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) < CDbl(Trim(LigneQuantite)) Then
                                                ExisteLecture = False
                                                ErreurJrn.WriteLine("Le Stock : " & (DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) & "  dans le dépôt d'origine : " & Trim(IDDepotEnteteOrigine) & "   ne permet pas de transférer l'article : " & LigneCodeArticle & " de Quantité : " & CDbl(Trim(LigneQuantite)) & "  ,  N° :" & Trim(EntetePieceInterne))
                                            Else
                                                If ListeStock.Count <> 0 Then
                                                    Dim ExisteArtStock As Boolean = False
                                                    For i As Integer = 0 To ListeStock.Count - 1
                                                        Dim StockListe() As String = Split(ListeStock.Item(i), ControlChars.Tab)
                                                        If StockListe(0) = CInt(Trim(IDDepotEnteteOrigine)) And Trim(StockListe(1)) = Trim(fournisseurTab.Rows(0).Item("Correspond")) Then
                                                            ExisteArtStock = True
                                                            If StockListe(2) < CDbl(Trim(LigneQuantite)) Then
                                                                ExisteLecture = False
                                                                ErreurJrn.WriteLine("Le Stock Calculé " & StockListe(2) & " dans le dépôt d'origine : " & StockListe(0) & "   ne permet pas de transférer l'article : " & Trim(fournisseurTab.Rows(0).Item("Correspond")) & " de Quantité : " & CDbl(Trim(LigneQuantite)) & " ,  N° :" & Trim(EntetePieceInterne))
                                                            Else
                                                                ListeStock.RemoveAt(i)
                                                                ListeStock.Add(CInt(Trim(IDDepotEnteteOrigine)) & ControlChars.Tab & Trim(fournisseurTab.Rows(0).Item("Correspond")) & ControlChars.Tab & (StockListe(2) - CDbl(Trim(LigneQuantite))))
                                                            End If
                                                        End If
                                                    Next i
                                                    If ExisteArtStock = False Then
                                                        ListeStock.Add(CInt(Trim(IDDepotEnteteOrigine)) & ControlChars.Tab & Trim(fournisseurTab.Rows(0).Item("Correspond")) & ControlChars.Tab & ((DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) - CDbl(Trim(LigneQuantite))))
                                                    End If
                                                Else
                                                    ListeStock.Add(CInt(Trim(IDDepotEnteteOrigine)) & ControlChars.Tab & Trim(fournisseurTab.Rows(0).Item("Correspond")) & ControlChars.Tab & ((DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) - CDbl(Trim(LigneQuantite))))
                                                End If
                                            End If
                                        Else
                                            If DossierTab.Rows(0).Item("AS_QteSto") < CDbl(Trim(LigneQuantite)) Then
                                                ExisteLecture = False
                                                ErreurJrn.WriteLine("Le Stock : " & DossierTab.Rows(0).Item("AS_QteSto") & " dans le dépôt d'origine : " & Trim(IDDepotEnteteOrigine) & "   ne permet pas de transférer l'article : " & LigneCodeArticle & " de Quantité : " & CDbl(Trim(LigneQuantite)) & " ,  N° :" & Trim(EntetePieceInterne))
                                            Else
                                                If ListeStock.Count <> 0 Then
                                                    Dim ExisteArtStock As Boolean = False
                                                    For i As Integer = 0 To ListeStock.Count - 1
                                                        Dim StockListe() As String = Split(ListeStock.Item(i), ControlChars.Tab)
                                                        If StockListe(0) = CInt(Trim(IDDepotEnteteOrigine)) And Trim(StockListe(1)) = Trim(fournisseurTab.Rows(0).Item("Correspond")) Then
                                                            ExisteArtStock = True
                                                            If StockListe(2) < CDbl(Trim(LigneQuantite)) Then
                                                                ExisteLecture = False
                                                                ErreurJrn.WriteLine("Le Stock Calculé " & StockListe(2) & " dans le dépôt d'origine : " & StockListe(0) & "   ne permet pas de transférer l'article : " & Trim(fournisseurTab.Rows(0).Item("Correspond")) & " de Quantité : " & CDbl(Trim(LigneQuantite)) & " ,  N° :" & Trim(EntetePieceInterne))
                                                            Else
                                                                ListeStock.RemoveAt(i)
                                                                ListeStock.Add(CInt(Trim(IDDepotEnteteOrigine)) & ControlChars.Tab & Trim(fournisseurTab.Rows(0).Item("Correspond")) & ControlChars.Tab & (StockListe(2) - CDbl(Trim(LigneQuantite))))
                                                            End If
                                                        End If
                                                    Next i
                                                    If ExisteArtStock = False Then
                                                        ListeStock.Add(CInt(Trim(IDDepotEnteteOrigine)) & ControlChars.Tab & Trim(fournisseurTab.Rows(0).Item("Correspond")) & ControlChars.Tab & (DossierTab.Rows(0).Item("AS_QteSto") - CDbl(Trim(LigneQuantite))))
                                                    End If
                                                Else
                                                    ListeStock.Add(CInt(Trim(IDDepotEnteteOrigine)) & ControlChars.Tab & Trim(fournisseurTab.Rows(0).Item("Correspond")) & ControlChars.Tab & (DossierTab.Rows(0).Item("AS_QteSto") - CDbl(Trim(LigneQuantite))))
                                                End If
                                            End If
                                        End If
                                    Else
                                        ExisteLecture = False
                                        ErreurJrn.WriteLine("Le Stock dans le dépôt d'origine : " & Trim(IDDepotEnteteOrigine) & "  est NULL et ne permet pas de transférer l'article : " & LigneCodeArticle & " ,  N° :" & Trim(EntetePieceInterne))
                                    End If
                                Else
                                    ExisteLecture = False
                                    ErreurJrn.WriteLine("Le Stock dans le dépôt d'origine : " & Trim(IDDepotEnteteOrigine) & "   ne permet pas de transférer l'article : " & LigneCodeArticle & " ,  N° :" & Trim(EntetePieceInterne))
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If

        If Datagridaffiche.Columns.Contains("LigneCodeArticle") = True Then
            If BaseCial.FactoryArticle.ExistReference(Trim(LigneCodeArticle)) = True Then
                If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                    If Datagridaffiche.Columns.Contains("EnteteIntituleDepotOrigine") = True Then
                        If BaseCial.FactoryDepot.ExistIntitule(Trim(EnteteIntituleDepotOrigine)) = True Then
                            DossierAdap = New OleDbDataAdapter("select * from F_ARTSTOCK where DE_No =(Select DE_No from F_DEPOT WHERE DE_Intitule='" & Trim(EnteteIntituleDepotOrigine) & "') And AR_Ref='" & Trim(LigneCodeArticle) & "'", OleSocieteConnect)
                            DossierDs = New DataSet
                            DossierAdap.Fill(DossierDs)
                            DossierTab = DossierDs.Tables(0)
                            If DossierTab.Rows.Count <> 0 Then
                                If Convert.IsDBNull(DossierTab.Rows(0).Item("AS_QteSto")) = False Then
                                    If Convert.IsDBNull(DossierTab.Rows(0).Item("AS_QtePrepa")) = False Then
                                        If (DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) < CDbl(Trim(LigneQuantite)) Then
                                            ExisteLecture = False
                                            ErreurJrn.WriteLine("Le Stock : " & (DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) & " dans le dépôt d'origine : " & Trim(EnteteIntituleDepotOrigine) & "   ne permet pas de transférer l'article : " & LigneCodeArticle & " de Quantité : " & CDbl(Trim(LigneQuantite)) & " ,  N° :" & Trim(EntetePieceInterne))
                                        Else
                                            If ListeStock.Count <> 0 Then
                                                Dim ExisteArtStock As Boolean = False
                                                For i As Integer = 0 To ListeStock.Count - 1
                                                    Dim StockListe() As String = Split(ListeStock.Item(i), ControlChars.Tab)
                                                    If StockListe(0) = Trim(EnteteIntituleDepotOrigine) And Trim(StockListe(1)) = Trim(LigneCodeArticle) Then
                                                        ExisteArtStock = True
                                                        If StockListe(2) < CDbl(Trim(LigneQuantite)) Then
                                                            ExisteLecture = False
                                                            ErreurJrn.WriteLine("Le Stock Calculé " & StockListe(2) & " dans le dépôt d'origine : " & StockListe(0) & "   ne permet pas de transférer l'article : " & LigneCodeArticle & " de Quantité : " & CDbl(Trim(LigneQuantite)) & " ,  N° :" & Trim(EntetePieceInterne))
                                                        Else
                                                            ListeStock.RemoveAt(i)
                                                            ListeStock.Add(Trim(EnteteIntituleDepotOrigine) & ControlChars.Tab & Trim(LigneCodeArticle) & ControlChars.Tab & (StockListe(2) - CDbl(Trim(LigneQuantite))))
                                                        End If
                                                    End If
                                                Next i
                                                If ExisteArtStock = False Then
                                                    ListeStock.Add(Trim(EnteteIntituleDepotOrigine) & ControlChars.Tab & Trim(LigneCodeArticle) & ControlChars.Tab & ((DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) - CDbl(Trim(LigneQuantite))))
                                                End If
                                            Else
                                                ListeStock.Add(Trim(EnteteIntituleDepotOrigine) & ControlChars.Tab & Trim(LigneCodeArticle) & ControlChars.Tab & ((DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) - CDbl(Trim(LigneQuantite))))
                                            End If
                                        End If
                                    Else
                                        If DossierTab.Rows(0).Item("AS_QteSto") < CDbl(Trim(LigneQuantite)) Then
                                            ExisteLecture = False
                                            ErreurJrn.WriteLine("Le Stock : " & DossierTab.Rows(0).Item("AS_QteSto") & " dans le dépôt d'origine : " & Trim(EnteteIntituleDepotOrigine) & "   ne permet pas de transférer l'article : " & LigneCodeArticle & " de Quantité : " & CDbl(Trim(LigneQuantite)) & " ,  N° :" & Trim(EntetePieceInterne))
                                        Else
                                            If ListeStock.Count <> 0 Then
                                                Dim ExisteArtStock As Boolean = False
                                                For i As Integer = 0 To ListeStock.Count - 1
                                                    Dim StockListe() As String = Split(ListeStock.Item(i), ControlChars.Tab)
                                                    If StockListe(0) = Trim(EnteteIntituleDepotOrigine) And Trim(StockListe(1)) = Trim(LigneCodeArticle) Then
                                                        ExisteArtStock = True
                                                        If StockListe(2) < CDbl(Trim(LigneQuantite)) Then
                                                            ExisteLecture = False
                                                            ErreurJrn.WriteLine("Le Stock Calculé " & StockListe(2) & " dans le dépôt d'origine : " & StockListe(0) & "   ne permet pas de transférer l'article : " & LigneCodeArticle & " de Quantité : " & CDbl(Trim(LigneQuantite)) & " ,  N° :" & Trim(EntetePieceInterne))
                                                        Else
                                                            ListeStock.RemoveAt(i)
                                                            ListeStock.Add(Trim(EnteteIntituleDepotOrigine) & ControlChars.Tab & Trim(LigneCodeArticle) & ControlChars.Tab & (StockListe(2) - CDbl(Trim(LigneQuantite))))
                                                        End If
                                                    End If
                                                Next i
                                                If ExisteArtStock = False Then
                                                    ListeStock.Add(Trim(EnteteIntituleDepotOrigine) & ControlChars.Tab & Trim(LigneCodeArticle) & ControlChars.Tab & (DossierTab.Rows(0).Item("AS_QteSto") - CDbl(Trim(LigneQuantite))))
                                                End If
                                            Else
                                                ListeStock.Add(Trim(EnteteIntituleDepotOrigine) & ControlChars.Tab & Trim(LigneCodeArticle) & ControlChars.Tab & (DossierTab.Rows(0).Item("AS_QteSto") - CDbl(Trim(LigneQuantite))))
                                            End If
                                        End If
                                    End If
                                Else
                                    ExisteLecture = False
                                    ErreurJrn.WriteLine("Le Stock dans le dépôt d'origine : " & Trim(EnteteIntituleDepotOrigine) & " est NULL et  ne permet pas de transférer l'article : " & LigneCodeArticle & " ,  N° :" & Trim(EntetePieceInterne))
                                End If
                            Else
                                ExisteLecture = False
                                ErreurJrn.WriteLine("Le Stock dans le dépôt d'origine : " & Trim(EnteteIntituleDepotOrigine) & "   ne permet pas de transférer l'article : " & LigneCodeArticle & " ,  N° :" & Trim(EntetePieceInterne))
                            End If
                        Else
                            fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Depot Stockage>' and Valeurlue ='" & Join(Split(Trim(EnteteIntituleDepotOrigine), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                            fournisseurDs = New DataSet
                            fournisseurAdap.Fill(fournisseurDs)
                            fournisseurTab = fournisseurDs.Tables(0)
                            If fournisseurTab.Rows.Count > 0 Then
                                If BaseCial.FactoryDepot.ExistIntitule(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                                    DossierAdap = New OleDbDataAdapter("select * from F_ARTSTOCK where DE_No =(Select DE_No from F_DEPOT WHERE DE_Intitule='" & Trim(fournisseurTab.Rows(0).Item("Correspond")) & "') And AR_Ref='" & Trim(LigneCodeArticle) & "'", OleSocieteConnect)
                                    DossierDs = New DataSet
                                    DossierAdap.Fill(DossierDs)
                                    DossierTab = DossierDs.Tables(0)
                                    If DossierTab.Rows.Count <> 0 Then
                                        If Convert.IsDBNull(DossierTab.Rows(0).Item("AS_QteSto")) = False Then
                                            If Convert.IsDBNull(DossierTab.Rows(0).Item("AS_QtePrepa")) = False Then
                                                If (DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) < CDbl(Trim(LigneQuantite)) Then
                                                    ExisteLecture = False
                                                    ErreurJrn.WriteLine("Le Stock : " & (DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) & " dans le dépôt d'origine : " & Trim(fournisseurTab.Rows(0).Item("Correspond")) & "   ne permet pas de transférer l'article : " & LigneCodeArticle & " de Quantité : " & CDbl(Trim(LigneQuantite)) & " ,  N° :" & Trim(EntetePieceInterne))
                                                Else
                                                    If ListeStock.Count <> 0 Then
                                                        Dim ExisteArtStock As Boolean = False
                                                        For i As Integer = 0 To ListeStock.Count - 1
                                                            Dim StockListe() As String = Split(ListeStock.Item(i), ControlChars.Tab)
                                                            If Trim(StockListe(0)) = Trim(fournisseurTab.Rows(0).Item("Correspond")) And Trim(StockListe(1)) = Trim(LigneCodeArticle) Then
                                                                ExisteArtStock = True
                                                                If StockListe(2) < CDbl(Trim(LigneQuantite)) Then
                                                                    ExisteLecture = False
                                                                    ErreurJrn.WriteLine("Le Stock Calculé " & StockListe(2) & " dans le dépôt d'origine : " & StockListe(0) & "   ne permet pas de transférer l'article : " & Trim(LigneCodeArticle) & " de Quantité : " & CDbl(Trim(LigneQuantite)) & " ,  N° :" & Trim(EntetePieceInterne))
                                                                Else
                                                                    ListeStock.RemoveAt(i)
                                                                    ListeStock.Add(Trim(fournisseurTab.Rows(0).Item("Correspond")) & ControlChars.Tab & Trim(LigneCodeArticle) & ControlChars.Tab & (StockListe(2) - CDbl(Trim(LigneQuantite))))
                                                                End If
                                                            End If
                                                        Next i
                                                        If ExisteArtStock = False Then
                                                            ListeStock.Add(Trim(fournisseurTab.Rows(0).Item("Correspond")) & ControlChars.Tab & Trim(LigneCodeArticle) & ControlChars.Tab & ((DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) - CDbl(Trim(LigneQuantite))))
                                                        End If
                                                    Else
                                                        ListeStock.Add(Trim(fournisseurTab.Rows(0).Item("Correspond")) & ControlChars.Tab & Trim(LigneCodeArticle) & ControlChars.Tab & ((DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) - CDbl(Trim(LigneQuantite))))
                                                    End If
                                                End If
                                            Else
                                                If DossierTab.Rows(0).Item("AS_QteSto") < CDbl(Trim(LigneQuantite)) Then
                                                    ExisteLecture = False
                                                    ErreurJrn.WriteLine("Le Stock : " & DossierTab.Rows(0).Item("AS_QteSto") & " dans le dépôt d'origine : " & Trim(fournisseurTab.Rows(0).Item("Correspond")) & "   ne permet pas de transférer l'article : " & LigneCodeArticle & " de Quantité : " & CDbl(Trim(LigneQuantite)) & " ,  N° :" & Trim(EntetePieceInterne))
                                                Else
                                                    If ListeStock.Count <> 0 Then
                                                        Dim ExisteArtStock As Boolean = False
                                                        For i As Integer = 0 To ListeStock.Count - 1
                                                            Dim StockListe() As String = Split(ListeStock.Item(i), ControlChars.Tab)
                                                            If Trim(StockListe(0)) = Trim(fournisseurTab.Rows(0).Item("Correspond")) And Trim(StockListe(1)) = Trim(LigneCodeArticle) Then
                                                                ExisteArtStock = True
                                                                If StockListe(2) < CDbl(Trim(LigneQuantite)) Then
                                                                    ExisteLecture = False
                                                                    ErreurJrn.WriteLine("Le Stock Calculé " & StockListe(2) & " dans le dépôt d'origine : " & StockListe(0) & "   ne permet pas de transférer l'article : " & Trim(LigneCodeArticle) & " de Quantité : " & CDbl(Trim(LigneQuantite)) & " ,  N° :" & Trim(EntetePieceInterne))
                                                                Else
                                                                    ListeStock.RemoveAt(i)
                                                                    ListeStock.Add(Trim(fournisseurTab.Rows(0).Item("Correspond")) & ControlChars.Tab & Trim(LigneCodeArticle) & ControlChars.Tab & (StockListe(2) - CDbl(Trim(LigneQuantite))))
                                                                End If
                                                            End If
                                                        Next i
                                                        If ExisteArtStock = False Then
                                                            ListeStock.Add(Trim(fournisseurTab.Rows(0).Item("Correspond")) & ControlChars.Tab & Trim(LigneCodeArticle) & ControlChars.Tab & (DossierTab.Rows(0).Item("AS_QteSto") - CDbl(Trim(LigneQuantite))))
                                                        End If
                                                    Else
                                                        ListeStock.Add(Trim(fournisseurTab.Rows(0).Item("Correspond")) & ControlChars.Tab & Trim(LigneCodeArticle) & ControlChars.Tab & (DossierTab.Rows(0).Item("AS_QteSto") - CDbl(Trim(LigneQuantite))))
                                                    End If
                                                End If
                                            End If
                                        Else
                                            ExisteLecture = False
                                            ErreurJrn.WriteLine("Le Stock dans le dépôt d'origine : " & Trim(fournisseurTab.Rows(0).Item("Correspond")) & " est NULL et  ne permet pas de transférer l'article : " & LigneCodeArticle & " ,  N° :" & Trim(EntetePieceInterne))
                                        End If
                                    Else
                                        ExisteLecture = False
                                        ErreurJrn.WriteLine("Le Stock dans le dépôt d'origine : " & Trim(fournisseurTab.Rows(0).Item("Correspond")) & "   ne permet pas de transférer l'article : " & LigneCodeArticle & " ,  N° :" & Trim(EntetePieceInterne))
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<ArticleStock>' and Valeurlue ='" & Join(Split(Trim(LigneCodeArticle), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                fournisseurDs = New DataSet
                fournisseurAdap.Fill(fournisseurDs)
                fournisseurTab = fournisseurDs.Tables(0)
                If fournisseurTab.Rows.Count <> 0 Then
                    If BaseCial.FactoryArticle.ExistReference(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                        If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                            If Datagridaffiche.Columns.Contains("EnteteIntituleDepotOrigine") = True Then
                                If BaseCial.FactoryDepot.ExistIntitule(Trim(EnteteIntituleDepotOrigine)) = True Then
                                    DossierAdap = New OleDbDataAdapter("select * from F_ARTSTOCK where DE_No =(Select DE_No from F_DEPOT WHERE DE_Intitule='" & Trim(EnteteIntituleDepotOrigine) & "') And AR_Ref='" & Trim(fournisseurTab.Rows(0).Item("Correspond")) & "'", OleSocieteConnect)
                                    DossierDs = New DataSet
                                    DossierAdap.Fill(DossierDs)
                                    DossierTab = DossierDs.Tables(0)
                                    If DossierTab.Rows.Count <> 0 Then
                                        If Convert.IsDBNull(DossierTab.Rows(0).Item("AS_QteSto")) = False Then
                                            If Convert.IsDBNull(DossierTab.Rows(0).Item("AS_QtePrepa")) = False Then
                                                If (DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) < CDbl(Trim(LigneQuantite)) Then
                                                    ExisteLecture = False
                                                    ErreurJrn.WriteLine("Le Stock : " & (DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) & " dans le dépôt d'origine : " & Trim(EnteteIntituleDepotOrigine) & "   ne permet pas de transférer l'article : " & LigneCodeArticle & " de Quantité : " & CDbl(Trim(LigneQuantite)) & " ,  N° :" & Trim(EntetePieceInterne))
                                                Else
                                                    If ListeStock.Count <> 0 Then
                                                        Dim ExisteArtStock As Boolean = False
                                                        For i As Integer = 0 To ListeStock.Count - 1
                                                            Dim StockListe() As String = Split(ListeStock.Item(i), ControlChars.Tab)
                                                            If Trim(StockListe(0)) = Trim(EnteteIntituleDepotOrigine) And Trim(StockListe(1)) = Trim(fournisseurTab.Rows(0).Item("Correspond")) Then
                                                                ExisteArtStock = True
                                                                If StockListe(2) < CDbl(Trim(LigneQuantite)) Then
                                                                    ExisteLecture = False
                                                                    ErreurJrn.WriteLine("Le Stock Calculé " & StockListe(2) & " dans le dépôt d'origine : " & StockListe(0) & "   ne permet pas de transférer l'article : " & Trim(fournisseurTab.Rows(0).Item("Correspond")) & " de Quantité : " & CDbl(Trim(LigneQuantite)) & " ,  N° :" & Trim(EntetePieceInterne))
                                                                Else
                                                                    ListeStock.RemoveAt(i)
                                                                    ListeStock.Add(Trim(EnteteIntituleDepotOrigine) & ControlChars.Tab & Trim(fournisseurTab.Rows(0).Item("Correspond")) & ControlChars.Tab & (StockListe(2) - CDbl(Trim(LigneQuantite))))
                                                                End If
                                                            End If
                                                        Next i
                                                        If ExisteArtStock = False Then
                                                            ListeStock.Add(Trim(EnteteIntituleDepotOrigine) & ControlChars.Tab & Trim(fournisseurTab.Rows(0).Item("Correspond")) & ControlChars.Tab & ((DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) - CDbl(Trim(LigneQuantite))))
                                                        End If
                                                    Else
                                                        ListeStock.Add(Trim(EnteteIntituleDepotOrigine) & ControlChars.Tab & Trim(fournisseurTab.Rows(0).Item("Correspond")) & ControlChars.Tab & ((DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) - CDbl(Trim(LigneQuantite))))
                                                    End If
                                                End If
                                            Else
                                                If DossierTab.Rows(0).Item("AS_QteSto") < CDbl(Trim(LigneQuantite)) Then
                                                    ExisteLecture = False
                                                    ErreurJrn.WriteLine("Le Stock : " & DossierTab.Rows(0).Item("AS_QteSto") & " dans le dépôt d'origine : " & Trim(EnteteIntituleDepotOrigine) & "   ne permet pas de transférer l'article : " & LigneCodeArticle & " de Quantité : " & CDbl(Trim(LigneQuantite)) & " ,  N° :" & Trim(EntetePieceInterne))
                                                Else
                                                    If ListeStock.Count <> 0 Then
                                                        Dim ExisteArtStock As Boolean = False
                                                        For i As Integer = 0 To ListeStock.Count - 1
                                                            Dim StockListe() As String = Split(ListeStock.Item(i), ControlChars.Tab)
                                                            If Trim(StockListe(0)) = Trim(EnteteIntituleDepotOrigine) And Trim(StockListe(1)) = Trim(fournisseurTab.Rows(0).Item("Correspond")) Then
                                                                ExisteArtStock = True
                                                                If StockListe(2) < CDbl(Trim(LigneQuantite)) Then
                                                                    ExisteLecture = False
                                                                    ErreurJrn.WriteLine("Le Stock Calculé " & StockListe(2) & " dans le dépôt d'origine : " & StockListe(0) & "   ne permet pas de transférer l'article : " & Trim(fournisseurTab.Rows(0).Item("Correspond")) & " de Quantité : " & CDbl(Trim(LigneQuantite)) & " ,  N° :" & Trim(EntetePieceInterne))
                                                                Else
                                                                    ListeStock.RemoveAt(i)
                                                                    ListeStock.Add(Trim(EnteteIntituleDepotOrigine) & ControlChars.Tab & Trim(fournisseurTab.Rows(0).Item("Correspond")) & ControlChars.Tab & (StockListe(2) - CDbl(Trim(LigneQuantite))))
                                                                End If
                                                            End If
                                                        Next i
                                                        If ExisteArtStock = False Then
                                                            ListeStock.Add(Trim(EnteteIntituleDepotOrigine) & ControlChars.Tab & Trim(fournisseurTab.Rows(0).Item("Correspond")) & ControlChars.Tab & (DossierTab.Rows(0).Item("AS_QteSto") - CDbl(Trim(LigneQuantite))))
                                                        End If
                                                    Else
                                                        ListeStock.Add(Trim(EnteteIntituleDepotOrigine) & ControlChars.Tab & Trim(fournisseurTab.Rows(0).Item("Correspond")) & ControlChars.Tab & (DossierTab.Rows(0).Item("AS_QteSto") - CDbl(Trim(LigneQuantite))))
                                                    End If
                                                End If
                                            End If
                                        Else
                                            ExisteLecture = False
                                            ErreurJrn.WriteLine("Le Stock dans le dépôt d'origine : " & Trim(EnteteIntituleDepotOrigine) & "  est NULL et ne permet pas de transférer l'article : " & LigneCodeArticle & " ,  N° :" & Trim(EntetePieceInterne))
                                        End If
                                    Else
                                        ExisteLecture = False
                                        ErreurJrn.WriteLine("Le Stock dans le dépôt d'origine : " & Trim(EnteteIntituleDepotOrigine) & "   ne permet pas de transférer l'article : " & LigneCodeArticle & " ,  N° :" & Trim(EntetePieceInterne))
                                    End If
                                Else
                                    Dim DepotAdap As OleDbDataAdapter
                                    Dim DepotDs As DataSet
                                    Dim DepotTab As DataTable
                                    DepotAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Depot Stockage>' and Valeurlue ='" & Join(Split(Trim(EnteteIntituleDepotOrigine), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                    DepotDs = New DataSet
                                    DepotAdap.Fill(DepotDs)
                                    DepotTab = DepotDs.Tables(0)
                                    If DepotTab.Rows.Count > 0 Then
                                        If BaseCial.FactoryDepot.ExistIntitule(Trim(DepotTab.Rows(0).Item("Correspond"))) = True Then
                                            DossierAdap = New OleDbDataAdapter("select * from F_ARTSTOCK where DE_No =(Select DE_No from F_DEPOT WHERE DE_Intitule='" & Trim(DepotTab.Rows(0).Item("Correspond")) & "') And AR_Ref='" & Trim(fournisseurTab.Rows(0).Item("Correspond")) & "'", OleSocieteConnect)
                                            DossierDs = New DataSet
                                            DossierAdap.Fill(DossierDs)
                                            DossierTab = DossierDs.Tables(0)
                                            If DossierTab.Rows.Count <> 0 Then
                                                If Convert.IsDBNull(DossierTab.Rows(0).Item("AS_QteSto")) = False Then
                                                    If Convert.IsDBNull(DossierTab.Rows(0).Item("AS_QtePrepa")) = False Then
                                                        If (DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) < CDbl(Trim(LigneQuantite)) Then
                                                            ExisteLecture = False
                                                            ErreurJrn.WriteLine("Le Stock : " & (DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) & " dans le dépôt d'origine : " & Trim(DepotTab.Rows(0).Item("Correspond")) & "   ne permet pas de transférer l'article : " & LigneCodeArticle & " de Quantité : " & CDbl(Trim(LigneQuantite)) & " ,  N° :" & Trim(EntetePieceInterne))
                                                        Else
                                                            If ListeStock.Count <> 0 Then
                                                                Dim ExisteArtStock As Boolean = False
                                                                For i As Integer = 0 To ListeStock.Count - 1
                                                                    Dim StockListe() As String = Split(ListeStock.Item(i), ControlChars.Tab)
                                                                    If Trim(StockListe(0)) = Trim(DepotTab.Rows(0).Item("Correspond")) And Trim(StockListe(1)) = Trim(fournisseurTab.Rows(0).Item("Correspond")) Then
                                                                        ExisteArtStock = True
                                                                        If StockListe(2) < CDbl(Trim(LigneQuantite)) Then
                                                                            ExisteLecture = False
                                                                            ErreurJrn.WriteLine("Le Stock Calculé " & StockListe(2) & " dans le dépôt d'origine : " & StockListe(0) & "   ne permet pas de transférer l'article : " & Trim(fournisseurTab.Rows(0).Item("Correspond")) & " de Quantité : " & CDbl(Trim(LigneQuantite)) & " ,  N° :" & Trim(EntetePieceInterne))
                                                                        Else
                                                                            ListeStock.RemoveAt(i)
                                                                            ListeStock.Add(Trim(DepotTab.Rows(0).Item("Correspond")) & ControlChars.Tab & Trim(fournisseurTab.Rows(0).Item("Correspond")) & ControlChars.Tab & (StockListe(2) - CDbl(Trim(LigneQuantite))))
                                                                        End If
                                                                    End If
                                                                Next i
                                                                If ExisteArtStock = False Then
                                                                    ListeStock.Add(Trim(DepotTab.Rows(0).Item("Correspond")) & ControlChars.Tab & Trim(fournisseurTab.Rows(0).Item("Correspond")) & ControlChars.Tab & ((DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) - CDbl(Trim(LigneQuantite))))
                                                                End If
                                                            Else
                                                                ListeStock.Add(Trim(DepotTab.Rows(0).Item("Correspond")) & ControlChars.Tab & Trim(fournisseurTab.Rows(0).Item("Correspond")) & ControlChars.Tab & ((DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) - CDbl(Trim(LigneQuantite))))
                                                            End If
                                                        End If
                                                    Else
                                                        If DossierTab.Rows(0).Item("AS_QteSto") < CDbl(Trim(LigneQuantite)) Then
                                                            ExisteLecture = False
                                                            ErreurJrn.WriteLine("Le Stock : " & DossierTab.Rows(0).Item("AS_QteSto") & " dans le dépôt d'origine : " & Trim(DepotTab.Rows(0).Item("Correspond")) & "   ne permet pas de transférer l'article : " & LigneCodeArticle & " de Quantité : " & CDbl(Trim(LigneQuantite)) & " ,  N° :" & Trim(EntetePieceInterne))
                                                        Else
                                                            If ListeStock.Count <> 0 Then
                                                                Dim ExisteArtStock As Boolean = False
                                                                For i As Integer = 0 To ListeStock.Count - 1
                                                                    Dim StockListe() As String = Split(ListeStock.Item(i), ControlChars.Tab)
                                                                    If Trim(StockListe(0)) = Trim(DepotTab.Rows(0).Item("Correspond")) And Trim(StockListe(1)) = Trim(fournisseurTab.Rows(0).Item("Correspond")) Then
                                                                        ExisteArtStock = True
                                                                        If StockListe(2) < CDbl(Trim(LigneQuantite)) Then
                                                                            ExisteLecture = False
                                                                            ErreurJrn.WriteLine("Le Stock Calculé " & StockListe(2) & " dans le dépôt d'origine : " & StockListe(0) & "   ne permet pas de transférer l'article : " & Trim(fournisseurTab.Rows(0).Item("Correspond")) & " de Quantité : " & CDbl(Trim(LigneQuantite)) & " ,  N° :" & Trim(EntetePieceInterne))
                                                                        Else
                                                                            ListeStock.RemoveAt(i)
                                                                            ListeStock.Add(Trim(DepotTab.Rows(0).Item("Correspond")) & ControlChars.Tab & Trim(fournisseurTab.Rows(0).Item("Correspond")) & ControlChars.Tab & (StockListe(2) - CDbl(Trim(LigneQuantite))))
                                                                        End If
                                                                    End If
                                                                Next i
                                                                If ExisteArtStock = False Then
                                                                    ListeStock.Add(Trim(DepotTab.Rows(0).Item("Correspond")) & ControlChars.Tab & Trim(fournisseurTab.Rows(0).Item("Correspond")) & ControlChars.Tab & (DossierTab.Rows(0).Item("AS_QteSto") - CDbl(Trim(LigneQuantite))))
                                                                End If
                                                            Else
                                                                ListeStock.Add(Trim(DepotTab.Rows(0).Item("Correspond")) & ControlChars.Tab & Trim(fournisseurTab.Rows(0).Item("Correspond")) & ControlChars.Tab & (DossierTab.Rows(0).Item("AS_QteSto") - CDbl(Trim(LigneQuantite))))
                                                            End If
                                                        End If
                                                    End If
                                                Else
                                                    ExisteLecture = False
                                                    ErreurJrn.WriteLine("Le Stock dans le dépôt d'origine : " & Trim(fournisseurTab.Rows(0).Item("Correspond")) & "  est NULL et ne permet pas de transférer l'article : " & LigneCodeArticle & " ,  N° :" & Trim(EntetePieceInterne))
                                                End If
                                            Else
                                                ExisteLecture = False
                                                ErreurJrn.WriteLine("Le Stock dans le dépôt d'origine : " & Trim(fournisseurTab.Rows(0).Item("Correspond")) & "   ne permet pas de transférer l'article : " & LigneCodeArticle & " ,  N° :" & Trim(EntetePieceInterne))
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        If Datagridaffiche.Columns.Contains("EnteteIntituleDepotOrigine") = True Then
            If BaseCial.FactoryDepot.ExistIntitule(Trim(EnteteIntituleDepotOrigine)) = True Then
            Else
                fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Depot Stockage>' and Valeurlue ='" & Join(Split(Trim(EnteteIntituleDepotOrigine), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                fournisseurDs = New DataSet
                fournisseurAdap.Fill(fournisseurDs)
                fournisseurTab = fournisseurDs.Tables(0)
                If fournisseurTab.Rows.Count > 0 Then
                    If BaseCial.FactoryDepot.ExistIntitule(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                    Else
                        ExisteLecture = False
                        ErreurJrn.WriteLine("< L'Intitulé dépôt Origine : " & Trim(fournisseurTab.Rows(0).Item("Correspond")) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'existe pas dans Sage >")
                    End If
                Else
                    ExisteLecture = False
                    ErreurJrn.WriteLine("< L'Intitulé dépôt Origine : " & Trim(EnteteIntituleDepotOrigine) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'existe pas dans la table de paramétrage et dans Sage >")
                End If
            End If
        End If
        If Datagridaffiche.Columns.Contains("EnteteIntituleDepotDestination") = True Then
            If BaseCial.FactoryDepot.ExistIntitule(Trim(EnteteIntituleDepotDestination)) = True Then
            Else
                fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Depot Stockage>' and Valeurlue ='" & Join(Split(Trim(EnteteIntituleDepotDestination), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                fournisseurDs = New DataSet
                fournisseurAdap.Fill(fournisseurDs)
                fournisseurTab = fournisseurDs.Tables(0)
                If fournisseurTab.Rows.Count > 0 Then
                    If BaseCial.FactoryDepot.ExistIntitule(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                    Else
                        ExisteLecture = False
                        ErreurJrn.WriteLine("< L'Intitulé dépôt Destination : " & Trim(fournisseurTab.Rows(0).Item("Correspond")) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'existe pas dans Sage >")
                    End If
                Else
                    ExisteLecture = False
                    ErreurJrn.WriteLine("< L'Intitulé dépôt Destination : " & Trim(EnteteIntituleDepotDestination) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'existe pas dans la table de paramétrage et dans Sage >")
                End If
            End If
        End If
        If Datagridaffiche.Columns.Contains("IDDepotEnteteOrigine") = True Then
            If IsNumeric(Trim(IDDepotEnteteOrigine)) = True Then
                statistAdap = New OleDbDataAdapter("select * from F_DEPOT where DE_No ='" & CInt(Trim(IDDepotEnteteOrigine)) & "'", OleSocieteConnect)
                statistDs = New DataSet
                statistAdap.Fill(statistDs)
                statistTab = statistDs.Tables(0)
                If statistTab.Rows.Count <> 0 Then

                Else
                    fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Depot Stockage>' and Valeurlue ='" & Join(Split(Trim(IDDepotEnteteOrigine), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                    fournisseurDs = New DataSet
                    fournisseurAdap.Fill(fournisseurDs)
                    fournisseurTab = fournisseurDs.Tables(0)
                    If fournisseurTab.Rows.Count > 0 Then
                        If IsNumeric(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                            statistAdap = New OleDbDataAdapter("select * from F_DEPOT where DE_No ='" & CInt(Trim(fournisseurTab.Rows(0).Item("Correspond"))) & "'", OleSocieteConnect)
                            statistDs = New DataSet
                            statistAdap.Fill(statistDs)
                            statistTab = statistDs.Tables(0)
                            If statistTab.Rows.Count <> 0 Then

                            Else
                                ExisteLecture = False
                                ErreurJrn.WriteLine("< L'ID dépôt Origine : " & fournisseurTab.Rows(0).Item("Correspond") & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'existe pas dans Sage >")
                            End If
                        Else
                            ExisteLecture = False
                            ErreurJrn.WriteLine("< L'ID dépôt Origine : " & fournisseurTab.Rows(0).Item("Correspond") & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " existant dans la table de paramétrage n'est pas numérique>")
                        End If
                    Else
                        ExisteLecture = False
                        ErreurJrn.WriteLine("< L'ID dépôt Origine : " & Trim(IDDepotEnteteOrigine) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'existe pas dans la table de paramétrage et dans Sage >")
                    End If
                End If
            Else
                ExisteLecture = False
                ErreurJrn.WriteLine("< L'ID dépôt Origine : " & IDDepotEnteteOrigine & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " existant dans le fichier n'est pas numérique>")
            End If
        End If
        If Datagridaffiche.Columns.Contains("IDDepotEnteteDestination") = True Then
            If IsNumeric(Trim(IDDepotEnteteDestination)) = True Then
                statistAdap = New OleDbDataAdapter("select * from F_DEPOT where DE_No ='" & CInt(Trim(IDDepotEnteteDestination)) & "'", OleSocieteConnect)
                statistDs = New DataSet
                statistAdap.Fill(statistDs)
                statistTab = statistDs.Tables(0)
                If statistTab.Rows.Count <> 0 Then

                Else
                    fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Depot Stockage>' and Valeurlue ='" & Join(Split(Trim(IDDepotEnteteDestination), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                    fournisseurDs = New DataSet
                    fournisseurAdap.Fill(fournisseurDs)
                    fournisseurTab = fournisseurDs.Tables(0)
                    If fournisseurTab.Rows.Count > 0 Then
                        If IsNumeric(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                            statistAdap = New OleDbDataAdapter("select * from F_DEPOT where DE_No ='" & CInt(Trim(fournisseurTab.Rows(0).Item("Correspond"))) & "'", OleSocieteConnect)
                            statistDs = New DataSet
                            statistAdap.Fill(statistDs)
                            statistTab = statistDs.Tables(0)
                            If statistTab.Rows.Count <> 0 Then

                            Else
                                ExisteLecture = False
                                ErreurJrn.WriteLine("< L'ID dépôt Destination : " & fournisseurTab.Rows(0).Item("Correspond") & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'existe pas dans Sage >")
                            End If
                        Else
                            ExisteLecture = False
                            ErreurJrn.WriteLine("< L'ID dépôt Destination : " & fournisseurTab.Rows(0).Item("Correspond") & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " existant dans la table de paramétrage n'est pas numérique>")
                        End If
                    Else
                        ExisteLecture = False
                        ErreurJrn.WriteLine("< L'ID dépôt Destination : " & Trim(IDDepotEnteteDestination) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'existe pas dans la table de paramétrage et dans Sage >")
                    End If
                End If
            Else
                ExisteLecture = False
                ErreurJrn.WriteLine("< L'ID dépôt Destination : " & IDDepotEnteteDestination & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " existant dans le fichier n'est pas numérique>")
            End If
        End If
        If TypeImport = "Import" Then
            If Datagridaffiche.Columns.Contains("IDDepotLigne") = True Then
                If IsNumeric(Trim(IDDepotLigne)) = True Then
                    statistAdap = New OleDbDataAdapter("select * from F_DEPOT where DE_No ='" & CInt(Trim(IDDepotLigne)) & "'", OleSocieteConnect)
                    statistDs = New DataSet
                    statistAdap.Fill(statistDs)
                    statistTab = statistDs.Tables(0)
                    If statistTab.Rows.Count <> 0 Then

                    Else
                        fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Depot Stockage>' and Valeurlue ='" & Join(Split(Trim(IDDepotLigne), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                        fournisseurDs = New DataSet
                        fournisseurAdap.Fill(fournisseurDs)
                        fournisseurTab = fournisseurDs.Tables(0)
                        If fournisseurTab.Rows.Count > 0 Then
                            If IsNumeric(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                                statistAdap = New OleDbDataAdapter("select * from F_DEPOT where DE_No ='" & CInt(Trim(fournisseurTab.Rows(0).Item("Correspond"))) & "'", OleSocieteConnect)
                                statistDs = New DataSet
                                statistAdap.Fill(statistDs)
                                statistTab = statistDs.Tables(0)
                                If statistTab.Rows.Count <> 0 Then

                                Else
                                    ExisteLecture = False
                                    ErreurJrn.WriteLine("< L'ID dépôt Ligne : " & fournisseurTab.Rows(0).Item("Correspond") & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'existe pas dans Sage >")
                                End If
                            Else
                                ExisteLecture = False
                                ErreurJrn.WriteLine("< L'ID dépôt Ligne : " & fournisseurTab.Rows(0).Item("Correspond") & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " existant dans la table de paramétrage n'est pas numérique>")
                            End If
                        Else
                            ExisteLecture = False
                            ErreurJrn.WriteLine("< L'ID dépôt Ligne : " & Trim(IDDepotLigne) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'existe pas dans la table de paramétrage et dans Sage >")
                        End If
                    End If
                Else
                    ExisteLecture = False
                    ErreurJrn.WriteLine("< L'ID dépôt Ligne : " & IDDepotLigne & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " existant dans le fichier n'est pas numérique>")
                End If
            Else
                If Datagridaffiche.Columns.Contains("LigneIntituleDepot") = True Then
                    If BaseCial.FactoryDepot.ExistIntitule(Trim(LigneIntituleDepot)) = True Then
                    Else
                        fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Depot Stockage>' and Valeurlue ='" & Join(Split(Trim(LigneIntituleDepot), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                        fournisseurDs = New DataSet
                        fournisseurAdap.Fill(fournisseurDs)
                        fournisseurTab = fournisseurDs.Tables(0)
                        If fournisseurTab.Rows.Count > 0 Then
                            If BaseCial.FactoryDepot.ExistIntitule(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                            Else
                                ExisteLecture = False
                                ErreurJrn.WriteLine("< L'Intitulé dépôt Ligne : " & Trim(fournisseurTab.Rows(0).Item("Correspond")) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'existe pas dans Sage >")
                            End If
                        Else
                            ExisteLecture = False
                            ErreurJrn.WriteLine("< L'Intitulé dépôt Ligne : " & Trim(LigneIntituleDepot) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'existe pas dans la table de paramétrage et dans Sage >")
                        End If
                    End If
                Else
                    ExisteLecture = False
                    ErreurJrn.WriteLine("Un dépôt en Ligne est Obligatoire dans le descriptif d'import! Il s'agit d'un import de document avec Ligne d'Origine et Destination")
                End If
            End If
        End If
        If Datagridaffiche.Columns.Contains("EnteteDateDocument") = True Then
            If Trim(EnteteDateDocument) <> "" Then
                If Verificatdate(Trim(EnteteDateDocument), FormatDatefichier, "Date de Document") = True Then
                Else
                    ExisteLecture = False
                End If
            End If
        End If
        If Datagridaffiche.Columns.Contains("EnteteSoucheDocument") = True Then
            If Trim(EnteteSoucheDocument) <> "" Then
                If EstNumeric(Trim(EnteteSoucheDocument), DecimalNomb, DecimalMone) = True Then
                    statistAdap = New OleDbDataAdapter("select * from P_SOUCHEINTERNE where cbIndice ='" & CInt(CInt(RenvoiTaux(Trim(EnteteSoucheDocument), DecimalNomb, DecimalMone)) + 1) & "'", OleSocieteConnect)
                    statistDs = New DataSet
                    statistAdap.Fill(statistDs)
                    statistTab = statistDs.Tables(0)
                    If statistTab.Rows.Count <> 0 Then

                    Else
                        ExisteLecture = False
                        ErreurJrn.WriteLine("< L'Indice de la  Souche du Document : " & CInt(CInt(RenvoiTaux(Trim(EnteteSoucheDocument), DecimalNomb, DecimalMone)) + 1) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'existe pas dans Sage >")
                    End If
                Else
                    If BaseCial.FactorySoucheStock.ExistIntitule(Trim(EnteteSoucheDocument)) = True Then


                    Else
                        fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Intitule Souche>' and Valeurlue ='" & Join(Split(Trim(EnteteSoucheDocument), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                        fournisseurDs = New DataSet
                        fournisseurAdap.Fill(fournisseurDs)
                        fournisseurTab = fournisseurDs.Tables(0)
                        If fournisseurTab.Rows.Count <> 0 Then
                            If BaseCial.FactorySoucheStock.ExistIntitule(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then

                            Else
                                ExisteLecture = False
                                ErreurJrn.WriteLine("< L'Intitulé de la  Souche du Document : " & Trim(fournisseurTab.Rows(0).Item("Correspond")) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'existe pas dans Sage >")
                            End If
                        Else
                            ExisteLecture = False
                            ErreurJrn.WriteLine("< L'Intitulé de la  Souche du Document : " & Trim(EnteteSoucheDocument) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'existe pas dans la table de paramétrage et dans Sage >")
                        End If
                    End If
                End If
            End If
        End If
        If Datagridaffiche.Columns.Contains("LignePoidsNet") = True Then
            If Trim(LignePoidsNet) <> "" Then
                If EstNumeric(Trim(LignePoidsNet), DecimalNomb, DecimalMone) = True Then
                Else
                    ExisteLecture = False
                    ErreurJrn.WriteLine("< Le poids Net : " & Trim(LignePoidsNet) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'est pas numérique >")
                End If
            End If
        End If
        If Datagridaffiche.Columns.Contains("LignePoidsBrut") = True Then
            If Trim(LignePoidsBrut) <> "" Then
                If EstNumeric(Trim(LignePoidsBrut), DecimalNomb, DecimalMone) = True Then
                Else
                    ExisteLecture = False
                    ErreurJrn.WriteLine("< Le poids brut : " & Trim(LignePoidsBrut) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'est pas numérique >")
                End If
            End If
        End If
        If Datagridaffiche.Columns.Contains("LignePrixUnitaire") = True Then
            If Trim(LignePrixUnitaire) <> "" Then
                If EstNumeric(Trim(LignePrixUnitaire), DecimalNomb, DecimalMone) = True Then
                Else
                    ExisteLecture = False
                    ErreurJrn.WriteLine("< Le prix unitaire : " & Trim(LignePrixUnitaire) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'est pas numérique >")
                End If
            End If
        End If
        If Datagridaffiche.Columns.Contains("LigneDatedeFabrication") = True Then
            If Trim(LigneDatedeFabrication) <> "" Then
                If IsDate(Trim(LigneDatedeFabrication)) = True Then
                Else
                    ExisteLecture = False
                    ErreurJrn.WriteLine("< La date de Fabrication : " & Trim(LigneDatedeFabrication) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'est pas au format date >")
                End If
            End If
        End If
        If Datagridaffiche.Columns.Contains("LigneDatedePeremption") = True Then
            If Trim(LigneDatedePeremption) <> "" Then
                If IsDate(Trim(LigneDatedePeremption)) = True Then
                Else
                    ExisteLecture = False
                    ErreurJrn.WriteLine("< La date de Peremption : " & Trim(LigneDatedePeremption) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'est pas au format date >")
                End If
            End If
        End If
        If Datagridaffiche.Columns.Contains("LigneQuantite") = True Then
            If Trim(LigneQuantite) <> "" Then
                If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then

                Else
                    ExisteLecture = False
                    ErreurJrn.WriteLine("< La Quantité : " & Trim(LigneQuantite) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'est pas numérique >")
                End If
            End If
        End If

        If Datagridaffiche.Columns.Contains("LigneCodeArticle") = True Then
            If Trim(LigneCodeArticle) <> "" Then
                If BaseCial.FactoryArticle.ExistReference(Trim(LigneCodeArticle)) = True Then
                    If BaseCial.FactoryArticle.ReadReference(Trim(LigneCodeArticle)).AR_SuiviStock = SuiviStockType.SuiviStockTypeSerie Then
                        ExisteLecture = False
                        ErreurJrn.WriteLine("< La Référence Article : " & Trim(LigneCodeArticle) & " existant en Gestion Commerciale - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " est suivi en Série >")
                    End If
                    If BaseCial.FactoryArticle.ReadReference(Trim(LigneCodeArticle)).AR_SuiviStock = SuiviStockType.SuiviStockTypeLot Then
                        ExisteLecture = False
                        ErreurJrn.WriteLine("< La Référence Article : " & Trim(LigneCodeArticle) & " existant en Gestion Commerciale - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " est suivi en Lot >")
                    End If
                    If BaseCial.FactoryArticle.ReadReference(Trim(LigneCodeArticle)).AR_Type = ArticleType.ArticleTypeGamme Then
                        ExisteLecture = False
                        ErreurJrn.WriteLine("< La Référence Article : " & Trim(LigneCodeArticle) & " existant en Gestion Commerciale - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " est de type Gamme >")
                    End If
                    If Trim(LigneQuantite) <> "" Then

                    Else
                        ExisteLecture = False
                        ErreurJrn.WriteLine("< La quantité pour La Référence Article : " & Trim(LigneCodeArticle) & " existant en Gestion Commerciale - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " doit être obligatoire >")
                    End If
                Else
                    fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<ArticleStock>' and Valeurlue ='" & Join(Split(Trim(LigneCodeArticle), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                    fournisseurDs = New DataSet
                    fournisseurAdap.Fill(fournisseurDs)
                    fournisseurTab = fournisseurDs.Tables(0)
                    If fournisseurTab.Rows.Count <> 0 Then
                        If BaseCial.FactoryArticle.ExistReference(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                            If BaseCial.FactoryArticle.ReadReference(Trim(fournisseurTab.Rows(0).Item("Correspond"))).AR_SuiviStock = SuiviStockType.SuiviStockTypeSerie Then
                                ExisteLecture = False
                                ErreurJrn.WriteLine("< La Référence Article : " & Trim(fournisseurTab.Rows(0).Item("Correspond")) & " existant en Gestion Commerciale - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " est suivi en Série >")
                            End If
                            If BaseCial.FactoryArticle.ReadReference(Trim(fournisseurTab.Rows(0).Item("Correspond"))).AR_SuiviStock = SuiviStockType.SuiviStockTypeLot Then
                                ExisteLecture = False
                                ErreurJrn.WriteLine("< La Référence Article : " & Trim(fournisseurTab.Rows(0).Item("Correspond")) & " existant en Gestion Commerciale - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " est suivi en Lot >")
                            End If
                            If BaseCial.FactoryArticle.ReadReference(Trim(fournisseurTab.Rows(0).Item("Correspond"))).AR_Type = ArticleType.ArticleTypeGamme Then
                                ExisteLecture = False
                                ErreurJrn.WriteLine("< La Référence Article : " & Trim(fournisseurTab.Rows(0).Item("Correspond")) & " existant en Gestion Commerciale - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " est de type Gamme >")
                            End If
                            If Trim(LigneQuantite) <> "" Then

                            Else
                                ExisteLecture = False
                                ErreurJrn.WriteLine("< La quantité pour La Référence Article : " & Trim(fournisseurTab.Rows(0).Item("Correspond")) & " existant en Gestion Commerciale - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " doit être obligatoire >")
                            End If
                        Else
                            ExisteLecture = False
                            ErreurJrn.WriteLine("< La Référence Article : " & Trim(fournisseurTab.Rows(0).Item("Correspond")) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'existe pas dans Sage >")
                        End If
                    Else
                        ExisteLecture = False
                        ErreurJrn.WriteLine("< La Référence Article : " & Trim(LigneCodeArticle) & " - Dépôt : " & Join(Split(Trim(EnteteIntituleDepotOrigine), "'"), "''") & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'existe pas dans la table de paramétrage et dans Sage >")
                    End If
                End If
            Else
                If Datagridaffiche.Columns.Contains("LigneQuantite") = True Then
                    If Trim(LigneQuantite) <> "" Then
                        ExisteLecture = False
                        ErreurJrn.WriteLine("< La quantité :" & Trim(LigneQuantite) & " ne doit pas être renseignée pour La Référence Article : " & Trim(LigneCodeArticle) & "vide - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " >")
                    End If
                End If
                If Datagridaffiche.Columns.Contains("LignePrixUnitaire") = True Then
                    If Trim(LignePrixUnitaire) <> "" Then
                        ExisteLecture = False
                        ErreurJrn.WriteLine("< Le prix unitaire :" & Trim(LignePrixUnitaire) & " ne doit pas être renseignée pour La Référence Article : " & Trim(LigneCodeArticle) & "vide - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " >")
                    End If
                End If
            End If
        Else
            If Datagridaffiche.Columns.Contains("LigneQuantite") = True Then
                If Trim(LigneQuantite) <> "" Then
                    ExisteLecture = False
                    ErreurJrn.WriteLine("< La quantité :" & Trim(LigneQuantite) & " ne doit pas être renseignée pour La Référence Article : " & Trim(LigneCodeArticle) & "vide - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " >")
                End If
            End If
            If Datagridaffiche.Columns.Contains("LignePrixUnitaire") = True Then
                If Trim(LignePrixUnitaire) <> "" Then
                    ExisteLecture = False
                    ErreurJrn.WriteLine("< Le prix unitaire :" & Trim(LignePrixUnitaire) & " ne doit pas être renseignée pour La Référence Article : " & Trim(LigneCodeArticle) & "vide - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " >")
                End If
            End If
        End If

        'Traitement des Infos Libres
        Try
            If infoLigne.Count > 0 Then
                While infoLigne.Count <> 0
                    OleAdaptaterDelete = New OleDbDataAdapter("select * From WIT_COL where Libelle='" & Trim(Datagridaffiche.Columns(infoLigne.Item(0)).Name) & "' And InfoLigne=True", OleConnenection)
                    OleDeleteDataset = New DataSet
                    OleAdaptaterDelete.Fill(OleDeleteDataset)
                    OledatableDelete = OleDeleteDataset.Tables(0)
                    If OledatableDelete.Rows.Count <> 0 Then
                        'L'info Libre Parametrée par l'utilisateur existe dans Sage
                        If IsNothing(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) = False Then
                            If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                statistAdap = New OleDbDataAdapter("select * from cbSysLibre where CB_File='F_DOCLIGNE' and CB_Name ='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "'", OleSocieteConnect)
                                statistDs = New DataSet
                                statistAdap.Fill(statistDs)
                                statistTab = statistDs.Tables(0)
                                If statistTab.Rows.Count <> 0 Then
                                    'Texte
                                    If statistTab.Rows(0).Item("CB_Type") = 9 Then
                                        OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                        OleRecherDataset = New DataSet
                                        OleRecherAdapter.Fill(OleRecherDataset)
                                        OleRechDatable = OleRecherDataset.Tables(0)
                                        If OleRechDatable.Rows.Count <> 0 Then
                                            If statistTab.Rows(0).Item("CB_Len") >= Strings.Len(Trim(OleRechDatable.Rows(0).Item("Correspond"))) Then

                                            Else
                                                ExisteLecture = False
                                                ErreurJrn.WriteLine("La Longueur de l'info libre : " & Datagridaffiche.Columns(infoLigne.Item(0)).Name & " est inférieur au valeur adressée: " & Trim(OleRechDatable.Rows(0).Item("Correspond")) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
                                            End If
                                        Else
                                            If statistTab.Rows(0).Item("CB_Len") >= Strings.Len(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value)) Then

                                            Else
                                                ExisteLecture = False
                                                ErreurJrn.WriteLine("La Longueur de l'info libre : " & Datagridaffiche.Columns(infoLigne.Item(0)).Name & " est inférieur au valeur adressée: " & Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
                                            End If
                                        End If
                                    End If
                                    'Table
                                    If statistTab.Rows(0).Item("CB_Type") = 22 Then
                                        OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                        OleRecherDataset = New DataSet
                                        OleRecherAdapter.Fill(OleRecherDataset)
                                        OleRechDatable = OleRecherDataset.Tables(0)
                                        If OleRechDatable.Rows.Count <> 0 Then
                                            If statistTab.Rows(0).Item("CB_Len") > Strings.Len(Trim(OleRechDatable.Rows(0).Item("Correspond"))) Then

                                            Else
                                                ExisteLecture = False
                                                ErreurJrn.WriteLine("La Longueur (21) de l'info libre de type table : " & Datagridaffiche.Columns(infoLigne.Item(0)).Name & " est inférieur au valeur adressée: " & Trim(OleRechDatable.Rows(0).Item("Correspond")) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
                                            End If
                                        Else
                                            If statistTab.Rows(0).Item("CB_Len") > Strings.Len(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value)) Then

                                            Else
                                                ExisteLecture = False
                                                ErreurJrn.WriteLine("La Longueur (21) de l'info libre de type table : " & Datagridaffiche.Columns(infoLigne.Item(0)).Name & " est inférieur au valeur adressée: " & Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
                                            End If
                                        End If
                                    End If
                                    'Montant
                                    If statistTab.Rows(0).Item("CB_Type") = 20 Then
                                        If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                            OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                            OleRecherDataset = New DataSet
                                            OleRecherAdapter.Fill(OleRecherDataset)
                                            OleRechDatable = OleRecherDataset.Tables(0)
                                            If OleRechDatable.Rows.Count <> 0 Then
                                                If EstNumeric(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone) = True Then

                                                Else
                                                    ExisteLecture = False
                                                    ErreurJrn.WriteLine("La valeur entrée : " & Trim(OleRechDatable.Rows(0).Item("Correspond")) & "doit être de type numérique pour l'info libre : " & Datagridaffiche.Columns(infoLigne.Item(0)).Name & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
                                                End If
                                            Else
                                                If EstNumeric(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), DecimalNomb, DecimalMone) = True Then

                                                Else
                                                    ExisteLecture = False
                                                    ErreurJrn.WriteLine("La valeur entrée : " & Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) & "doit être de type numérique pour l'info libre : " & Datagridaffiche.Columns(infoLigne.Item(0)).Name & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
                                                End If
                                            End If
                                        End If
                                    End If
                                    'Valeur
                                    If statistTab.Rows(0).Item("CB_Type") = 7 Then
                                        If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                            OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                            OleRecherDataset = New DataSet
                                            OleRecherAdapter.Fill(OleRecherDataset)
                                            OleRechDatable = OleRecherDataset.Tables(0)
                                            If OleRechDatable.Rows.Count <> 0 Then
                                                If EstNumeric(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone) = True Then

                                                Else
                                                    ExisteLecture = False
                                                    ErreurJrn.WriteLine("La valeur entrée : " & Trim(OleRechDatable.Rows(0).Item("Correspond")) & "doit être de type numérique pour l'info libre : " & Datagridaffiche.Columns(infoLigne.Item(0)).Name & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
                                                End If
                                            Else
                                                If EstNumeric(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), DecimalNomb, DecimalMone) = True Then

                                                Else
                                                    ExisteLecture = False
                                                    ErreurJrn.WriteLine("La valeur entrée : " & Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) & "doit être de type numérique pour l'info libre : " & Datagridaffiche.Columns(infoLigne.Item(0)).Name & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
                                                End If
                                            End If
                                        End If
                                    End If

                                    'Date Court
                                    If statistTab.Rows(0).Item("CB_Type") = 3 Then
                                        If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                            OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                            OleRecherDataset = New DataSet
                                            OleRecherAdapter.Fill(OleRecherDataset)
                                            OleRechDatable = OleRecherDataset.Tables(0)
                                            If OleRechDatable.Rows.Count <> 0 Then
                                                If Verificatdate(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then

                                                Else
                                                    ExisteLecture = False
                                                End If
                                            Else
                                                If Verificatdate(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then
                                                Else
                                                    ExisteLecture = False
                                                End If
                                            End If
                                        End If
                                    End If
                                    'Date Longue
                                    If statistTab.Rows(0).Item("CB_Type") = 14 Then
                                        If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                            OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                            OleRecherDataset = New DataSet
                                            OleRecherAdapter.Fill(OleRecherDataset)
                                            OleRechDatable = OleRecherDataset.Tables(0)
                                            If OleRechDatable.Rows.Count <> 0 Then
                                                If Verificatdate(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then

                                                Else
                                                    ExisteLecture = False
                                                End If
                                            Else
                                                If Verificatdate(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then
                                                Else
                                                    ExisteLecture = False
                                                End If
                                            End If
                                        End If
                                    End If
                                Else
                                    ExisteLecture = False
                                    ErreurJrn.WriteLine(" Impossible de traiter l'information libre en Ligne de Document :" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "  Il est inexistant dans la table CbsysLibre de Sage")
                                End If
                            End If
                        Else
                            'nothing
                        End If
                    Else
                        ExisteLecture = False
                        ErreurJrn.WriteLine(" Impossible de traiter l'information libre en Ligne de Document :" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "  Il est inexistant dans la table de Paramétrage")
                    End If
                    'L'info Libre Parametrée par l'utilisateur n'existe pas dans Sage
                    infoLigne.RemoveAt(0)
                End While
            End If
        Catch ex As Exception
            exceptionTrouve = True
            ExisteLecture = False
            ErreurJrn.WriteLine(" Erreur de Création de L'information Libre Ligne Document " & ex.Message & ", vérifiez la longueur de la chaine avec celle paramétrée/ ou la cohérence des données")
        End Try
        'Traitement des Infos Libres
        Try
            If infoListe.Count > 0 Then
                While infoListe.Count <> 0
                    OleAdaptaterDelete = New OleDbDataAdapter("select * From WIT_COL where Libelle='" & Trim(Datagridaffiche.Columns(infoListe.Item(0)).Name) & "' And Libre=True", OleConnenection)
                    OleDeleteDataset = New DataSet
                    OleAdaptaterDelete.Fill(OleDeleteDataset)
                    OledatableDelete = OleDeleteDataset.Tables(0)
                    If OledatableDelete.Rows.Count <> 0 Then
                        'L'info Libre Parametrée par l'utilisateur existe dans Sage
                        If IsNothing(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) = False Then
                            If Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) <> "" Then
                                statistAdap = New OleDbDataAdapter("select * from cbSysLibre where CB_File='F_DOCENTETE' and CB_Name ='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "'", OleSocieteConnect)
                                statistDs = New DataSet
                                statistAdap.Fill(statistDs)
                                statistTab = statistDs.Tables(0)
                                If statistTab.Rows.Count <> 0 Then
                                    'Texte
                                    If statistTab.Rows(0).Item("CB_Type") = 9 Then
                                        OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                        OleRecherDataset = New DataSet
                                        OleRecherAdapter.Fill(OleRecherDataset)
                                        OleRechDatable = OleRecherDataset.Tables(0)
                                        If OleRechDatable.Rows.Count <> 0 Then
                                            If statistTab.Rows(0).Item("CB_Len") >= Strings.Len(Trim(OleRechDatable.Rows(0).Item("Correspond"))) Then

                                            Else
                                                ExisteLecture = False
                                                ErreurJrn.WriteLine("La Longueur de l'info libre : " & Datagridaffiche.Columns(infoListe.Item(0)).Name & " est inférieur au valeur adressée: " & Trim(OleRechDatable.Rows(0).Item("Correspond")) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
                                            End If
                                        Else
                                            If statistTab.Rows(0).Item("CB_Len") >= Strings.Len(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value)) Then

                                            Else
                                                ExisteLecture = False
                                                ErreurJrn.WriteLine("La Longueur de l'info libre : " & Datagridaffiche.Columns(infoListe.Item(0)).Name & " est inférieur au valeur adressée: " & Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
                                            End If
                                        End If
                                    End If
                                    'Table
                                    If statistTab.Rows(0).Item("CB_Type") = 22 Then
                                        OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                        OleRecherDataset = New DataSet
                                        OleRecherAdapter.Fill(OleRecherDataset)
                                        OleRechDatable = OleRecherDataset.Tables(0)
                                        If OleRechDatable.Rows.Count <> 0 Then
                                            If statistTab.Rows(0).Item("CB_Len") > Strings.Len(Trim(OleRechDatable.Rows(0).Item("Correspond"))) Then

                                            Else
                                                ExisteLecture = False
                                                ErreurJrn.WriteLine("La Longueur (21) de l'info libre de type table : " & Datagridaffiche.Columns(infoListe.Item(0)).Name & " est inférieur au valeur adressée: " & Trim(OleRechDatable.Rows(0).Item("Correspond")) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
                                            End If
                                        Else
                                            If statistTab.Rows(0).Item("CB_Len") > Strings.Len(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value)) Then

                                            Else
                                                ExisteLecture = False
                                                ErreurJrn.WriteLine("La Longueur (21) de l'info libre de type table : " & Datagridaffiche.Columns(infoListe.Item(0)).Name & " est inférieur au valeur adressée: " & Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
                                            End If
                                        End If
                                    End If
                                    'Montant
                                    If statistTab.Rows(0).Item("CB_Type") = 20 Then
                                        If Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) <> "" Then
                                            OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                            OleRecherDataset = New DataSet
                                            OleRecherAdapter.Fill(OleRecherDataset)
                                            OleRechDatable = OleRecherDataset.Tables(0)
                                            If OleRechDatable.Rows.Count <> 0 Then
                                                If EstNumeric(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone) = True Then

                                                Else
                                                    ExisteLecture = False
                                                    ErreurJrn.WriteLine("La valeur entrée : " & Trim(OleRechDatable.Rows(0).Item("Correspond")) & "doit être de type numérique pour l'info libre : " & Datagridaffiche.Columns(infoLigne.Item(0)).Name & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
                                                End If
                                            Else
                                                If EstNumeric(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), DecimalNomb, DecimalMone) = True Then

                                                Else
                                                    ExisteLecture = False
                                                    ErreurJrn.WriteLine("La valeur entrée : " & Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) & "doit être de type numérique pour l'info libre : " & Datagridaffiche.Columns(infoLigne.Item(0)).Name & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
                                                End If
                                            End If
                                        End If
                                    End If
                                    'Valeur
                                    If statistTab.Rows(0).Item("CB_Type") = 7 Then
                                        If Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) <> "" Then
                                            OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                            OleRecherDataset = New DataSet
                                            OleRecherAdapter.Fill(OleRecherDataset)
                                            OleRechDatable = OleRecherDataset.Tables(0)
                                            If OleRechDatable.Rows.Count <> 0 Then
                                                If EstNumeric(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone) = True Then

                                                Else
                                                    ExisteLecture = False
                                                    ErreurJrn.WriteLine("La valeur entrée : " & Trim(OleRechDatable.Rows(0).Item("Correspond")) & "doit être de type numérique pour l'info libre : " & Datagridaffiche.Columns(infoLigne.Item(0)).Name & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
                                                End If
                                            Else
                                                If EstNumeric(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), DecimalNomb, DecimalMone) = True Then

                                                Else
                                                    ExisteLecture = False
                                                    ErreurJrn.WriteLine("La valeur entrée : " & Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) & "doit être de type numérique pour l'info libre : " & Datagridaffiche.Columns(infoLigne.Item(0)).Name & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
                                                End If
                                            End If
                                        End If
                                    End If

                                    'Date Court
                                    If statistTab.Rows(0).Item("CB_Type") = 3 Then
                                        If Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) <> "" Then
                                            OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                            OleRecherDataset = New DataSet
                                            OleRecherAdapter.Fill(OleRecherDataset)
                                            OleRechDatable = OleRecherDataset.Tables(0)
                                            If OleRechDatable.Rows.Count <> 0 Then
                                                If Verificatdate(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier, Datagridaffiche.Columns(infoListe.Item(0)).Name) = True Then

                                                Else
                                                    ExisteLecture = False
                                                End If
                                            Else
                                                If Verificatdate(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), FormatDatefichier, Datagridaffiche.Columns(infoListe.Item(0)).Name) = True Then
                                                Else
                                                    ExisteLecture = False
                                                End If
                                            End If
                                        End If
                                    End If
                                    'Date Longue
                                    If statistTab.Rows(0).Item("CB_Type") = 14 Then
                                        If Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) <> "" Then
                                            OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                            OleRecherDataset = New DataSet
                                            OleRecherAdapter.Fill(OleRecherDataset)
                                            OleRechDatable = OleRecherDataset.Tables(0)
                                            If OleRechDatable.Rows.Count <> 0 Then
                                                If Verificatdate(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier, Datagridaffiche.Columns(infoListe.Item(0)).Name) = True Then

                                                Else
                                                    ExisteLecture = False
                                                End If
                                            Else
                                                If Verificatdate(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), FormatDatefichier, Datagridaffiche.Columns(infoListe.Item(0)).Name) = True Then
                                                Else
                                                    ExisteLecture = False
                                                End If
                                            End If
                                        End If
                                    End If
                                Else
                                    ExisteLecture = False
                                    ErreurJrn.WriteLine(" Impossible de traiter l'information libre en Entête de Document :" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "  Il est inexistant dans la table CbsysLibre de Sage")
                                End If
                            End If
                        Else
                            'nothing
                        End If
                    Else
                        ExisteLecture = False
                        ErreurJrn.WriteLine(" Impossible de traiter l'information libre en Entête de Document :" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "  Il est inexistant dans la table de Paramétrage")
                    End If
                    'L'info Libre Parametrée par l'utilisateur n'existe pas dans Sage
                    infoListe.RemoveAt(0)
                End While
            End If
        Catch ex As Exception
            ExisteLecture = False
            exceptionTrouve = True
            ErreurJrn.WriteLine(" Erreur de Création de L'information Libre Entête de Document " & ex.Message & " , vérifiez la longueur de la chaine avec celle paramétrée/ ou la cohérence des données")
        End Try
    End Sub
    Private Sub Modification_Integration_Fichier(ByVal sPathFilexporter As String, ByVal spathFileFormat As String, ByRef Formatype As String, ByRef Base_Excel As String, ByRef sColumnsSepar As String, ByRef IdentifiantPiece As String, ByRef FormatdeDatefich As String)
        Dim OleAdaptater As OleDbDataAdapter
        Dim OleAfficheDataset As DataSet
        Dim Oledatable As DataTable
        Dim m As Integer
        Dim jColD As Integer
        Dim iLine As Integer
        Dim aRows() As String
        Dim iColPosition, iColGauchetxt As Integer
        Dim i As Integer, aCols() As String
        Initialiser()
        iLine = 0
        aRows = Nothing
        Dim DossierAdap As OleDbDataAdapter
        Dim DossierDs As DataSet
        Dim DossierTab As DataTable
        Dim FormatQte As Integer = 0
        DossierAdap = New OleDbDataAdapter("select * from P_DOSSIERCIAL", OleSocieteConnect)
        DossierDs = New DataSet
        DossierAdap.Fill(DossierDs)
        DossierTab = DossierDs.Tables(0)
        If DossierTab.Rows.Count <> 0 Then
            If InStr(DossierTab.Rows(0).Item("D_FormatQte"), ",") <> 0 Then
                If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ",")), ".") <> 0 Then
                    FormatQte = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ","))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ",")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ","))), ".")))
                Else
                    FormatQte = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ",")))
                End If
            Else
                If InStr(DossierTab.Rows(0).Item("D_FormatQte"), ".") <> 0 Then
                    If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ".")), ",") <> 0 Then
                        FormatQte = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), "."))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ".")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), "."))), ",")))
                    Else
                        FormatQte = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ".")))
                    End If
                End If
            End If
        End If
        If Trim(Formatype) = "Excel" Then
            Try
                If AffichFormatintegration(spathFileFormat, Formatype) = True Then
                    ProgressBar1.Value = ProgressBar1.Minimum
                    Datagridaffiche.Rows.Clear()
                    NbreTotal = DecFormat
                    OleAdaptater = New OleDbDataAdapter("select * from [" & Base_Excel & "$] ", OleExcelConnected)
                    OleAfficheDataset = New DataSet
                    OleAdaptater.Fill(OleAfficheDataset)
                    Oledatable = OleAfficheDataset.Tables(0)
                    If Oledatable.Rows.Count <> 0 Then
                        ProgresMax = Oledatable.Rows.Count - DecFormat
                        m = 0
                        infoListe = New List(Of Integer)
                        infoLigne = New List(Of Integer)
                        For i = DecFormat To Oledatable.Rows.Count - 1
                            Datagridaffiche.RowCount = iLine + 1 - m
                            For jColD = 0 To Datagridaffiche.ColumnCount - 1
                                iColPosition = CInt(Strings.Right(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1), Strings.Len(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1)) - InStr(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1), "(")))
                                If iColPosition <= Oledatable.Columns.Count Then
                                    If iColPosition <> 0 Then
                                        If Convert.IsDBNull(Oledatable.Rows(i).Item(iColPosition - 1)) = False Then
                                            Datagridaffiche.Item(jColD, iLine - m).Value = Trim(Oledatable.Rows(i).Item(iColPosition - 1))
                                        Else
                                            Datagridaffiche.Item(jColD, iLine - m).Value = ""
                                        End If
                                    Else
                                        Datagridaffiche.Item(jColD, iLine - m).Value = LireFichierFormat(spathFileFormat, Datagridaffiche.Columns(jColD).Name, Formatype)
                                    End If
                                Else
                                    Datagridaffiche.Item(jColD, iLine - m).Value = ""
                                End If
                            Next jColD
                            iLine = iLine + 1
                            If i Mod 10 = 0 Then
                                Me.Refresh()
                                Modification_Integrer_Ecriture(IdentifiantPiece, FormatdeDatefich, Trim(Formatype), LirePUDefaut(spathFileFormat, Trim(Formatype)), LireLigneArticle(spathFileFormat, Trim(Formatype)), FormatQte)
                                m = iLine
                            Else
                                If i = (Oledatable.Rows.Count - 1) Then
                                    Me.Refresh()
                                    Modification_Integrer_Ecriture(IdentifiantPiece, FormatdeDatefich, Trim(Formatype), LirePUDefaut(spathFileFormat, Trim(Formatype)), LireLigneArticle(spathFileFormat, Trim(Formatype)), FormatQte)
                                    m = iLine
                                End If
                            End If
                        Next i
                    End If
                Else
                    ErreurJrn.WriteLine("Impossible d' Integrer le Fichier " & sPathFilexporter & " car Erreur de lecture du Fichier Format " & spathFileFormat)
                End If
            Catch ex As Exception
                exceptionTrouve = True
            End Try
        Else
            If Trim(Formatype) = "Délimité" Or Trim(Formatype) = "Tabulation" Or Trim(Formatype) = "Pipe" Then
                Try
                    If AffichFormatintegration(spathFileFormat, Formatype) = True Then
                        aRows = GetArrayFile(sPathFilexporter, aRows)
                        NbreTotal = DecFormat
                        ProgressBar1.Value = ProgressBar1.Minimum
                        Datagridaffiche.Rows.Clear()
                        ProgresMax = UBound(aRows) + 1 - DecFormat
                        m = 0
                        infoListe = New List(Of Integer)
                        infoLigne = New List(Of Integer)
                        For i = DecFormat To UBound(aRows)
                            aCols = Split(aRows(i), sColumnsSepar)
                            Datagridaffiche.RowCount = iLine + 1 - m
                            For jColD = 0 To Datagridaffiche.ColumnCount - 1
                                iColPosition = CInt(Strings.Right(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1), Strings.Len(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1)) - InStr(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1), "(")))
                                If iColPosition <> 0 Then
                                    If iColPosition <= (UBound(aCols) + 1) Then
                                        Datagridaffiche.Item(jColD, iLine - m).Value = Trim(aCols(iColPosition - 1))
                                    Else
                                        Datagridaffiche.Item(jColD, iLine - m).Value = ""
                                    End If
                                Else
                                    Datagridaffiche.Item(jColD, iLine - m).Value = LireFichierFormat(spathFileFormat, Datagridaffiche.Columns(jColD).Name, Formatype)
                                End If
                            Next jColD
                            iLine = iLine + 1
                            If i Mod 10 = 0 Then
                                Me.Refresh()
                                Modification_Integrer_Ecriture(IdentifiantPiece, FormatdeDatefich, Trim(Formatype), LirePUDefaut(spathFileFormat, Trim(Formatype)), LireLigneArticle(spathFileFormat, Trim(Formatype)), FormatQte)
                                m = iLine
                            Else
                                If i = UBound(aRows) Then
                                    Me.Refresh()
                                    Modification_Integrer_Ecriture(IdentifiantPiece, FormatdeDatefich, Trim(Formatype), LirePUDefaut(spathFileFormat, Trim(Formatype)), LireLigneArticle(spathFileFormat, Trim(Formatype)), FormatQte)
                                    m = iLine
                                End If
                            End If
                        Next i
                    Else
                        ErreurJrn.WriteLine("Impossible d' Integrer le Fichier " & sPathFilexporter & " car Erreur de lecture du Fichier Format " & spathFileFormat)
                    End If
                Catch ex As Exception
                    exceptionTrouve = True
                End Try
            Else
                If Trim(Formatype) = "Longueur Fixe" Then
                    Try
                        If AffichFormatintegration(spathFileFormat, Formatype) = True Then
                            aRows = GetArrayFile(sPathFilexporter, aRows)
                            NbreTotal = DecFormat
                            ProgressBar1.Value = ProgressBar1.Minimum
                            Datagridaffiche.Rows.Clear()
                            ProgresMax = UBound(aRows) + 1 - DecFormat
                            m = 0
                            infoListe = New List(Of Integer)
                            infoLigne = New List(Of Integer)
                            For i = DecFormat To UBound(aRows)
                                Datagridaffiche.RowCount = iLine + 1 - m
                                For jColD = 0 To Datagridaffiche.ColumnCount - 1
                                    iColPosition = CInt(Strings.Left(Strings.Right(Datagridaffiche.Columns(jColD).HeaderText, Strings.Len(Datagridaffiche.Columns(jColD).HeaderText) - InStr(Datagridaffiche.Columns(jColD).HeaderText, "[")), InStr(Strings.Right(Datagridaffiche.Columns(jColD).HeaderText, Strings.Len(Datagridaffiche.Columns(jColD).HeaderText) - InStr(Datagridaffiche.Columns(jColD).HeaderText, "[")), "]") - 1))
                                    iColGauchetxt = CInt(Strings.Right(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1), Strings.Len(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1)) - InStr(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1), "(")))
                                    If iColPosition <> 0 Or iColGauchetxt <> 0 Then
                                        Datagridaffiche.Item(jColD, iLine - m).Value = Trim(Strings.Mid(aRows(i), iColPosition, iColGauchetxt))
                                    Else
                                        Datagridaffiche.Item(jColD, iLine - m).Value = LireFichierFormat(spathFileFormat, Datagridaffiche.Columns(jColD).Name, Formatype)
                                    End If
                                Next jColD
                                iLine = iLine + 1
                                If i Mod 10 = 0 Then
                                    Me.Refresh()
                                    Modification_Integrer_Ecriture(IdentifiantPiece, FormatdeDatefich, Trim(Formatype), LirePUDefaut(spathFileFormat, Trim(Formatype)), LireLigneArticle(spathFileFormat, Trim(Formatype)), FormatQte)
                                    m = iLine
                                Else
                                    If i = UBound(aRows) Then
                                        Me.Refresh()
                                        Modification_Integrer_Ecriture(IdentifiantPiece, FormatdeDatefich, Trim(Formatype), LirePUDefaut(spathFileFormat, Trim(Formatype)), LireLigneArticle(spathFileFormat, Trim(Formatype)), FormatQte)
                                        m = iLine
                                    End If
                                End If
                            Next i
                        Else
                            ErreurJrn.WriteLine("Impossible d' Integrer le Fichier " & sPathFilexporter & " car Erreur de lecture du Fichier Format " & spathFileFormat)
                        End If
                    Catch ex As Exception
                        exceptionTrouve = True
                    End Try
                End If
            End If
        End If

    End Sub
    Private Sub Modification_Integrer_Ecriture(ByRef IdentifiantPiece As String, ByRef FormatfichierDate As String, ByRef FormatIntegrer As Object, ByRef Punitaire As String, ByRef IdentifiantArticle As String, ByRef FormatQteU As Integer)
        Me.Cursor = Cursors.WaitCursor
        BT_integrer.Enabled = False
        If Datagridaffiche.RowCount >= 0 Then
            ProgressBar1.Maximum = ProgresMax
            Try
                For numLigne = 0 To Datagridaffiche.RowCount - 1
                    vidage()
                    NbreTotal = NbreTotal + 1
                    Label5.Refresh()
                    Label5.Text = "Modification des Integrations!"
                    For numColonne = 0 To Datagridaffiche.ColumnCount - 1
                        'Entête Document
                        If Datagridaffiche.Columns.Contains(IdentifiantPiece) = True Then
                            If Strings.Len(Trim(Datagridaffiche.Rows(numLigne).Cells(IdentifiantPiece).Value)) <= 8 Then
                                EntetePieceInterne = Formatage_Chaine(Trim(Datagridaffiche.Rows(numLigne).Cells(IdentifiantPiece).Value))
                            Else
                                EntetePieceInterne = Formatage_Chaine(Strings.Left(Trim(Datagridaffiche.Rows(numLigne).Cells(IdentifiantPiece).Value), 8))
                            End If
                        End If
                        If Datagridaffiche.Columns.Contains(IdentifiantArticle) = True Then
                            PieceArticle = Trim(Datagridaffiche.Rows(numLigne).Cells(IdentifiantArticle).Value)
                        End If

                        If Datagridaffiche.Columns(numColonne).Name = "EnteteTyPeDocument" Then
                            EnteteTyPeDocument = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "EnteteCodeAffaire" Then
                            EnteteCodeAffaire = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "EntetePlanAnalytique" Then
                            EntetePlanAnalytique = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "EnteteReference" Then
                            If Strings.Len(Trim(Datagridaffiche.Item(numColonne, numLigne).Value)) <= 17 Then
                                EnteteReference = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                            Else
                                EnteteReference = Strings.Left(Trim(Datagridaffiche.Item(numColonne, numLigne).Value), 17)
                            End If
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "LigneDesignationArticle" Then
                            If Strings.Len(Trim(Datagridaffiche.Item(numColonne, numLigne).Value)) <= 69 Then
                                LigneDesignationArticle = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                            Else
                                LigneDesignationArticle = Strings.Left(Trim(Datagridaffiche.Item(numColonne, numLigne).Value), 69)
                            End If
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "LignePoidsBrut" Then
                            LignePoidsBrut = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "LignePoidsNet" Then
                            LignePoidsNet = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "LignePrixUnitaire" Then
                            LignePrixUnitaire = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns.Contains("LigneQuantite") = True Then
                            LigneQuantite = Trim(Datagridaffiche.Rows(numLigne).Cells("LigneQuantite").Value)
                        End If

                        If Datagridaffiche.Columns(numColonne).Name = "LigneReference" Then
                            If Strings.Len(Trim(Datagridaffiche.Item(numColonne, numLigne).Value)) <= 17 Then
                                LigneReference = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                            Else
                                LigneReference = Strings.Left(Trim(Datagridaffiche.Item(numColonne, numLigne).Value), 17)
                            End If
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "LigneCodeArticle" Then
                            LigneCodeArticle = Formatage_Article(Trim(Datagridaffiche.Item(numColonne, numLigne).Value))
                        End If

                        'RECHERCHE DE L'INTITULE DE L'INFO LIBRE
                        If Trim(FormatIntegrer) = "Longueur Fixe" Then
                            Dim InfoTableau() As String = Split(Trim(Strings.Left(Strings.Right(Datagridaffiche.Columns(numColonne).HeaderText, Strings.Len(Datagridaffiche.Columns(numColonne).HeaderText) - InStr(Datagridaffiche.Columns(numColonne).HeaderText, "{")), Strings.Len(Strings.Right(Datagridaffiche.Columns(numColonne).HeaderText, Strings.Len(Datagridaffiche.Columns(numColonne).HeaderText) - InStr(Datagridaffiche.Columns(numColonne).HeaderText, "{"))) - 1)), "-")
                            If Trim(InfoTableau(0)) = "oui" Then
                                If Trim(InfoTableau(1)) = "F_DOCENTETE" Then
                                    infoListe.Add(numColonne)
                                End If
                                If Trim(InfoTableau(1)) = "F_DOCLIGNE" Then
                                    infoLigne.Add(numColonne)
                                End If
                            End If
                        Else
                            Dim InfoTableau() As String = Split(Trim(Strings.Left(Strings.Right(Datagridaffiche.Columns(numColonne).HeaderText, Strings.Len(Datagridaffiche.Columns(numColonne).HeaderText) - InStr(Datagridaffiche.Columns(numColonne).HeaderText, "[")), Strings.Len(Strings.Right(Datagridaffiche.Columns(numColonne).HeaderText, Strings.Len(Datagridaffiche.Columns(numColonne).HeaderText) - InStr(Datagridaffiche.Columns(numColonne).HeaderText, "["))) - 1)), "-")
                            If Trim(InfoTableau(0)) = "oui" Then
                                If Trim(InfoTableau(1)) = "F_DOCENTETE" Then
                                    infoListe.Add(numColonne)
                                End If
                                If Trim(InfoTableau(1)) = "F_DOCLIGNE" Then
                                    infoLigne.Add(numColonne)
                                End If
                            End If
                        End If
                    Next numColonne
                    'Creation Effective du Document Commercial
                    Document = Nothing
                    If CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQteU, DecimalNomb, DecimalMone)) <> 0 Then

                        If Trim(EnteteTyPeDocument) = "23" Then 'Transfert
                            If BaseCial.FactoryDocumentStock.ExistPiece(DocumentType.DocumentTypeStockVirement, Trim(EntetePieceInterne)) = True Then
                                ModificationTouslespiece(EnteteIntituleDepotOrigine, EntetePieceInterne, EnteteTyPeDocument, Document, infoListe, FormatfichierDate, PieceArticle, Punitaire, IdentifiantArticle)
                            Else
                                ExisteLecture = False
                                exceptionTrouve = True
                                ErreurJrn.WriteLine("Le Mouvement de Transfert N°: " & Trim(EntetePieceInterne) & " n'existe pas dans la base commerciale")
                            End If
                        End If
                    End If
                    NbreLigne = NbreLigne + 1
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label8.Text = NbreLigne & "/" & ProgresMax
                    Label8.Refresh()
                Next numLigne
            Catch ex As Exception
                ExisteLecture = False
                exceptionTrouve = True
                ErreurJrn.WriteLine("Une erreur s'est produit au moment de la lecture du fichier  : " & Trim(EntetePieceInterne))
            End Try
        End If
        Datagridaffiche.Rows.Clear()
        Me.Cursor = Cursors.Default
        BT_integrer.Enabled = True
    End Sub
    Private Sub Modification_Ligne_Article(ByRef CodeArticle As String, ByRef FormatDatefichier As String, ByRef PieceArticle As String, ByRef Punitaire As String, ByRef IdentifiantArticle As String)
        Dim statistAdap As OleDbDataAdapter
        Dim statistDs As DataSet
        Dim statistTab As DataTable
        Dim OleAdaptaterDelete As OleDbDataAdapter
        Dim OleDeleteDataset As DataSet
        Dim OledatableDelete As DataTable
        Dim Tarifsupdate As String = Nothing
        Dim Tarifsupdateperso As String = Nothing
        Dim OleRecherAdapter As OleDbDataAdapter
        Dim OleRecherDataset As DataSet
        Dim OleRechDatable As DataTable
        Dim DossierAdap As OleDbDataAdapter
        Dim DossierDs As DataSet
        Dim DossierTab As DataTable
        Dim FormatQte As Integer = 0
        Dim FormatMnt As Integer = 0
        DossierAdap = New OleDbDataAdapter("select * from P_DOSSIERCIAL", OleSocieteConnect)
        DossierDs = New DataSet
        DossierAdap.Fill(DossierDs)
        DossierTab = DossierDs.Tables(0)
        If DossierTab.Rows.Count <> 0 Then
            If InStr(DossierTab.Rows(0).Item("D_FormatQte"), ",") <> 0 Then
                If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ",")), ".") <> 0 Then
                    FormatQte = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ","))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ",")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ","))), ".")))
                Else
                    FormatQte = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ",")))
                End If
            Else
                If InStr(DossierTab.Rows(0).Item("D_FormatQte"), ".") <> 0 Then
                    If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ".")), ",") <> 0 Then
                        FormatQte = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), "."))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ".")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), "."))), ",")))
                    Else
                        FormatQte = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ".")))
                    End If
                End If
            End If

            If InStr(DossierTab.Rows(0).Item("D_FormatPrix"), ",") <> 0 Then
                If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ",")), ".") <> 0 Then
                    FormatMnt = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ","))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ",")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ","))), ".")))
                Else
                    FormatMnt = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ",")))
                End If
            Else
                If InStr(DossierTab.Rows(0).Item("D_FormatPrix"), ".") <> 0 Then
                    If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ".")), ",") <> 0 Then
                        FormatMnt = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), "."))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ".")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), "."))), ",")))
                    Else
                        FormatMnt = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ".")))
                    End If
                End If
            End If
        End If
        Try
            For Each LigneDocument In Document.FactoryDocumentLigne.List
                If LigneDocument.Article.AR_Ref = Trim(CodeArticle) Then
                    With LigneDocument
                        If Datagridaffiche.Columns.Contains("LigneDesignationArticle") = True Then
                            If Trim(LigneDesignationArticle) <> "" Then
                                .DL_Design = LigneDesignationArticle
                            End If
                        End If
                        If Datagridaffiche.Columns.Contains("LignePoidsNet") = True Then
                            If EstNumeric(Trim(LignePoidsNet), DecimalNomb, DecimalMone) = True Then
                                .DL_PoidsNet = CDbl(RenvoiMontant(Trim(LignePoidsNet), FormatQte, DecimalNomb, DecimalMone))
                            End If
                        End If
                        If Datagridaffiche.Columns.Contains("LignePoidsBrut") = True Then
                            If EstNumeric(Trim(LignePoidsBrut), DecimalNomb, DecimalMone) = True Then
                                .DL_PoidsBrut = CDbl(RenvoiMontant(Trim(LignePoidsBrut), FormatQte, DecimalNomb, DecimalMone))
                            End If
                        End If

                        If Datagridaffiche.Columns.Contains("LigneReference") = True Then
                            If Trim(LigneReference) <> "" Then
                                .DO_Ref = Trim(LigneReference)
                            End If
                        End If
                        If Datagridaffiche.Columns.Contains("LigneQuantite") = True Then
                            If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                                .DL_Qte = CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone))
                            End If
                        End If
                        If Datagridaffiche.Columns.Contains("LignePrixUnitaire") = True Then
                            If EstNumeric(Trim(LignePrixUnitaire), DecimalNomb, DecimalMone) = True Then
                                .DL_PrixUnitaire = CDbl(RenvoiMontant(Trim(LignePrixUnitaire), FormatMnt, DecimalNomb, DecimalMone))
                            End If
                        End If
                        If Punitaire = "oui" Then
                            .WriteDefault()
                        Else
                            .Write()
                        End If
                        If IsNothing(LigneDocument.Article) = False Then
                            ErreurJrn.WriteLine("Code Article : " & Trim(LigneDocument.Article.AR_Ref) & " Modifié Pour la pièce Sage N° : " & Document.DO_Piece & " et Pour la pièce du fichier N° :" & Trim(EntetePieceInterne))
                        Else
                            ErreurJrn.WriteLine("Code Article : " & Trim("Vide") & " Modifié Pour la pièce Sage N° : " & Document.DO_Piece & " et Pour la pièce du fichier N° :" & Trim(EntetePieceInterne))
                        End If
                        'Traitement des Infos Libres
                        Try
                            If infoLigne.Count > 0 Then
                                While infoLigne.Count <> 0
                                    OleAdaptaterDelete = New OleDbDataAdapter("select * From WIT_COL where Libelle='" & Trim(Datagridaffiche.Columns(infoLigne.Item(0)).Name) & "' And InfoLigne=True", OleConnenection)
                                    OleDeleteDataset = New DataSet
                                    OleAdaptaterDelete.Fill(OleDeleteDataset)
                                    OledatableDelete = OleDeleteDataset.Tables(0)
                                    If OledatableDelete.Rows.Count <> 0 Then
                                        'L'info Libre Parametrée par l'utilisateur existe dans Sage
                                        If LigneDocument.InfoLibre.Count <> 0 Then
                                            If IsNothing(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) = False Then
                                                If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                                    statistAdap = New OleDbDataAdapter("select * from cbSysLibre where CB_File='F_DOCLIGNE' and CB_Name ='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "'", OleSocieteConnect)
                                                    statistDs = New DataSet
                                                    statistAdap.Fill(statistDs)
                                                    statistTab = statistDs.Tables(0)
                                                    If statistTab.Rows.Count <> 0 Then
                                                        'Texte
                                                        If statistTab.Rows(0).Item("CB_Type") = 9 Then
                                                            OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                                            OleRecherDataset = New DataSet
                                                            OleRecherAdapter.Fill(OleRecherDataset)
                                                            OleRechDatable = OleRecherDataset.Tables(0)
                                                            If OleRechDatable.Rows.Count <> 0 Then
                                                                If statistTab.Rows(0).Item("CB_Len") >= Strings.Len(Trim(OleRechDatable.Rows(0).Item("Correspond"))) Then
                                                                    LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = Trim(OleRechDatable.Rows(0).Item("Correspond"))
                                                                    LigneDocument.Write()
                                                                End If
                                                            Else
                                                                If statistTab.Rows(0).Item("CB_Len") >= Strings.Len(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value)) Then
                                                                    LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value)
                                                                    LigneDocument.Write()
                                                                End If
                                                            End If
                                                        End If
                                                        'Table
                                                        If statistTab.Rows(0).Item("CB_Type") = 22 Then
                                                            OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                                            OleRecherDataset = New DataSet
                                                            OleRecherAdapter.Fill(OleRecherDataset)
                                                            OleRechDatable = OleRecherDataset.Tables(0)
                                                            If OleRechDatable.Rows.Count <> 0 Then
                                                                If statistTab.Rows(0).Item("CB_Len") > Strings.Len(Trim(OleRechDatable.Rows(0).Item("Correspond"))) Then
                                                                    LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = Trim(OleRechDatable.Rows(0).Item("Correspond"))
                                                                    LigneDocument.Write()
                                                                End If
                                                            Else
                                                                If statistTab.Rows(0).Item("CB_Len") > Strings.Len(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value)) Then
                                                                    LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value)
                                                                    LigneDocument.Write()
                                                                End If
                                                            End If
                                                        End If
                                                        'Montant
                                                        If statistTab.Rows(0).Item("CB_Type") = 20 Then
                                                            If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                                                OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                                                OleRecherDataset = New DataSet
                                                                OleRecherAdapter.Fill(OleRecherDataset)
                                                                OleRechDatable = OleRecherDataset.Tables(0)
                                                                If OleRechDatable.Rows.Count <> 0 Then
                                                                    If EstNumeric(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone) = True Then
                                                                        LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone))
                                                                        LigneDocument.Write()
                                                                    End If
                                                                Else
                                                                    If EstNumeric(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), DecimalNomb, DecimalMone) = True Then
                                                                        LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), DecimalNomb, DecimalMone))
                                                                        LigneDocument.Write()
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                        'Valeur
                                                        If statistTab.Rows(0).Item("CB_Type") = 7 Then
                                                            If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                                                OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                                                OleRecherDataset = New DataSet
                                                                OleRecherAdapter.Fill(OleRecherDataset)
                                                                OleRechDatable = OleRecherDataset.Tables(0)
                                                                If OleRechDatable.Rows.Count <> 0 Then
                                                                    If EstNumeric(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone) = True Then
                                                                        LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone))
                                                                        LigneDocument.Write()
                                                                    End If
                                                                Else
                                                                    If EstNumeric(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), DecimalNomb, DecimalMone) = True Then
                                                                        LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), DecimalNomb, DecimalMone))
                                                                        LigneDocument.Write()
                                                                    End If
                                                                End If
                                                            End If
                                                        End If

                                                        'Date Court
                                                        If statistTab.Rows(0).Item("CB_Type") = 3 Then
                                                            If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                                                OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                                                OleRecherDataset = New DataSet
                                                                OleRecherAdapter.Fill(OleRecherDataset)
                                                                OleRechDatable = OleRecherDataset.Tables(0)
                                                                If OleRechDatable.Rows.Count <> 0 Then
                                                                    If Verificatdate(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then
                                                                        LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = RenvoieDateValide(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier)
                                                                        LigneDocument.Write()
                                                                    End If
                                                                Else
                                                                    If Verificatdate(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then
                                                                        LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = RenvoieDateValide(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), FormatDatefichier)
                                                                        LigneDocument.Write()
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                        'Date Longue
                                                        If statistTab.Rows(0).Item("CB_Type") = 14 Then
                                                            If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                                                OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                                                OleRecherDataset = New DataSet
                                                                OleRecherAdapter.Fill(OleRecherDataset)
                                                                OleRechDatable = OleRecherDataset.Tables(0)
                                                                If OleRechDatable.Rows.Count <> 0 Then
                                                                    If Verificatdate(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then
                                                                        LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = RenvoieDateValide(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier)
                                                                        LigneDocument.Write()
                                                                    End If
                                                                Else
                                                                    If Verificatdate(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then
                                                                        LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = RenvoieDateValide(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), FormatDatefichier)
                                                                        LigneDocument.Write()
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    Else
                                                        If IsNothing(LigneDocument.Article) = False Then
                                                            ErreurJrn.WriteLine("Code Article : " & Trim(LigneDocument.Article.AR_Ref) & " Impossible de traiter l'information libre :" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "  Il est inexistant dans la table CbsysLibre de Sage")
                                                        Else
                                                            ErreurJrn.WriteLine("Code Article : " & Trim("Vide") & " Impossible de traiter l'information libre :" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "  Il est inexistant dans la table CbsysLibre de Sage")
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                'nothing
                                            End If
                                        End If
                                    End If
                                    'L'info Libre Parametrée par l'utilisateur n'existe pas dans Sage
                                    infoLigne.RemoveAt(0)
                                End While
                            End If
                        Catch ex As Exception
                            exceptionTrouve = True
                            If IsNothing(LigneDocument.Article) = False Then
                                ErreurJrn.WriteLine("Code Article : " & Trim(LigneDocument.Article.AR_Ref) & " Erreur de Modification de L'information Libre , vérifiez la longueur de la chaine avec celle paramétrée/ ou la cohérence des données")
                            Else
                                ErreurJrn.WriteLine("Code Article : " & Trim("Vide") & " Erreur de Modification de L'information Libre , vérifiez la longueur de la chaine avec celle paramétrée/ ou la cohérence des données")
                            End If
                        End Try
                    End With
                End If
            Next
        Catch ex As Exception
            exceptionTrouve = True
            ErreurJrn.WriteLine("Code Article : " & Trim(LigneCodeArticle) & " N°Pièce : " & EntetePieceInterne & " Erreur système de Modification de l'article : " & ex.Message)
            ListePiece.Add(EntetePieceInterne)
        End Try
        'l'article existe dans la base
    End Sub
    Private Sub Modification_Creation_Ligne_Article(ByRef FormatDatefichier As String, ByRef PieceArticle As String, ByRef Punitaire As String, ByRef IdentifiantArticle As String)
        Dim fournisseurAdap As OleDbDataAdapter
        Dim fournisseurDs As DataSet
        Dim fournisseurTab As DataTable
        Dim statistAdap As OleDbDataAdapter
        Dim statistDs As DataSet
        Dim statistTab As DataTable
        Dim OleAdaptaterDelete As OleDbDataAdapter
        Dim OleDeleteDataset As DataSet
        Dim OledatableDelete As DataTable
        Dim Tarifsupdate As String = Nothing
        Dim Tarifsupdateperso As String = Nothing
        Dim OleRecherAdapter As OleDbDataAdapter
        Dim OleRecherDataset As DataSet
        Dim OleRechDatable As DataTable
        Dim DossierAdap As OleDbDataAdapter
        Dim DossierDs As DataSet
        Dim DossierTab As DataTable
        Dim FormatQte As Integer = 0
        Dim FormatMnt As Integer = 0
        DossierAdap = New OleDbDataAdapter("select * from P_DOSSIERCIAL", OleSocieteConnect)
        DossierDs = New DataSet
        DossierAdap.Fill(DossierDs)
        DossierTab = DossierDs.Tables(0)
        If DossierTab.Rows.Count <> 0 Then
            If InStr(DossierTab.Rows(0).Item("D_FormatQte"), ",") <> 0 Then
                If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ",")), ".") <> 0 Then
                    FormatQte = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ","))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ",")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ","))), ".")))
                Else
                    FormatQte = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ",")))
                End If
            Else
                If InStr(DossierTab.Rows(0).Item("D_FormatQte"), ".") <> 0 Then
                    If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ".")), ",") <> 0 Then
                        FormatQte = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), "."))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ".")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), "."))), ",")))
                    Else
                        FormatQte = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ".")))
                    End If
                End If
            End If

            If InStr(DossierTab.Rows(0).Item("D_FormatPrix"), ",") <> 0 Then
                If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ",")), ".") <> 0 Then
                    FormatMnt = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ","))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ",")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ","))), ".")))
                Else
                    FormatMnt = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ",")))
                End If
            Else
                If InStr(DossierTab.Rows(0).Item("D_FormatPrix"), ".") <> 0 Then
                    If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ".")), ",") <> 0 Then
                        FormatMnt = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), "."))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ".")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), "."))), ",")))
                    Else
                        FormatMnt = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ".")))
                    End If
                End If
            End If
        End If
        Try
            LigneDocument = Document.FactoryDocumentLigne.Create
            With LigneDocument
                If Datagridaffiche.Columns.Contains("LigneDesignationArticle") = True Then
                    If Trim(LigneDesignationArticle) <> "" Then
                        .DL_Design = LigneDesignationArticle
                    End If
                End If
                If Datagridaffiche.Columns.Contains("LignePoidsNet") = True Then
                    If EstNumeric(Trim(LignePoidsNet), DecimalNomb, DecimalMone) = True Then
                        .DL_PoidsNet = CDbl(RenvoiMontant(Trim(LignePoidsNet), FormatQte, DecimalNomb, DecimalMone))
                    End If
                End If
                If Datagridaffiche.Columns.Contains("LignePoidsBrut") = True Then
                    If EstNumeric(Trim(LignePoidsBrut), DecimalNomb, DecimalMone) = True Then
                        .DL_PoidsBrut = CDbl(RenvoiMontant(Trim(LignePoidsBrut), FormatQte, DecimalNomb, DecimalMone))
                    End If
                End If
                If Datagridaffiche.Columns.Contains("LigneReference") = True Then
                    If Trim(LigneReference) <> "" Then
                        .DO_Ref = Trim(LigneReference)
                    End If
                End If

                If Datagridaffiche.Columns.Contains("LigneCodeArticle") = True Then
                    If BaseCial.FactoryArticle.ExistReference(Trim(LigneCodeArticle)) = True Then
                        If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                            .SetDefaultArticle(BaseCial.FactoryArticle.ReadReference(Trim(LigneCodeArticle)), CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)))
                        End If
                    Else
                        fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<ArticleStock>' and Valeurlue ='" & Join(Split(Trim(LigneCodeArticle), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                        fournisseurDs = New DataSet
                        fournisseurAdap.Fill(fournisseurDs)
                        fournisseurTab = fournisseurDs.Tables(0)
                        If fournisseurTab.Rows.Count <> 0 Then
                            If BaseCial.FactoryArticle.ExistReference(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                                If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                                    .SetDefaultArticle(BaseCial.FactoryArticle.ReadReference(Trim(fournisseurTab.Rows(0).Item("Correspond"))), CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)))
                                End If
                            End If
                        End If
                    End If
                End If

                If Datagridaffiche.Columns.Contains("LignePrixUnitaire") = True Then
                    If EstNumeric(Trim(LignePrixUnitaire), DecimalNomb, DecimalMone) = True Then
                        .DL_PrixUnitaire = CDbl(RenvoiMontant(Trim(LignePrixUnitaire), FormatMnt, DecimalNomb, DecimalMone))
                    End If
                End If
                If Punitaire = "oui" Then
                    .WriteDefault()
                Else
                    .Write()
                End If
                If IsNothing(LigneDocument.Article) = False Then
                    ErreurJrn.WriteLine("Code Article : " & Trim(LigneDocument.Article.AR_Ref) & " Créé Pour la pièce du fichier N° :" & Trim(EntetePieceInterne))
                Else
                    ErreurJrn.WriteLine("Code Article : " & Trim("Vide") & " Créé Pour la pièce du fichier N° :" & Trim(EntetePieceInterne))
                End If
                'Traitement des Infos Libres
                Try
                    If infoLigne.Count > 0 Then
                        While infoLigne.Count <> 0
                            OleAdaptaterDelete = New OleDbDataAdapter("select * From WIT_COL where Libelle='" & Trim(Datagridaffiche.Columns(infoLigne.Item(0)).Name) & "' And InfoLigne=True", OleConnenection)
                            OleDeleteDataset = New DataSet
                            OleAdaptaterDelete.Fill(OleDeleteDataset)
                            OledatableDelete = OleDeleteDataset.Tables(0)
                            If OledatableDelete.Rows.Count <> 0 Then
                                'L'info Libre Parametrée par l'utilisateur existe dans Sage
                                If LigneDocument.InfoLibre.Count <> 0 Then
                                    If IsNothing(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) = False Then
                                        If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                            statistAdap = New OleDbDataAdapter("select * from cbSysLibre where CB_File='F_DOCLIGNE' and CB_Name ='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "'", OleSocieteConnect)
                                            statistDs = New DataSet
                                            statistAdap.Fill(statistDs)
                                            statistTab = statistDs.Tables(0)
                                            If statistTab.Rows.Count <> 0 Then
                                                'Texte
                                                If statistTab.Rows(0).Item("CB_Type") = 9 Then
                                                    OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                                    OleRecherDataset = New DataSet
                                                    OleRecherAdapter.Fill(OleRecherDataset)
                                                    OleRechDatable = OleRecherDataset.Tables(0)
                                                    If OleRechDatable.Rows.Count <> 0 Then
                                                        If statistTab.Rows(0).Item("CB_Len") >= Strings.Len(Trim(OleRechDatable.Rows(0).Item("Correspond"))) Then
                                                            LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = Trim(OleRechDatable.Rows(0).Item("Correspond"))
                                                            LigneDocument.Write()
                                                        End If
                                                    Else
                                                        If statistTab.Rows(0).Item("CB_Len") >= Strings.Len(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value)) Then
                                                            LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value)
                                                            LigneDocument.Write()
                                                        End If
                                                    End If
                                                End If
                                                'Table
                                                If statistTab.Rows(0).Item("CB_Type") = 22 Then
                                                    OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                                    OleRecherDataset = New DataSet
                                                    OleRecherAdapter.Fill(OleRecherDataset)
                                                    OleRechDatable = OleRecherDataset.Tables(0)
                                                    If OleRechDatable.Rows.Count <> 0 Then
                                                        If statistTab.Rows(0).Item("CB_Len") > Strings.Len(Trim(OleRechDatable.Rows(0).Item("Correspond"))) Then
                                                            LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = Trim(OleRechDatable.Rows(0).Item("Correspond"))
                                                            LigneDocument.Write()
                                                        End If
                                                    Else
                                                        If statistTab.Rows(0).Item("CB_Len") > Strings.Len(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value)) Then
                                                            LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value)
                                                            LigneDocument.Write()
                                                        End If
                                                    End If
                                                End If
                                                'Montant
                                                If statistTab.Rows(0).Item("CB_Type") = 20 Then
                                                    If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                                        OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                                        OleRecherDataset = New DataSet
                                                        OleRecherAdapter.Fill(OleRecherDataset)
                                                        OleRechDatable = OleRecherDataset.Tables(0)
                                                        If OleRechDatable.Rows.Count <> 0 Then
                                                            If EstNumeric(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone) = True Then
                                                                LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone))
                                                                LigneDocument.Write()
                                                            End If
                                                        Else
                                                            If EstNumeric(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), DecimalNomb, DecimalMone) = True Then
                                                                LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), DecimalNomb, DecimalMone))
                                                                LigneDocument.Write()
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                                'Valeur
                                                If statistTab.Rows(0).Item("CB_Type") = 7 Then
                                                    If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                                        OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                                        OleRecherDataset = New DataSet
                                                        OleRecherAdapter.Fill(OleRecherDataset)
                                                        OleRechDatable = OleRecherDataset.Tables(0)
                                                        If OleRechDatable.Rows.Count <> 0 Then
                                                            If EstNumeric(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone) = True Then
                                                                LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone))
                                                                LigneDocument.Write()
                                                            End If
                                                        Else
                                                            If EstNumeric(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), DecimalNomb, DecimalMone) = True Then
                                                                LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), DecimalNomb, DecimalMone))
                                                                LigneDocument.Write()
                                                            End If
                                                        End If
                                                    End If
                                                End If

                                                'Date Court
                                                If statistTab.Rows(0).Item("CB_Type") = 3 Then
                                                    If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                                        OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                                        OleRecherDataset = New DataSet
                                                        OleRecherAdapter.Fill(OleRecherDataset)
                                                        OleRechDatable = OleRecherDataset.Tables(0)
                                                        If OleRechDatable.Rows.Count <> 0 Then
                                                            If Verificatdate(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then
                                                                LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = RenvoieDateValide(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier)
                                                                LigneDocument.Write()
                                                            End If
                                                        Else
                                                            If Verificatdate(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then
                                                                LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = RenvoieDateValide(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), FormatDatefichier)
                                                                LigneDocument.Write()
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                                'Date Longue
                                                If statistTab.Rows(0).Item("CB_Type") = 14 Then
                                                    If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                                        OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                                        OleRecherDataset = New DataSet
                                                        OleRecherAdapter.Fill(OleRecherDataset)
                                                        OleRechDatable = OleRecherDataset.Tables(0)
                                                        If OleRechDatable.Rows.Count <> 0 Then
                                                            If Verificatdate(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then
                                                                LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = RenvoieDateValide(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier)
                                                                LigneDocument.Write()
                                                            End If
                                                        Else
                                                            If Verificatdate(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then
                                                                LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = RenvoieDateValide(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), FormatDatefichier)
                                                                LigneDocument.Write()
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If IsNothing(LigneDocument.Article) = False Then
                                                    ErreurJrn.WriteLine("Code Article : " & Trim(LigneDocument.Article.AR_Ref) & " Impossible de traiter l'information libre :" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "  Il est inexistant dans la table CbsysLibre de Sage")
                                                Else
                                                    ErreurJrn.WriteLine("Code Article : " & Trim("Vide") & " Impossible de traiter l'information libre :" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "  Il est inexistant dans la table CbsysLibre de Sage")
                                                End If
                                            End If
                                        End If
                                    Else
                                        'nothing
                                    End If
                                End If
                            End If
                            'L'info Libre Parametrée par l'utilisateur n'existe pas dans Sage
                            infoLigne.RemoveAt(0)
                        End While
                    End If
                Catch ex As Exception
                    exceptionTrouve = True
                    If IsNothing(LigneDocument.Article) = False Then
                        ErreurJrn.WriteLine("Code Article : " & Trim(LigneDocument.Article.AR_Ref) & " Erreur de Création de L'information Libre , vérifiez la longueur de la chaine avec celle paramétrée/ ou la cohérence des données")
                    Else
                        ErreurJrn.WriteLine("Code Article : " & Trim("Vide") & " Erreur de Création de L'information Libre , vérifiez la longueur de la chaine avec celle paramétrée/ ou la cohérence des données")
                    End If
                End Try

            End With
        Catch ex As Exception
            exceptionTrouve = True
            ErreurJrn.WriteLine("Code Article : " & Trim(LigneCodeArticle) & " N°Pièce : " & EntetePieceInterne & " Erreur système de Création de l'article : " & ex.Message)
            ListePiece.Add(EntetePieceInterne)
        End Try
    End Sub
    Private Sub Modification_Entete_Document(ByRef typedoc As String, ByRef FormatDatefichier As String)
        Dim fournisseurAdap As OleDbDataAdapter
        Dim fournisseurDs As DataSet
        Dim fournisseurTab As DataTable
        Dim statistAdap As OleDbDataAdapter
        Dim statistDs As DataSet
        Dim statistTab As DataTable
        Dim OleAdaptaterDelete As OleDbDataAdapter
        Dim OleDeleteDataset As DataSet
        Dim OledatableDelete As DataTable
        Dim OleRecherAdapter As OleDbDataAdapter
        Dim OleRecherDataset As DataSet
        Dim OleRechDatable As DataTable
        Dim DossierAdap As OleDbDataAdapter
        Dim DossierDs As DataSet
        Dim DossierTab As DataTable
        Dim FormatQte As Integer = 0
        Dim FormatMnt As Integer = 0
        DossierAdap = New OleDbDataAdapter("select * from P_DOSSIERCIAL", OleSocieteConnect)
        DossierDs = New DataSet
        DossierAdap.Fill(DossierDs)
        DossierTab = DossierDs.Tables(0)
        If DossierTab.Rows.Count <> 0 Then
            If InStr(DossierTab.Rows(0).Item("D_FormatQte"), ",") <> 0 Then
                If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ",")), ".") <> 0 Then
                    FormatQte = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ","))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ",")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ","))), ".")))
                Else
                    FormatQte = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ",")))
                End If
            Else
                If InStr(DossierTab.Rows(0).Item("D_FormatQte"), ".") <> 0 Then
                    If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ".")), ",") <> 0 Then
                        FormatQte = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), "."))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ".")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), "."))), ",")))
                    Else
                        FormatQte = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ".")))
                    End If
                End If
            End If

            If InStr(DossierTab.Rows(0).Item("D_FormatPrix"), ",") <> 0 Then
                If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ",")), ".") <> 0 Then
                    FormatMnt = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ","))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ",")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ","))), ".")))
                Else
                    FormatMnt = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ",")))
                End If
            Else
                If InStr(DossierTab.Rows(0).Item("D_FormatPrix"), ".") <> 0 Then
                    If InStr(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ".")), ",") <> 0 Then
                        FormatMnt = Len(Strings.Right(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), "."))), Len(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ".")))) - InStr(Trim(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), "."))), ",")))
                    Else
                        FormatMnt = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ".")))
                    End If
                End If
            End If
        End If
        With Document
            If Datagridaffiche.Columns.Contains("EntetePlanAnalytique") = True Then
                If Datagridaffiche.Columns.Contains("EnteteCodeAffaire") = True Then
                    If BaseCpta.FactoryAnalytique.ExistIntitule(Trim(EntetePlanAnalytique)) = True Then
                        PlanAna = BaseCpta.FactoryAnalytique.ReadIntitule(Trim(EntetePlanAnalytique))
                        If BaseCpta.FactoryCompteA.ExistNumero(PlanAna, Trim(EnteteCodeAffaire)) = True Then
                            .CompteA = BaseCpta.FactoryCompteA.ReadNumero(PlanAna, Trim(EnteteCodeAffaire))
                        Else
                            statistAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Section Analytique>' and Valeurlue ='" & Join(Split(Trim(EnteteCodeAffaire), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                            statistDs = New DataSet
                            statistAdap.Fill(statistDs)
                            statistTab = statistDs.Tables(0)
                            If statistTab.Rows.Count <> 0 Then
                                If BaseCpta.FactoryCompteA.ExistNumero(PlanAna, Trim(statistTab.Rows(0).Item("Correspond"))) = True Then
                                    .CompteA = BaseCpta.FactoryCompteA.ReadNumero(PlanAna, Trim(statistTab.Rows(0).Item("Correspond")))
                                End If
                            End If
                        End If
                    Else
                        fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Plan Analytique>' and Valeurlue ='" & Join(Split(Trim(EntetePlanAnalytique), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                        fournisseurDs = New DataSet
                        fournisseurAdap.Fill(fournisseurDs)
                        fournisseurTab = fournisseurDs.Tables(0)
                        If fournisseurTab.Rows.Count > 0 Then
                            If BaseCpta.FactoryAnalytique.ExistIntitule(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                                PlanAna = BaseCpta.FactoryAnalytique.ReadIntitule(Trim(fournisseurTab.Rows(0).Item("Correspond")))
                                If BaseCpta.FactoryCompteA.ExistNumero(PlanAna, Trim(EnteteCodeAffaire)) = True Then
                                    .CompteA = BaseCpta.FactoryCompteA.ReadNumero(PlanAna, Trim(EnteteCodeAffaire))
                                Else
                                    statistAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Section Analytique>' and Valeurlue ='" & Join(Split(Trim(EnteteCodeAffaire), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                    statistDs = New DataSet
                                    statistAdap.Fill(statistDs)
                                    statistTab = statistDs.Tables(0)
                                    If statistTab.Rows.Count <> 0 Then
                                        If BaseCpta.FactoryCompteA.ExistNumero(PlanAna, Trim(statistTab.Rows(0).Item("Correspond"))) = True Then
                                            .CompteA = BaseCpta.FactoryCompteA.ReadNumero(PlanAna, Trim(statistTab.Rows(0).Item("Correspond")))
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If

            If Datagridaffiche.Columns.Contains("EnteteReference") = True Then
                If Trim(EnteteReference) <> "" Then
                    .DO_Ref = EnteteReference
                End If
            End If
            .Write()
            ErreurJrn.WriteLine("-----------------------------------------------------------------------------------------------------")
            ErreurJrn.WriteLine("")
            If typedoc = "23" Then
                ErreurJrn.WriteLine("Mouvement de Transfert N° : " & Trim(Document.DO_Piece) & " Modifié Pour la pièce N° :" & Trim(EntetePieceInterne))
            End If
            'Traitement des Infos Libres
            Try
                If infoListe.Count > 0 Then
                    While infoListe.Count <> 0
                        OleAdaptaterDelete = New OleDbDataAdapter("select * From WIT_COL where Libelle='" & Trim(Datagridaffiche.Columns(infoListe.Item(0)).Name) & "' And Libre=True", OleConnenection)
                        OleDeleteDataset = New DataSet
                        OleAdaptaterDelete.Fill(OleDeleteDataset)
                        OledatableDelete = OleDeleteDataset.Tables(0)
                        If OledatableDelete.Rows.Count <> 0 Then
                            'L'info Libre Parametrée par l'utilisateur existe dans Sage
                            If Document.InfoLibre.Count <> 0 Then
                                If IsNothing(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) = False Then
                                    If Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) <> "" Then
                                        statistAdap = New OleDbDataAdapter("select * from cbSysLibre where CB_File='F_DOCENTETE' and CB_Name ='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "'", OleSocieteConnect)
                                        statistDs = New DataSet
                                        statistAdap.Fill(statistDs)
                                        statistTab = statistDs.Tables(0)
                                        If statistTab.Rows.Count <> 0 Then
                                            'Texte
                                            If statistTab.Rows(0).Item("CB_Type") = 9 Then
                                                OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                                OleRecherDataset = New DataSet
                                                OleRecherAdapter.Fill(OleRecherDataset)
                                                OleRechDatable = OleRecherDataset.Tables(0)
                                                If OleRechDatable.Rows.Count <> 0 Then
                                                    If statistTab.Rows(0).Item("CB_Len") >= Strings.Len(Trim(OleRechDatable.Rows(0).Item("Correspond"))) Then
                                                        Document.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = Trim(OleRechDatable.Rows(0).Item("Correspond"))
                                                        Document.Write()
                                                    End If
                                                Else
                                                    If statistTab.Rows(0).Item("CB_Len") >= Strings.Len(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value)) Then
                                                        Document.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value)
                                                        Document.Write()
                                                    End If
                                                End If
                                            End If
                                            'Table
                                            If statistTab.Rows(0).Item("CB_Type") = 22 Then
                                                OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                                OleRecherDataset = New DataSet
                                                OleRecherAdapter.Fill(OleRecherDataset)
                                                OleRechDatable = OleRecherDataset.Tables(0)
                                                If OleRechDatable.Rows.Count <> 0 Then
                                                    If statistTab.Rows(0).Item("CB_Len") > Strings.Len(Trim(OleRechDatable.Rows(0).Item("Correspond"))) Then
                                                        Document.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = Trim(OleRechDatable.Rows(0).Item("Correspond"))
                                                        Document.Write()
                                                    End If
                                                Else
                                                    If statistTab.Rows(0).Item("CB_Len") > Strings.Len(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value)) Then
                                                        Document.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value)
                                                        Document.Write()
                                                    End If
                                                End If
                                            End If
                                            'Montant
                                            If statistTab.Rows(0).Item("CB_Type") = 20 Then
                                                If Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) <> "" Then
                                                    OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                                    OleRecherDataset = New DataSet
                                                    OleRecherAdapter.Fill(OleRecherDataset)
                                                    OleRechDatable = OleRecherDataset.Tables(0)
                                                    If OleRechDatable.Rows.Count <> 0 Then
                                                        If EstNumeric(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone) = True Then
                                                            Document.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone))
                                                            Document.Write()
                                                        End If
                                                    Else
                                                        If EstNumeric(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), DecimalNomb, DecimalMone) = True Then
                                                            Document.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), DecimalNomb, DecimalMone))
                                                            Document.Write()
                                                        End If
                                                    End If
                                                End If
                                            End If
                                            'Valeur
                                            If statistTab.Rows(0).Item("CB_Type") = 7 Then
                                                If Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) <> "" Then
                                                    OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                                    OleRecherDataset = New DataSet
                                                    OleRecherAdapter.Fill(OleRecherDataset)
                                                    OleRechDatable = OleRecherDataset.Tables(0)
                                                    If OleRechDatable.Rows.Count <> 0 Then
                                                        If EstNumeric(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone) = True Then
                                                            Document.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone))
                                                            Document.Write()
                                                        End If
                                                    Else
                                                        If EstNumeric(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), DecimalNomb, DecimalMone) = True Then
                                                            Document.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), DecimalNomb, DecimalMone))
                                                            Document.Write()
                                                        End If
                                                    End If
                                                End If
                                            End If

                                            'Date Court
                                            If statistTab.Rows(0).Item("CB_Type") = 3 Then
                                                If Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) <> "" Then
                                                    OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                                    OleRecherDataset = New DataSet
                                                    OleRecherAdapter.Fill(OleRecherDataset)
                                                    OleRechDatable = OleRecherDataset.Tables(0)
                                                    If OleRechDatable.Rows.Count <> 0 Then
                                                        If Verificatdate(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier, Datagridaffiche.Columns(infoListe.Item(0)).Name) = True Then
                                                            Document.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = RenvoieDateValide(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier)
                                                            Document.Write()
                                                        End If
                                                    Else
                                                        If Verificatdate(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), FormatDatefichier, Datagridaffiche.Columns(infoListe.Item(0)).Name) = True Then
                                                            Document.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = RenvoieDateValide(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), FormatDatefichier)
                                                            Document.Write()
                                                        End If
                                                    End If
                                                End If
                                            End If
                                            'Date Longue
                                            If statistTab.Rows(0).Item("CB_Type") = 14 Then
                                                If Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) <> "" Then
                                                    OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                                    OleRecherDataset = New DataSet
                                                    OleRecherAdapter.Fill(OleRecherDataset)
                                                    OleRechDatable = OleRecherDataset.Tables(0)
                                                    If OleRechDatable.Rows.Count <> 0 Then
                                                        If Verificatdate(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier, Datagridaffiche.Columns(infoListe.Item(0)).Name) = True Then
                                                            Document.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = RenvoieDateValide(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier)
                                                            Document.Write()
                                                        End If
                                                    Else
                                                        If Verificatdate(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), FormatDatefichier, Datagridaffiche.Columns(infoListe.Item(0)).Name) = True Then
                                                            Document.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = RenvoieDateValide(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), FormatDatefichier)
                                                            Document.Write()
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        Else

                                            If typedoc = "23" Then
                                                ErreurJrn.WriteLine("Mouvement de Transfert N° : " & Trim(Document.DO_Piece) & " Impossible de traiter l'information libre de type Date Longue :" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "  De  valeur entrée '" & Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) & " dans Sage")

                                            End If

                                            If typedoc = "23" Then
                                                ErreurJrn.WriteLine("Mouvement de Transfert N° : " & Trim(Document.DO_Piece) & " Impossible de traiter l'information libre :" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "  Il est inexistant dans la table CbsysLibre de Sage")
                                            End If
                                        End If
                                    End If
                                Else
                                    'nothing
                                End If
                            End If
                        End If
                        'L'info Libre Parametrée par l'utilisateur n'existe pas dans Sage
                        infoListe.RemoveAt(0)
                    End While
                End If
            Catch ex As Exception
                exceptionTrouve = True
                If typedoc = "23" Then
                    ErreurJrn.WriteLine("Mouvement de Transfert N° : " & Trim(Document.DO_Piece) & " Erreur de Modification de L'information Libre , vérifiez la longueur de la chaine avec celle paramétrée/ ou la cohérence des données")
                End If
            End Try
        End With
    End Sub
    Private Sub ModificationTouslespiece(ByRef EnteteIntituleDepotOrigine As String, ByRef EntetePieceInterne As String, ByRef EnteteTyPeDocument As String, ByRef Document As IBODocumentStock3, ByRef infoListe As List(Of Integer), ByRef FormatDatefichier As String, ByRef PieceArticle As String, ByRef Punitaire As String, ByRef IdentifiantArticle As String)
        Dim fournisseurAdap As OleDbDataAdapter
        Dim fournisseurDs As DataSet
        Dim fournisseurTab As DataTable
        Dim TarifrsAdap As OleDbDataAdapter
        Dim TarifrsDs As DataSet
        Dim TarifrsTab As DataTable
        Dim Tarifsupdate As String = Nothing
        Dim Tarifsupdateperso As String = Nothing
        If Trim(EntetePieceInterne) = Trim(EntetePiecePrecedent) Then
            If IsNothing(Document) = True Then

                If Trim(EnteteTyPeDocument) = "23" Then
                    Try
                        Document = BaseCial.FactoryDocumentStock.ReadPiece(DocumentType.DocumentTypeStockVirement, Trim(EntetePieceInterne))
                        Modification_Entete_Document(EnteteTyPeDocument, FormatDatefichier)
                    Catch ex As Exception
                        exceptionTrouve = True
                        ErreurJrn.WriteLine("Erreur de Modification Entête du Mouvement de Transfert N°Pièce Fchier : " & EntetePieceInterne & " Erreur système : " & ex.Message)
                        ListePiece.Add(EntetePieceInterne)
                    End Try
                End If
                If IsNothing(Document) = False Then
                    If Datagridaffiche.Columns.Contains("LigneCodeArticle") = True Then
                        If BaseCial.FactoryArticle.ExistReference(Trim(LigneCodeArticle)) = True Then
                            If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                                Tarifsupdateperso = Nothing
                                TarifrsAdap = New OleDbDataAdapter("select * from F_DOCLIGNE where DO_Piece ='" & Trim(Document.DO_Piece) & "' AND AR_Ref='" & Join(Split(Trim(LigneCodeArticle), "'"), "''") & "'", OleSocieteConnect)
                                TarifrsDs = New DataSet
                                TarifrsAdap.Fill(TarifrsDs)
                                TarifrsTab = TarifrsDs.Tables(0)
                                If TarifrsTab.Rows.Count = 1 Then
                                    Modification_Ligne_Article(Join(Split(Trim(LigneCodeArticle), "'"), "''"), FormatDatefichier, PieceArticle, Punitaire, IdentifiantArticle)
                                    'l'article existe dans la base
                                Else
                                    If TarifrsTab.Rows.Count = 0 Then
                                        Modification_Creation_Ligne_Article(FormatDatefichier, PieceArticle, Punitaire, IdentifiantArticle)
                                        'l article n'existe pas dans la pièce 
                                    End If
                                End If
                            End If
                        Else
                            fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<ArticleStock>' and Valeurlue ='" & Join(Split(Trim(LigneCodeArticle), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                            fournisseurDs = New DataSet
                            fournisseurAdap.Fill(fournisseurDs)
                            fournisseurTab = fournisseurDs.Tables(0)
                            If fournisseurTab.Rows.Count <> 0 Then
                                If BaseCial.FactoryArticle.ExistReference(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                                    If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                                        Tarifsupdateperso = Nothing
                                        TarifrsAdap = New OleDbDataAdapter("select * from F_DOCLIGNE where DO_Piece ='" & Trim(Document.DO_Piece) & "' AND AR_Ref='" & Join(Split(Trim(fournisseurTab.Rows(0).Item("Correspond")), "'"), "''") & "'", OleSocieteConnect)
                                        TarifrsDs = New DataSet
                                        TarifrsAdap.Fill(TarifrsDs)
                                        TarifrsTab = TarifrsDs.Tables(0)
                                        If TarifrsTab.Rows.Count = 1 Then
                                            Modification_Ligne_Article(Join(Split(Trim(fournisseurTab.Rows(0).Item("Correspond")), "'"), "''"), FormatDatefichier, PieceArticle, Punitaire, IdentifiantArticle)
                                            'l'article existe dans la base
                                        Else
                                            If TarifrsTab.Rows.Count = 0 Then
                                                Modification_Creation_Ligne_Article(FormatDatefichier, PieceArticle, Punitaire, IdentifiantArticle)
                                                'l'article n'existe pas dans la pièce
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else 'document nothing
                If Datagridaffiche.Columns.Contains("LigneCodeArticle") = True Then
                    If BaseCial.FactoryArticle.ExistReference(Trim(LigneCodeArticle)) = True Then
                        If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                            Tarifsupdateperso = Nothing
                            TarifrsAdap = New OleDbDataAdapter("select * from F_DOCLIGNE where DO_Piece ='" & Trim(Document.DO_Piece) & "' AND AR_Ref='" & Join(Split(Trim(LigneCodeArticle), "'"), "''") & "'", OleSocieteConnect)
                            TarifrsDs = New DataSet
                            TarifrsAdap.Fill(TarifrsDs)
                            TarifrsTab = TarifrsDs.Tables(0)
                            If TarifrsTab.Rows.Count = 1 Then
                                Modification_Ligne_Article(Join(Split(Trim(LigneCodeArticle), "'"), "''"), FormatDatefichier, PieceArticle, Punitaire, IdentifiantArticle)
                                'l'article existe dans la base
                            Else
                                If TarifrsTab.Rows.Count = 0 Then
                                    Modification_Creation_Ligne_Article(FormatDatefichier, PieceArticle, Punitaire, IdentifiantArticle)
                                    'l article n'existe pas dans la pièce 
                                End If
                            End If
                        End If
                    Else
                        fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<ArticleStock>' and Valeurlue ='" & Join(Split(Trim(LigneCodeArticle), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                        fournisseurDs = New DataSet
                        fournisseurAdap.Fill(fournisseurDs)
                        fournisseurTab = fournisseurDs.Tables(0)
                        If fournisseurTab.Rows.Count <> 0 Then
                            If BaseCial.FactoryArticle.ExistReference(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                                If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                                    Tarifsupdateperso = Nothing
                                    TarifrsAdap = New OleDbDataAdapter("select * from F_DOCLIGNE where DO_Piece ='" & Trim(Document.DO_Piece) & "' AND AR_Ref='" & Join(Split(Trim(fournisseurTab.Rows(0).Item("Correspond")), "'"), "''") & "'", OleSocieteConnect)
                                    TarifrsDs = New DataSet
                                    TarifrsAdap.Fill(TarifrsDs)
                                    TarifrsTab = TarifrsDs.Tables(0)
                                    If TarifrsTab.Rows.Count = 1 Then
                                        Modification_Ligne_Article(Join(Split(Trim(fournisseurTab.Rows(0).Item("Correspond")), "'"), "''"), FormatDatefichier, PieceArticle, Punitaire, IdentifiantArticle)
                                        'l'article existe dans la base
                                    Else
                                        If TarifrsTab.Rows.Count = 0 Then
                                            Modification_Creation_Ligne_Article(FormatDatefichier, PieceArticle, Punitaire, IdentifiantArticle)
                                            'l'article n'existe pas dans la pièce
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Else
            ' piece precedent <> piece en cours
            If IsNothing(Document) = True Then

                If Trim(EnteteTyPeDocument) = "23" Then
                    Try
                        Document = BaseCial.FactoryDocumentStock.ReadPiece(DocumentType.DocumentTypeStockVirement, Trim(EntetePieceInterne))
                        Modification_Entete_Document(EnteteTyPeDocument, FormatDatefichier)
                    Catch ex As Exception
                        exceptionTrouve = True
                        ErreurJrn.WriteLine("Erreur de Modification Entête du Mouvement de Transfert N°Pièce Fchier : " & EntetePieceInterne & " Erreur système : " & ex.Message)
                        ListePiece.Add(EntetePieceInterne)
                    End Try
                End If
                If IsNothing(Document) = False Then
                    If Datagridaffiche.Columns.Contains("LigneCodeArticle") = True Then
                        If BaseCial.FactoryArticle.ExistReference(Trim(LigneCodeArticle)) = True Then
                            If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                                Tarifsupdateperso = Nothing
                                TarifrsAdap = New OleDbDataAdapter("select * from F_DOCLIGNE where DO_Piece ='" & Trim(Document.DO_Piece) & "' AND AR_Ref='" & Join(Split(Trim(LigneCodeArticle), "'"), "''") & "'", OleSocieteConnect)
                                TarifrsDs = New DataSet
                                TarifrsAdap.Fill(TarifrsDs)
                                TarifrsTab = TarifrsDs.Tables(0)
                                If TarifrsTab.Rows.Count = 1 Then
                                    Modification_Ligne_Article(Join(Split(Trim(LigneCodeArticle), "'"), "''"), FormatDatefichier, PieceArticle, Punitaire, IdentifiantArticle)
                                    'l'article existe dans la base
                                Else
                                    If TarifrsTab.Rows.Count = 0 Then
                                        Modification_Creation_Ligne_Article(FormatDatefichier, PieceArticle, Punitaire, IdentifiantArticle)
                                        'l article n'existe pas dans la pièce 
                                    End If
                                End If
                            End If
                        Else
                            fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<ArticleStock>' and Valeurlue ='" & Join(Split(Trim(LigneCodeArticle), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                            fournisseurDs = New DataSet
                            fournisseurAdap.Fill(fournisseurDs)
                            fournisseurTab = fournisseurDs.Tables(0)
                            If fournisseurTab.Rows.Count <> 0 Then
                                If BaseCial.FactoryArticle.ExistReference(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                                    If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                                        Tarifsupdateperso = Nothing
                                        TarifrsAdap = New OleDbDataAdapter("select * from F_DOCLIGNE where DO_Piece ='" & Trim(Document.DO_Piece) & "' AND AR_Ref='" & Join(Split(Trim(fournisseurTab.Rows(0).Item("Correspond")), "'"), "''") & "'", OleSocieteConnect)
                                        TarifrsDs = New DataSet
                                        TarifrsAdap.Fill(TarifrsDs)
                                        TarifrsTab = TarifrsDs.Tables(0)
                                        If TarifrsTab.Rows.Count = 1 Then
                                            Modification_Ligne_Article(Join(Split(Trim(fournisseurTab.Rows(0).Item("Correspond")), "'"), "''"), FormatDatefichier, PieceArticle, Punitaire, IdentifiantArticle)
                                            'l'article existe dans la base
                                        Else
                                            If TarifrsTab.Rows.Count = 0 Then
                                                Modification_Creation_Ligne_Article(FormatDatefichier, PieceArticle, Punitaire, IdentifiantArticle)
                                                'l'article n'existe pas dans la pièce
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        EntetePiecePrecedent = EntetePieceInterne
    End Sub
    Private Sub RecuperationEnregistrementModifié(ByRef sPathFilexporter As String, ByRef spathFileFormat As String, ByRef typeFormat As String, ByRef Base_Excel As String, ByRef sColumnsSepar As String, ByRef ListePiece As List(Of String), ByRef IdentifiantPiece As String)
        Dim OleAdaptater As OleDbDataAdapter
        Dim OleAfficheDataset As DataSet
        Dim Oledatable As DataTable
        Dim jColD As Integer
        Dim aRows(), aCols(), Dataname(), FichierRecup As String
        Dim iColPosition, iColGauchetxt As Integer
        Dim i, j As Integer
        aRows = Nothing
        Dataname = Split(sPathFilexporter, "\")
        FichierRecup = "Recup_" & Strings.Right(DateAndTime.Year(Now), 2) & "" & Format(DateAndTime.Month(Now), "00") & "" & Format(DateAndTime.Day(Now), "00") & "[" & "" & Format(DateAndTime.Hour(Now), "00") & "-" & Format(DateAndTime.Minute(Now), "00") & "-" & Format(DateAndTime.Second(Now), "00") & "]" & Dataname(UBound(Dataname))
        If ListePiece.Count <> 0 Then
            If Trim(typeFormat) = "Excel" Then
                'Dim Cnapplica As New Microsoft.Office.Interop.Excel.Application
                'Dim Cnbook As Microsoft.Office.Interop.Excel.Workbook
                'Dim Cnsheet As Microsoft.Office.Interop.Excel.Worksheet
                'Try
                '    If AffichFormatintegration(spathFileFormat, typeFormat) = True Then
                '        Datagridaffiche.Rows.Clear()
                '        OleAdaptater = New OleDbDataAdapter("select * from [" & Base_Excel & "$] ", OleExcelConnected)
                '        OleAfficheDataset = New DataSet
                '        OleAdaptater.Fill(OleAfficheDataset)
                '        Oledatable = OleAfficheDataset.Tables(0)
                '        If Oledatable.Rows.Count <> 0 Then
                '            Cnapplica = CreateObject("Excel.Application")
                '            Cnbook = Cnapplica.Workbooks.Add
                '            Cnsheet = Cnbook.Worksheets.Add
                '            For i = Cnbook.Sheets.Count To 1 Step -1
                '                If Cnbook.Sheets(i).name() = Base_Excel Then
                '                    Cnbook.Worksheets(i).Delete()
                '                End If
                '            Next i
                '            Cnsheet.Name = Base_Excel
                '            ProgressBar1.Value = ProgressBar1.Minimum
                '            ProgressBar1.Maximum = Oledatable.Rows.Count - DecFormat
                '            j = 0 + CInt(DecFormat)
                '            For i = DecFormat To Oledatable.Rows.Count - 1
                '                If Datagridaffiche.Columns.Contains(IdentifiantPiece) = True Then
                '                    iColPosition = CInt(Strings.Right(Strings.Left(Datagridaffiche.Columns(IdentifiantPiece).HeaderText, InStr(Datagridaffiche.Columns(IdentifiantPiece).HeaderText, ")") - 1), Strings.Len(Strings.Left(Datagridaffiche.Columns(IdentifiantPiece).HeaderText, InStr(Datagridaffiche.Columns(IdentifiantPiece).HeaderText, ")") - 1)) - InStr(Strings.Left(Datagridaffiche.Columns(IdentifiantPiece).HeaderText, InStr(Datagridaffiche.Columns(IdentifiantPiece).HeaderText, ")") - 1), "(")))
                '                    If iColPosition <= Oledatable.Columns.Count Then
                '                        If iColPosition <> 0 Then
                '                            If Convert.IsDBNull(Oledatable.Rows(i).Item(iColPosition - 1)) = False Then
                '                                If ListePiece.Contains(Trim(Oledatable.Rows(i).Item(iColPosition - 1))) = True Then
                '                                    j = j + 1
                '                                    For jColD = 0 To Oledatable.Columns.Count - 1
                '                                        If Convert.IsDBNull(Oledatable.Rows(i).Item(jColD)) = False Then
                '                                            Cnsheet.Cells(j, jColD + 1) = Oledatable.Rows(i).Item(jColD)
                '                                        End If
                '                                    Next jColD
                '                                End If
                '                            End If
                '                        Else
                '                            If ListePiece.Contains(LireFichierFormat(spathFileFormat, IdentifiantPiece, typeFormat)) = True Then
                '                                j = j + 1
                '                                For jColD = 0 To Oledatable.Columns.Count - 1
                '                                    If Convert.IsDBNull(Oledatable.Rows(i).Item(jColD)) = False Then
                '                                        Cnsheet.Cells(j, jColD + 1) = Oledatable.Rows(i).Item(jColD)
                '                                    End If
                '                                Next jColD
                '                            End If
                '                        End If
                '                    End If
                '                End If
                '                ProgressBar1.Value = ProgressBar1.Value + 1
                '            Next i
                '            Cnbook.SaveCopyAs(PathsFileRecuperer & "" & FichierRecup)
                '            Cnapplica.DefaultSaveFormat = Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel5
                '            Cnbook.Close(False) 'Ferme le classeur
                '            Cnapplica.Quit()
                '            Cnbook = Nothing
                '            Cnapplica = Nothing
                '        End If
                '    Else
                '        ErreurJrn.WriteLine("Impossible d'integrer le Fichier " & sPathFilexporter & " car Erreur de lecture du Fichier Format " & spathFileFormat)
                '    End If
                'Catch ex As Exception
                '    exceptionTrouve = True
                'End Try
            Else
                If Trim(typeFormat) = "Délimité" Or Trim(typeFormat) = "Tabulation" Or Trim(typeFormat) = "Pipe" Then
                    Try
                        Error_journal = File.AppendText(PathsFileRecuperer & "" & FichierRecup)
                        If AffichFormatintegration(spathFileFormat, typeFormat) = True Then
                            aRows = GetArrayFile(sPathFilexporter, aRows)
                            Datagridaffiche.Rows.Clear()
                            ProgressBar1.Value = ProgressBar1.Minimum
                            ProgressBar1.Maximum = UBound(aRows) + 1 - DecFormat
                            For i = DecFormat To UBound(aRows)
                                aCols = Split(aRows(i), sColumnsSepar)
                                If Datagridaffiche.Columns.Contains(IdentifiantPiece) = True Then
                                    iColPosition = CInt(Strings.Right(Strings.Left(Datagridaffiche.Columns(IdentifiantPiece).HeaderText, InStr(Datagridaffiche.Columns(IdentifiantPiece).HeaderText, ")") - 1), Strings.Len(Strings.Left(Datagridaffiche.Columns(IdentifiantPiece).HeaderText, InStr(Datagridaffiche.Columns(IdentifiantPiece).HeaderText, ")") - 1)) - InStr(Strings.Left(Datagridaffiche.Columns(IdentifiantPiece).HeaderText, InStr(Datagridaffiche.Columns(IdentifiantPiece).HeaderText, ")") - 1), "(")))
                                    If iColPosition <> 0 Then
                                        If iColPosition <= (UBound(aCols) + 1) Then
                                            If Strings.Len(Trim(aCols(iColPosition - 1))) <= 8 Then
                                                If ListePiece.Contains(Formatage_Chaine(Trim(aCols(iColPosition - 1)))) = True Then
                                                    Error_journal.WriteLine(aRows(i))
                                                End If
                                            Else
                                                If ListePiece.Contains(Formatage_Chaine(Strings.Left(Trim(aCols(iColPosition - 1)), 8))) = True Then
                                                    Error_journal.WriteLine(aRows(i))
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ListePiece.Contains(LireFichierFormat(spathFileFormat, IdentifiantPiece, typeFormat)) = True Then
                                            Error_journal.WriteLine(aRows(i))
                                        End If
                                    End If
                                End If
                                ProgressBar1.Value = ProgressBar1.Value + 1
                            Next i
                            Error_journal.Close()
                        Else
                            ErreurJrn.WriteLine("Impossible d' Integrer le Fichier " & sPathFilexporter & " car Erreur de lecture du Fichier Format " & spathFileFormat)
                        End If
                    Catch ex As Exception
                        exceptionTrouve = True
                    End Try
                Else
                    If Trim(typeFormat) = "Longueur Fixe" Then
                        Try
                            Error_journal = File.AppendText(PathsFileRecuperer & "" & FichierRecup)
                            If AffichFormatintegration(spathFileFormat, typeFormat) = True Then
                                aRows = GetArrayFile(sPathFilexporter, aRows)
                                Datagridaffiche.Rows.Clear()
                                ProgressBar1.Value = ProgressBar1.Minimum
                                ProgressBar1.Maximum = UBound(aRows) + 1 - DecFormat
                                For i = DecFormat To UBound(aRows)
                                    If Datagridaffiche.Columns.Contains(IdentifiantPiece) = True Then
                                        iColPosition = CInt(Strings.Left(Strings.Right(Datagridaffiche.Columns(IdentifiantPiece).HeaderText, Strings.Len(Datagridaffiche.Columns(IdentifiantPiece).HeaderText) - InStr(Datagridaffiche.Columns(IdentifiantPiece).HeaderText, "[")), InStr(Strings.Right(Datagridaffiche.Columns(IdentifiantPiece).HeaderText, Strings.Len(Datagridaffiche.Columns(IdentifiantPiece).HeaderText) - InStr(Datagridaffiche.Columns(IdentifiantPiece).HeaderText, "[")), "]") - 1))
                                        iColGauchetxt = CInt(Strings.Right(Strings.Left(Datagridaffiche.Columns(IdentifiantPiece).HeaderText, InStr(Datagridaffiche.Columns(IdentifiantPiece).HeaderText, ")") - 1), Strings.Len(Strings.Left(Datagridaffiche.Columns(IdentifiantPiece).HeaderText, InStr(Datagridaffiche.Columns(IdentifiantPiece).HeaderText, ")") - 1)) - InStr(Strings.Left(Datagridaffiche.Columns(IdentifiantPiece).HeaderText, InStr(Datagridaffiche.Columns(IdentifiantPiece).HeaderText, ")") - 1), "(")))
                                        If iColPosition <> 0 Or iColGauchetxt <> 0 Then
                                            If Strings.Len(Trim(Strings.Mid(aRows(i), iColPosition, iColGauchetxt))) <= 8 Then
                                                If ListePiece.Contains(Formatage_Chaine(Trim(Strings.Mid(aRows(i), iColPosition, iColGauchetxt)))) = True Then
                                                    Error_journal.WriteLine(aRows(i))
                                                End If
                                            Else
                                                If ListePiece.Contains(Formatage_Chaine(Strings.Left(Trim(Strings.Mid(aRows(i), iColPosition, iColGauchetxt)), 8))) = True Then
                                                    Error_journal.WriteLine(aRows(i))
                                                End If
                                            End If
                                        Else
                                            If ListePiece.Contains(LireFichierFormat(spathFileFormat, IdentifiantPiece, typeFormat)) = True Then
                                                Error_journal.WriteLine(aRows(i))
                                            End If
                                        End If
                                    End If
                                    ProgressBar1.Value = ProgressBar1.Value + 1
                                Next i
                                Error_journal.Close()
                            Else
                                ErreurJrn.WriteLine("Impossible d'integrer le Fichier " & sPathFilexporter & " car Erreur de lecture du Fichier Format " & spathFileFormat)
                            End If
                        Catch ex As Exception
                            exceptionTrouve = True
                        End Try
                    End If
                End If
            End If
        End If
    End Sub
    Private Sub BT_Quitter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Quitter.Click
        Me.Close()
    End Sub
    Private Function Modification_Verification_Fichier(ByVal sPathFilexporter As String, ByVal spathFileFormat As String, ByRef Formatype As String, ByRef Base_Excel As String, ByRef sColumnsSepar As String, ByRef IdentifiantPiece As String, ByRef FormatdeDatefich As String) As Boolean
        Dim OleAdaptater As OleDbDataAdapter
        Dim OleAfficheDataset As DataSet
        Dim Oledatable As DataTable
        Dim m As Integer
        Dim jColD As Integer
        Dim iLine As Integer
        Dim aRows() As String
        Dim iColPosition, iColGauchetxt As Integer
        Dim i As Integer, aCols() As String
        Initialiser()
        iLine = 0
        aRows = Nothing

        If Trim(Formatype) = "Excel" Then
            Try
                If AffichFormatintegration(spathFileFormat, Formatype) = True Then
                    ProgressBar1.Value = ProgressBar1.Minimum
                    Datagridaffiche.Rows.Clear()
                    NbreTotal = DecFormat
                    OleAdaptater = New OleDbDataAdapter("select * from [" & Base_Excel & "$] ", OleExcelConnected)
                    OleAfficheDataset = New DataSet
                    OleAdaptater.Fill(OleAfficheDataset)
                    Oledatable = OleAfficheDataset.Tables(0)
                    If Oledatable.Rows.Count <> 0 Then
                        ProgresMax = Oledatable.Rows.Count - DecFormat
                        m = 0
                        infoListe = New List(Of Integer)
                        infoLigne = New List(Of Integer)
                        For i = DecFormat To Oledatable.Rows.Count - 1
                            Datagridaffiche.RowCount = iLine + 1 - m
                            For jColD = 0 To Datagridaffiche.ColumnCount - 1
                                iColPosition = CInt(Strings.Right(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1), Strings.Len(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1)) - InStr(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1), "(")))
                                If iColPosition <= Oledatable.Columns.Count Then
                                    If iColPosition <> 0 Then
                                        If Convert.IsDBNull(Oledatable.Rows(i).Item(iColPosition - 1)) = False Then
                                            Datagridaffiche.Item(jColD, iLine - m).Value = Trim(Oledatable.Rows(i).Item(iColPosition - 1))
                                        Else
                                            Datagridaffiche.Item(jColD, iLine - m).Value = ""
                                        End If
                                    Else
                                        Datagridaffiche.Item(jColD, iLine - m).Value = LireFichierFormat(spathFileFormat, Datagridaffiche.Columns(jColD).Name, Formatype)
                                    End If
                                Else
                                    Datagridaffiche.Item(jColD, iLine - m).Value = ""
                                End If
                            Next jColD
                            iLine = iLine + 1
                            If i Mod 10 = 0 Then
                                Me.Refresh()
                                Modification_Verification_Integrer_Ecriture(IdentifiantPiece, FormatdeDatefich, Trim(Formatype), LirePUDefaut(spathFileFormat, Trim(Formatype)), LireLigneArticle(spathFileFormat, Trim(Formatype)))
                                m = iLine
                            Else
                                If i = (Oledatable.Rows.Count - 1) Then
                                    Me.Refresh()
                                    Modification_Verification_Integrer_Ecriture(IdentifiantPiece, FormatdeDatefich, Trim(Formatype), LirePUDefaut(spathFileFormat, Trim(Formatype)), LireLigneArticle(spathFileFormat, Trim(Formatype)))
                                    m = iLine
                                End If
                            End If
                        Next i
                    End If
                Else
                    ErreurJrn.WriteLine("Impossible d' Integrer le Fichier " & sPathFilexporter & " car Erreur de lecture du Fichier Format " & spathFileFormat)
                End If
            Catch ex As Exception
                exceptionTrouve = True
            End Try
        Else
            If Trim(Formatype) = "Délimité" Or Trim(Formatype) = "Tabulation" Or Trim(Formatype) = "Pipe" Then
                Try
                    If AffichFormatintegration(spathFileFormat, Formatype) = True Then
                        aRows = GetArrayFile(sPathFilexporter, aRows)
                        NbreTotal = DecFormat
                        ProgressBar1.Value = ProgressBar1.Minimum
                        Datagridaffiche.Rows.Clear()
                        ProgresMax = UBound(aRows) + 1 - DecFormat
                        m = 0
                        infoListe = New List(Of Integer)
                        infoLigne = New List(Of Integer)
                        For i = DecFormat To UBound(aRows)
                            aCols = Split(aRows(i), sColumnsSepar)
                            Datagridaffiche.RowCount = iLine + 1 - m
                            For jColD = 0 To Datagridaffiche.ColumnCount - 1
                                iColPosition = CInt(Strings.Right(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1), Strings.Len(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1)) - InStr(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1), "(")))
                                If iColPosition <> 0 Then
                                    If iColPosition <= (UBound(aCols) + 1) Then
                                        Datagridaffiche.Item(jColD, iLine - m).Value = Trim(aCols(iColPosition - 1))
                                    Else
                                        Datagridaffiche.Item(jColD, iLine - m).Value = ""
                                    End If
                                Else
                                    Datagridaffiche.Item(jColD, iLine - m).Value = LireFichierFormat(spathFileFormat, Datagridaffiche.Columns(jColD).Name, Formatype)
                                End If
                            Next jColD
                            iLine = iLine + 1
                            If i Mod 10 = 0 Then
                                Me.Refresh()
                                Modification_Verification_Integrer_Ecriture(IdentifiantPiece, FormatdeDatefich, Trim(Formatype), LirePUDefaut(spathFileFormat, Trim(Formatype)), LireLigneArticle(spathFileFormat, Trim(Formatype)))
                                m = iLine
                            Else
                                If i = UBound(aRows) Then
                                    Me.Refresh()
                                    Modification_Verification_Integrer_Ecriture(IdentifiantPiece, FormatdeDatefich, Trim(Formatype), LirePUDefaut(spathFileFormat, Trim(Formatype)), LireLigneArticle(spathFileFormat, Trim(Formatype)))
                                    m = iLine
                                End If
                            End If
                        Next i
                    Else
                        ErreurJrn.WriteLine("Impossible d' Integrer le Fichier " & sPathFilexporter & " car Erreur de lecture du Fichier Format " & spathFileFormat)
                    End If
                Catch ex As Exception
                    exceptionTrouve = True
                End Try
            Else
                If Trim(Formatype) = "Longueur Fixe" Then
                    Try
                        If AffichFormatintegration(spathFileFormat, Formatype) = True Then
                            aRows = GetArrayFile(sPathFilexporter, aRows)
                            NbreTotal = DecFormat
                            ProgressBar1.Value = ProgressBar1.Minimum
                            Datagridaffiche.Rows.Clear()
                            ProgresMax = UBound(aRows) + 1 - DecFormat
                            m = 0
                            infoListe = New List(Of Integer)
                            infoLigne = New List(Of Integer)
                            For i = DecFormat To UBound(aRows)
                                Datagridaffiche.RowCount = iLine + 1 - m
                                For jColD = 0 To Datagridaffiche.ColumnCount - 1
                                    iColPosition = CInt(Strings.Left(Strings.Right(Datagridaffiche.Columns(jColD).HeaderText, Strings.Len(Datagridaffiche.Columns(jColD).HeaderText) - InStr(Datagridaffiche.Columns(jColD).HeaderText, "[")), InStr(Strings.Right(Datagridaffiche.Columns(jColD).HeaderText, Strings.Len(Datagridaffiche.Columns(jColD).HeaderText) - InStr(Datagridaffiche.Columns(jColD).HeaderText, "[")), "]") - 1))
                                    iColGauchetxt = CInt(Strings.Right(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1), Strings.Len(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1)) - InStr(Strings.Left(Datagridaffiche.Columns(jColD).HeaderText, InStr(Datagridaffiche.Columns(jColD).HeaderText, ")") - 1), "(")))
                                    If iColPosition <> 0 Or iColGauchetxt <> 0 Then
                                        Datagridaffiche.Item(jColD, iLine - m).Value = Trim(Strings.Mid(aRows(i), iColPosition, iColGauchetxt))
                                    Else
                                        Datagridaffiche.Item(jColD, iLine - m).Value = LireFichierFormat(spathFileFormat, Datagridaffiche.Columns(jColD).Name, Formatype)
                                    End If
                                Next jColD
                                iLine = iLine + 1
                                If i Mod 10 = 0 Then
                                    Me.Refresh()
                                    Modification_Verification_Integrer_Ecriture(IdentifiantPiece, FormatdeDatefich, Trim(Formatype), LirePUDefaut(spathFileFormat, Trim(Formatype)), LireLigneArticle(spathFileFormat, Trim(Formatype)))
                                    m = iLine
                                Else
                                    If i = UBound(aRows) Then
                                        Me.Refresh()
                                        Modification_Verification_Integrer_Ecriture(IdentifiantPiece, FormatdeDatefich, Trim(Formatype), LirePUDefaut(spathFileFormat, Trim(Formatype)), LireLigneArticle(spathFileFormat, Trim(Formatype)))
                                        m = iLine
                                    End If
                                End If
                            Next i
                        Else
                            ErreurJrn.WriteLine("Impossible d' Integrer le Fichier " & sPathFilexporter & " car Erreur de lecture du Fichier Format " & spathFileFormat)
                        End If
                    Catch ex As Exception
                        exceptionTrouve = True
                    End Try
                End If
            End If
        End If
        Modification_Verification_Fichier = ExisteLecture
    End Function
    Private Sub Modification_Verification_Integrer_Ecriture(ByRef IdentifiantPiece As String, ByRef FormatfichierDate As String, ByRef FormatIntegrer As Object, ByRef Punitaire As String, ByRef IdentifiantArticle As String)
        Me.Cursor = Cursors.WaitCursor
        BT_integrer.Enabled = False
        If Datagridaffiche.RowCount >= 0 Then
            ProgressBar1.Maximum = ProgresMax
            Try
                For numLigne = 0 To Datagridaffiche.RowCount - 1
                    vidage()
                    NbreTotal = NbreTotal + 1
                    Label5.Refresh()
                    Label5.Text = "Vérification des Integrations!"
                    For numColonne = 0 To Datagridaffiche.ColumnCount - 1
                        'Entête Document
                        If Datagridaffiche.Columns.Contains(IdentifiantPiece) = True Then
                            If Strings.Len(Trim(Datagridaffiche.Rows(numLigne).Cells(IdentifiantPiece).Value)) <= 8 Then
                                EntetePieceInterne = Formatage_Chaine(Trim(Datagridaffiche.Rows(numLigne).Cells(IdentifiantPiece).Value))
                            Else
                                EntetePieceInterne = Formatage_Chaine(Strings.Left(Trim(Datagridaffiche.Rows(numLigne).Cells(IdentifiantPiece).Value), 8))
                                ErreurJrn.WriteLine("N°Pièce :" & Trim(Datagridaffiche.Rows(numLigne).Cells(IdentifiantPiece).Value) & "  a été tronquée")
                            End If
                        End If

                        If Datagridaffiche.Columns.Contains(IdentifiantArticle) = True Then
                            PieceArticle = Trim(Datagridaffiche.Rows(numLigne).Cells(IdentifiantArticle).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "EnteteTyPeDocument" Then
                            EnteteTyPeDocument = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "EnteteCodeAffaire" Then
                            EnteteCodeAffaire = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "EntetePlanAnalytique" Then
                            EntetePlanAnalytique = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "EnteteReference" Then
                            If Strings.Len(Trim(Datagridaffiche.Item(numColonne, numLigne).Value)) <= 17 Then
                                EnteteReference = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                            Else
                                EnteteReference = Strings.Left(Trim(Datagridaffiche.Item(numColonne, numLigne).Value), 17)
                                ErreurJrn.WriteLine("La Référence en Entête :" & EnteteReference & " de la Pièce : " & Trim(EntetePieceInterne) & " a été tronqué")
                            End If
                        End If
                        'Ligne Document
                        If Datagridaffiche.Columns(numColonne).Name = "LigneDesignationArticle" Then
                            If Strings.Len(Trim(Datagridaffiche.Item(numColonne, numLigne).Value)) <= 69 Then
                                LigneDesignationArticle = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                            Else
                                LigneDesignationArticle = Strings.Left(Trim(Datagridaffiche.Item(numColonne, numLigne).Value), 69)
                                ErreurJrn.WriteLine("La Désignation Article :" & LigneDesignationArticle & " de la Pièce : " & Trim(EntetePieceInterne) & " a été tronqué")
                            End If
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "EnteteIntituleDepotOrigine" Then
                            EnteteIntituleDepotOrigine = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "EnteteIntituleDepotDestination" Then
                            EnteteIntituleDepotDestination = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If

                        If Datagridaffiche.Columns(numColonne).Name = "IDDepotEnteteOrigine" Then
                            IDDepotEnteteOrigine = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "IDDepotEnteteDestination" Then
                            IDDepotEnteteDestination = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "IDDepotLigne" Then
                            IDDepotLigne = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "LigneIntituleDepot" Then
                            LigneIntituleDepot = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "LignePoidsBrut" Then
                            LignePoidsBrut = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "LignePoidsNet" Then
                            LignePoidsNet = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "LignePrixUnitaire" Then
                            LignePrixUnitaire = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                        End If

                        If Datagridaffiche.Columns.Contains("LigneQuantite") = True Then
                            LigneQuantite = Trim(Datagridaffiche.Rows(numLigne).Cells("LigneQuantite").Value)
                        End If

                        If Datagridaffiche.Columns(numColonne).Name = "LigneReference" Then
                            If Strings.Len(Trim(Datagridaffiche.Item(numColonne, numLigne).Value)) <= 17 Then
                                LigneReference = Trim(Datagridaffiche.Item(numColonne, numLigne).Value)
                            Else
                                LigneReference = Strings.Left(Trim(Datagridaffiche.Item(numColonne, numLigne).Value), 17)
                                ErreurJrn.WriteLine("La Référence en Ligne :" & LigneReference & " de la Pièce : " & Trim(EntetePieceInterne) & " a été tronqué")
                            End If
                        End If
                        If Datagridaffiche.Columns(numColonne).Name = "LigneCodeArticle" Then
                            LigneCodeArticle = Formatage_Article(Trim(Datagridaffiche.Item(numColonne, numLigne).Value))
                        End If

                        'RECHERCHE DE L'INTITULE DE L'INFO LIBRE
                        If Trim(FormatIntegrer) = "Longueur Fixe" Then
                            Dim InfoTableau() As String = Split(Trim(Strings.Left(Strings.Right(Datagridaffiche.Columns(numColonne).HeaderText, Strings.Len(Datagridaffiche.Columns(numColonne).HeaderText) - InStr(Datagridaffiche.Columns(numColonne).HeaderText, "{")), Strings.Len(Strings.Right(Datagridaffiche.Columns(numColonne).HeaderText, Strings.Len(Datagridaffiche.Columns(numColonne).HeaderText) - InStr(Datagridaffiche.Columns(numColonne).HeaderText, "{"))) - 1)), "-")
                            If Trim(InfoTableau(0)) = "oui" Then
                                If Trim(InfoTableau(1)) = "F_DOCENTETE" Then
                                    infoListe.Add(numColonne)
                                End If
                                If Trim(InfoTableau(1)) = "F_DOCLIGNE" Then
                                    infoLigne.Add(numColonne)
                                End If
                            End If
                        Else
                            Dim InfoTableau() As String = Split(Trim(Strings.Left(Strings.Right(Datagridaffiche.Columns(numColonne).HeaderText, Strings.Len(Datagridaffiche.Columns(numColonne).HeaderText) - InStr(Datagridaffiche.Columns(numColonne).HeaderText, "[")), Strings.Len(Strings.Right(Datagridaffiche.Columns(numColonne).HeaderText, Strings.Len(Datagridaffiche.Columns(numColonne).HeaderText) - InStr(Datagridaffiche.Columns(numColonne).HeaderText, "["))) - 1)), "-")
                            If Trim(InfoTableau(0)) = "oui" Then
                                If Trim(InfoTableau(1)) = "F_DOCENTETE" Then
                                    infoListe.Add(numColonne)
                                End If
                                If Trim(InfoTableau(1)) = "F_DOCLIGNE" Then
                                    infoLigne.Add(numColonne)
                                End If
                            End If
                        End If
                    Next numColonne
                    'Creation Effective du Document Commercial
                   
                    If Trim(EnteteTyPeDocument) = "23" Then 'Transfert
                        If BaseCial.FactoryDocumentStock.ExistPiece(DocumentType.DocumentTypeStockVirement, Trim(EntetePieceInterne)) = True Then
                            Document = BaseCial.FactoryDocumentStock.ReadPiece(DocumentType.DocumentTypeStockVirement, Trim(EntetePieceInterne))
                            Modification_Verification_Parametrage(EnteteIntituleDepotOrigine, EntetePieceInterne, EnteteTyPeDocument, Document, infoListe, FormatfichierDate, PieceArticle, Punitaire, IdentifiantArticle)
                        Else
                            ExisteLecture = False
                            exceptionTrouve = True
                            ErreurJrn.WriteLine("Le Mouvement de Transfert N°: " & Trim(EntetePieceInterne) & " n'existe pas dans la base commerciale")
                        End If
                    Else
                        ExisteLecture = False
                        exceptionTrouve = True
                        ErreurJrn.WriteLine("Le type de document ne correspond à aucune de ces valeurs (23)")
                    End If
                    NbreLigne = NbreLigne + 1
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    Label8.Text = NbreLigne & "/" & ProgresMax
                    Label8.Refresh()
                Next numLigne
            Catch ex As Exception
                ExisteLecture = False
                exceptionTrouve = True
                ErreurJrn.WriteLine("Une erreur s'est produit au moment de la lecture du fichier  : " & Trim(EnteteIntituleDepotOrigine))
            End Try
        End If
        Datagridaffiche.Rows.Clear()
        Me.Cursor = Cursors.Default
        BT_integrer.Enabled = True
    End Sub
    Private Sub Modification_Verification_Parametrage(ByRef EnteteIntituleDepotOrigine As String, ByRef EntetePieceInterne As String, ByRef EnteteTyPeDocument As String, ByRef Document As IBODocumentStock3, ByRef infoListe As List(Of Integer), ByRef FormatDatefichier As String, ByRef PieceArticle As String, ByRef Punitaire As String, ByRef IdentifiantArticle As String)
        Dim fournisseurAdap As OleDbDataAdapter
        Dim fournisseurDs As DataSet
        Dim fournisseurTab As DataTable
        Dim statistAdap As OleDbDataAdapter
        Dim statistDs As DataSet
        Dim statistTab As DataTable
        Dim OleAdaptaterDelete As OleDbDataAdapter
        Dim OleDeleteDataset As DataSet
        Dim OledatableDelete As DataTable
        Dim OleRecherAdapter As OleDbDataAdapter
        Dim OleRecherDataset As DataSet
        Dim OleRechDatable As DataTable
        Dim DossierAdap As OleDbDataAdapter
        Dim DossierDs As DataSet
        Dim DossierTab As DataTable
        Dim FormatQte As Integer = 0
        Dim FormatMnt As Integer = 0
        DossierAdap = New OleDbDataAdapter("select * from P_DOSSIERCIAL", OleSocieteConnect)
        DossierDs = New DataSet
        DossierAdap.Fill(DossierDs)
        DossierTab = DossierDs.Tables(0)
        If DossierTab.Rows.Count <> 0 Then
            If InStr(DossierTab.Rows(0).Item("D_FormatQte"), ",") <> 0 Then
                FormatQte = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatQte")), Len(Trim(DossierTab.Rows(0).Item("D_FormatQte"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatQte")), ",")))
            End If
            If InStr(DossierTab.Rows(0).Item("D_FormatPrix"), ",") <> 0 Then
                FormatMnt = Len(Strings.Right(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), Len(Trim(DossierTab.Rows(0).Item("D_FormatPrix"))) - InStr(Trim(DossierTab.Rows(0).Item("D_FormatPrix")), ",")))
            End If
        End If
        If Datagridaffiche.Columns.Contains("EnteteTyPeDocument") = True Then
            If Trim(EnteteTyPeDocument) <> "" Then
               
                If Trim(EnteteTyPeDocument) = "23" Then
                    If Datagridaffiche.Columns.Contains("LigneQuantite") = True Then
                        If Trim(LigneQuantite) <> "" Then
                            If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                                If CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)) < 0 Then
                                    ExisteLecture = False
                                    ErreurJrn.WriteLine("< La Quantité : " & Trim(LigneQuantite) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " ne doit pas être négative >")
                                End If
                            End If
                        End If
                    End If
                Else
                    ErreurJrn.WriteLine("Le statut du document " & EnteteTyPeDocument & " dois être égal à 23:Transfert : " & EntetePieceInterne & " le statut par défaut va être utilisé")
                End If
            End If
        End If

        If Datagridaffiche.Columns.Contains("EntetePlanAnalytique") = True Then
            If Datagridaffiche.Columns.Contains("EnteteCodeAffaire") = True Then
                If Trim(EntetePlanAnalytique) <> "" Then
                    If Trim(EnteteCodeAffaire) <> "" Then
                        If BaseCpta.FactoryAnalytique.ExistIntitule(Trim(EntetePlanAnalytique)) = True Then
                            PlanAna = BaseCpta.FactoryAnalytique.ReadIntitule(Trim(EntetePlanAnalytique))
                            If BaseCpta.FactoryCompteA.ExistNumero(PlanAna, Trim(EnteteCodeAffaire)) = True Then
                            Else
                                statistAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Section Analytique>' and Valeurlue ='" & Join(Split(Trim(EnteteCodeAffaire), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                statistDs = New DataSet
                                statistAdap.Fill(statistDs)
                                statistTab = statistDs.Tables(0)
                                If statistTab.Rows.Count <> 0 Then
                                    If BaseCpta.FactoryCompteA.ExistNumero(PlanAna, Trim(statistTab.Rows(0).Item("Correspond"))) = True Then

                                    Else
                                        ExisteLecture = False
                                        ErreurJrn.WriteLine("< Le Code de Section analytique : " & Trim(statistTab.Rows(0).Item("Correspond")) & " - N°Pièce : " & Trim(EntetePieceInterne) & " n'existe pas dans Sage >")
                                    End If
                                Else
                                    ExisteLecture = False
                                    ErreurJrn.WriteLine("< Le Code de Section analytique : " & Trim(EnteteCodeAffaire) & " - N°Pièce : " & Trim(EntetePieceInterne) & " n'existe pas dans la table de paramétrage et dans Sage >")
                                End If
                            End If
                        Else
                            fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Plan Analytique>' and Valeurlue ='" & Join(Split(Trim(EntetePlanAnalytique), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                            fournisseurDs = New DataSet
                            fournisseurAdap.Fill(fournisseurDs)
                            fournisseurTab = fournisseurDs.Tables(0)
                            If fournisseurTab.Rows.Count > 0 Then
                                If BaseCpta.FactoryAnalytique.ExistIntitule(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                                    PlanAna = BaseCpta.FactoryAnalytique.ReadIntitule(Trim(fournisseurTab.Rows(0).Item("Correspond")))
                                    If BaseCpta.FactoryCompteA.ExistNumero(PlanAna, Trim(EnteteCodeAffaire)) = True Then
                                    Else
                                        statistAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<Section Analytique>' and Valeurlue ='" & Join(Split(Trim(EnteteCodeAffaire), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                        statistDs = New DataSet
                                        statistAdap.Fill(statistDs)
                                        statistTab = statistDs.Tables(0)
                                        If statistTab.Rows.Count <> 0 Then
                                            If BaseCpta.FactoryCompteA.ExistNumero(PlanAna, Trim(statistTab.Rows(0).Item("Correspond"))) = True Then

                                            Else
                                                ExisteLecture = False
                                                ErreurJrn.WriteLine("< Le Code de Section analytique : " & Trim(statistTab.Rows(0).Item("Correspond")) & " - N°Pièce : " & Trim(EntetePieceInterne) & " n'existe pas dans Sage >")
                                            End If
                                        Else
                                            ExisteLecture = False
                                            ErreurJrn.WriteLine("< Le Code de Section analytique : " & Trim(EnteteCodeAffaire) & " - N°Pièce : " & Trim(EntetePieceInterne) & " n'existe pas dans la table de paramétrage et dans Sage >")
                                        End If
                                    End If
                                Else
                                    ExisteLecture = False
                                    ErreurJrn.WriteLine("< Le Code du plan analytique : " & Trim(fournisseurTab.Rows(0).Item("Correspond")) & " - N°Pièce : " & Trim(EntetePieceInterne) & " n'existe pas dans Sage >")
                                End If
                            Else
                                ExisteLecture = False
                                ErreurJrn.WriteLine("< Le Code du plan analytique : " & Trim(EntetePlanAnalytique) & " - N°Pièce : " & Trim(EntetePieceInterne) & " n'existe pas dans la table de paramétrage et dans Sage >")
                            End If
                        End If
                    End If
                End If
            End If
        End If
        If Datagridaffiche.Columns.Contains("LignePoidsNet") = True Then
            If Trim(LignePoidsNet) <> "" Then
                If EstNumeric(Trim(LignePoidsNet), DecimalNomb, DecimalMone) = True Then
                Else
                    ExisteLecture = False
                    ErreurJrn.WriteLine("< Le poids Net : " & Trim(LignePoidsNet) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'est pas numérique >")
                End If
            End If
        End If
        If Datagridaffiche.Columns.Contains("LignePoidsBrut") = True Then
            If Trim(LignePoidsBrut) <> "" Then
                If EstNumeric(Trim(LignePoidsBrut), DecimalNomb, DecimalMone) = True Then
                Else
                    ExisteLecture = False
                    ErreurJrn.WriteLine("< Le poids brut : " & Trim(LignePoidsBrut) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'est pas numérique >")
                End If
            End If
        End If

        If Datagridaffiche.Columns.Contains("LignePrixUnitaire") = True Then
            If Trim(LignePrixUnitaire) <> "" Then
                If EstNumeric(Trim(LignePrixUnitaire), DecimalNomb, DecimalMone) = True Then
                Else
                    ExisteLecture = False
                    ErreurJrn.WriteLine("< Le prix unitaire : " & Trim(LignePrixUnitaire) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'est pas numérique >")
                End If
            End If
        End If
        If Datagridaffiche.Columns.Contains("LigneQuantite") = True Then
            If Trim(LigneQuantite) <> "" Then
                If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then

                Else
                    ExisteLecture = False
                    ErreurJrn.WriteLine("< La Quantité : " & Trim(LigneQuantite) & " - N°Pièce : " & Trim(EntetePieceInterne) & " n'est pas numérique >")
                End If
            End If
        End If
        If Datagridaffiche.Columns.Contains("LigneCodeArticle") = True Then
            If Trim(LigneCodeArticle) <> "" Then
                If BaseCial.FactoryArticle.ExistReference(Trim(LigneCodeArticle)) = True Then
                    If Trim(LigneQuantite) <> "" Then
                        statistAdap = New OleDbDataAdapter("select * from F_DOCLIGNE where DO_Piece ='" & Join(Split(Trim(EntetePieceInterne), "'"), "''") & "' and AR_Ref ='" & Join(Split(Trim(LigneCodeArticle), "'"), "''") & "'", OleSocieteConnect)
                        statistDs = New DataSet
                        statistAdap.Fill(statistDs)
                        statistTab = statistDs.Tables(0)
                        If statistTab.Rows.Count <= 1 Then

                        Else
                            ExisteLecture = False
                            ErreurJrn.WriteLine("< La Référence Article : " & Trim(LigneCodeArticle) & " est present plusieurs fois dans la pièce Commerciale - N°Pièce : " & Trim(EntetePieceInterne) & " >")
                        End If
                    Else
                        ExisteLecture = False
                        ErreurJrn.WriteLine("< La quantité pour La Référence Article : " & Trim(LigneCodeArticle) & " existant en Gestion Commerciale - N°Pièce : " & Trim(EntetePieceInterne) & " doit être obligatoire >")
                    End If
                Else
                    fournisseurAdap = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='<ArticleStock>' and Valeurlue ='" & Join(Split(Trim(LigneCodeArticle), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                    fournisseurDs = New DataSet
                    fournisseurAdap.Fill(fournisseurDs)
                    fournisseurTab = fournisseurDs.Tables(0)
                    If fournisseurTab.Rows.Count <> 0 Then
                        If BaseCial.FactoryArticle.ExistReference(Trim(fournisseurTab.Rows(0).Item("Correspond"))) = True Then
                            If Trim(LigneQuantite) <> "" Then
                                statistAdap = New OleDbDataAdapter("select * from F_DOCLIGNE where DO_Piece ='" & Join(Split(Trim(EntetePieceInterne), "'"), "''") & "' and AR_Ref ='" & Join(Split(Trim(fournisseurTab.Rows(0).Item("Correspond")), "'"), "''") & "'", OleSocieteConnect)
                                statistDs = New DataSet
                                statistAdap.Fill(statistDs)
                                statistTab = statistDs.Tables(0)
                                If statistTab.Rows.Count <= 1 Then

                                Else
                                    ExisteLecture = False
                                    ErreurJrn.WriteLine("< La Référence Article : " & Trim(fournisseurTab.Rows(0).Item("Correspond")) & " est present plusieurs fois dans la pièce Commerciale - N°Pièce : " & Trim(EntetePieceInterne) & " >")
                                End If
                            Else
                                ExisteLecture = False
                                ErreurJrn.WriteLine("< La quantité pour La Référence Article : " & Trim(fournisseurTab.Rows(0).Item("Correspond")) & " existant en Gestion Commerciale - N°Pièce : " & Trim(EntetePieceInterne) & " doit être obligatoire >")
                            End If
                        Else
                            ExisteLecture = False
                            ErreurJrn.WriteLine("< La Référence Article : " & Trim(fournisseurTab.Rows(0).Item("Correspond")) & " - N°Pièce : " & Trim(EntetePieceInterne) & " n'existe pas dans Sage >")
                        End If
                    Else
                        ExisteLecture = False
                        ErreurJrn.WriteLine("< La Référence Article : " & Trim(LigneCodeArticle) & " - Dépôt : " & Join(Split(Trim(EnteteIntituleDepotOrigine), "'"), "''") & " - N°Pièce : " & Trim(EntetePieceInterne) & " n'existe pas dans la table de paramétrage et dans Sage >")
                    End If
                End If
            Else
                If Datagridaffiche.Columns.Contains("LigneQuantite") = True Then
                    If Trim(LigneQuantite) <> "" Then
                        ExisteLecture = False
                        ErreurJrn.WriteLine("< La quantité :" & Trim(LigneQuantite) & " ne doit pas être renseignée pour La Référence Article : " & Trim(LigneCodeArticle) & "vide - N°Pièce : " & Trim(EntetePieceInterne) & " >")
                    End If
                End If
                If Datagridaffiche.Columns.Contains("LignePrixUnitaire") = True Then
                    If Trim(LignePrixUnitaire) <> "" Then
                        ExisteLecture = False
                        ErreurJrn.WriteLine("< Le prix unitaire :" & Trim(LignePrixUnitaire) & " ne doit pas être renseignée pour La Référence Article : " & Trim(LigneCodeArticle) & "vide - N°Pièce : " & Trim(EntetePieceInterne) & " >")
                    End If
                End If
            End If
        Else
            If Datagridaffiche.Columns.Contains("LigneQuantite") = True Then
                If Trim(LigneQuantite) <> "" Then
                    ExisteLecture = False
                    ErreurJrn.WriteLine("< La quantité :" & Trim(LigneQuantite) & " ne doit pas être renseignée pour La Référence Article : " & Trim(LigneCodeArticle) & "vide - N°Pièce : " & Trim(EntetePieceInterne) & " >")
                End If
            End If
            If Datagridaffiche.Columns.Contains("LignePrixUnitaire") = True Then
                If Trim(LignePrixUnitaire) <> "" Then
                    ExisteLecture = False
                    ErreurJrn.WriteLine("< Le prix unitaire :" & Trim(LignePrixUnitaire) & " ne doit pas être renseignée pour La Référence Article : " & Trim(LigneCodeArticle) & "vide - N°Pièce : " & Trim(EntetePieceInterne) & " >")
                End If
            End If
        End If
        'Traitement des Infos Libres
        Try
            If infoLigne.Count > 0 Then
                While infoLigne.Count <> 0
                    OleAdaptaterDelete = New OleDbDataAdapter("select * From WIT_COL where Libelle='" & Trim(Datagridaffiche.Columns(infoLigne.Item(0)).Name) & "' And InfoLigne=True", OleConnenection)
                    OleDeleteDataset = New DataSet
                    OleAdaptaterDelete.Fill(OleDeleteDataset)
                    OledatableDelete = OleDeleteDataset.Tables(0)
                    If OledatableDelete.Rows.Count <> 0 Then
                        'L'info Libre Parametrée par l'utilisateur existe dans Sage
                        If IsNothing(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) = False Then
                            If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                statistAdap = New OleDbDataAdapter("select * from cbSysLibre where CB_File='F_DOCLIGNE' and CB_Name ='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "'", OleSocieteConnect)
                                statistDs = New DataSet
                                statistAdap.Fill(statistDs)
                                statistTab = statistDs.Tables(0)
                                If statistTab.Rows.Count <> 0 Then
                                    'Texte
                                    If statistTab.Rows(0).Item("CB_Type") = 9 Then
                                        OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                        OleRecherDataset = New DataSet
                                        OleRecherAdapter.Fill(OleRecherDataset)
                                        OleRechDatable = OleRecherDataset.Tables(0)
                                        If OleRechDatable.Rows.Count <> 0 Then
                                            If statistTab.Rows(0).Item("CB_Len") >= Strings.Len(Trim(OleRechDatable.Rows(0).Item("Correspond"))) Then

                                            Else
                                                ExisteLecture = False
                                                ErreurJrn.WriteLine("La Longueur de l'info libre : " & Datagridaffiche.Columns(infoLigne.Item(0)).Name & " est inférieur au valeur adressée: " & Trim(OleRechDatable.Rows(0).Item("Correspond")) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
                                            End If
                                        Else
                                            If statistTab.Rows(0).Item("CB_Len") >= Strings.Len(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value)) Then

                                            Else
                                                ExisteLecture = False
                                                ErreurJrn.WriteLine("La Longueur de l'info libre : " & Datagridaffiche.Columns(infoLigne.Item(0)).Name & " est inférieur au valeur adressée: " & Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
                                            End If
                                        End If
                                    End If
                                    'Table
                                    If statistTab.Rows(0).Item("CB_Type") = 22 Then
                                        OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                        OleRecherDataset = New DataSet
                                        OleRecherAdapter.Fill(OleRecherDataset)
                                        OleRechDatable = OleRecherDataset.Tables(0)
                                        If OleRechDatable.Rows.Count <> 0 Then
                                            If statistTab.Rows(0).Item("CB_Len") > Strings.Len(Trim(OleRechDatable.Rows(0).Item("Correspond"))) Then

                                            Else
                                                ExisteLecture = False
                                                ErreurJrn.WriteLine("La Longueur (21) de l'info libre de type table : " & Datagridaffiche.Columns(infoLigne.Item(0)).Name & " est inférieur au valeur adressée: " & Trim(OleRechDatable.Rows(0).Item("Correspond")) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
                                            End If
                                        Else
                                            If statistTab.Rows(0).Item("CB_Len") > Strings.Len(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value)) Then

                                            Else
                                                ExisteLecture = False
                                                ErreurJrn.WriteLine("La Longueur (21) de l'info libre de type table : " & Datagridaffiche.Columns(infoLigne.Item(0)).Name & " est inférieur au valeur adressée: " & Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
                                            End If
                                        End If
                                    End If
                                    'Montant
                                    If statistTab.Rows(0).Item("CB_Type") = 20 Then
                                        If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                            OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                            OleRecherDataset = New DataSet
                                            OleRecherAdapter.Fill(OleRecherDataset)
                                            OleRechDatable = OleRecherDataset.Tables(0)
                                            If OleRechDatable.Rows.Count <> 0 Then
                                                If EstNumeric(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone) = True Then

                                                Else
                                                    ExisteLecture = False
                                                    ErreurJrn.WriteLine("La valeur entrée : " & Trim(OleRechDatable.Rows(0).Item("Correspond")) & "doit être de type numérique pour l'info libre : " & Datagridaffiche.Columns(infoLigne.Item(0)).Name & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
                                                End If
                                            Else
                                                If EstNumeric(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), DecimalNomb, DecimalMone) = True Then

                                                Else
                                                    ExisteLecture = False
                                                    ErreurJrn.WriteLine("La valeur entrée : " & Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) & "doit être de type numérique pour l'info libre : " & Datagridaffiche.Columns(infoLigne.Item(0)).Name & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
                                                End If
                                            End If
                                        End If
                                    End If
                                    'Valeur
                                    If statistTab.Rows(0).Item("CB_Type") = 7 Then
                                        If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                            OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                            OleRecherDataset = New DataSet
                                            OleRecherAdapter.Fill(OleRecherDataset)
                                            OleRechDatable = OleRecherDataset.Tables(0)
                                            If OleRechDatable.Rows.Count <> 0 Then
                                                If EstNumeric(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone) = True Then

                                                Else
                                                    ExisteLecture = False
                                                    ErreurJrn.WriteLine("La valeur entrée : " & Trim(OleRechDatable.Rows(0).Item("Correspond")) & "doit être de type numérique pour l'info libre : " & Datagridaffiche.Columns(infoLigne.Item(0)).Name & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
                                                End If
                                            Else
                                                If EstNumeric(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), DecimalNomb, DecimalMone) = True Then

                                                Else
                                                    ExisteLecture = False
                                                    ErreurJrn.WriteLine("La valeur entrée : " & Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) & "doit être de type numérique pour l'info libre : " & Datagridaffiche.Columns(infoLigne.Item(0)).Name & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
                                                End If
                                            End If
                                        End If
                                    End If

                                    'Date Court
                                    If statistTab.Rows(0).Item("CB_Type") = 3 Then
                                        If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                            OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                            OleRecherDataset = New DataSet
                                            OleRecherAdapter.Fill(OleRecherDataset)
                                            OleRechDatable = OleRecherDataset.Tables(0)
                                            If OleRechDatable.Rows.Count <> 0 Then
                                                If Verificatdate(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then

                                                Else
                                                    ExisteLecture = False
                                                End If
                                            Else
                                                If Verificatdate(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then
                                                Else
                                                    ExisteLecture = False
                                                End If
                                            End If
                                        End If
                                    End If
                                    'Date Longue
                                    If statistTab.Rows(0).Item("CB_Type") = 14 Then
                                        If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                            OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Ligne=True", OleConnenection)
                                            OleRecherDataset = New DataSet
                                            OleRecherAdapter.Fill(OleRecherDataset)
                                            OleRechDatable = OleRecherDataset.Tables(0)
                                            If OleRechDatable.Rows.Count <> 0 Then
                                                If Verificatdate(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then

                                                Else
                                                    ExisteLecture = False
                                                End If
                                            Else
                                                If Verificatdate(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then
                                                Else
                                                    ExisteLecture = False
                                                End If
                                            End If
                                        End If
                                    End If
                                Else
                                    ErreurJrn.WriteLine(" Impossible de traiter l'information libre en Ligne de Document :" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "  Il est inexistant dans la table CbsysLibre de Sage")
                                    ExisteLecture = False
                                End If
                            End If
                        Else
                            'nothing
                        End If
                    Else
                        ErreurJrn.WriteLine(" Impossible de traiter l'information libre en Ligne de Document :" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "  Il est inexistant dans la table de Paramétrage")
                        ExisteLecture = False
                    End If
                    'L'info Libre Parametrée par l'utilisateur n'existe pas dans Sage
                    infoLigne.RemoveAt(0)
                End While
            End If
        Catch ex As Exception
            exceptionTrouve = True
            ExisteLecture = False
            ErreurJrn.WriteLine(" Erreur de Création de L'information Libre Ligne Document : " & ex.Message & ", vérifiez la longueur de la chaine avec celle paramétrée/ ou la cohérence des données")
        End Try
        'Traitement des Infos Libres
        Try
            If infoListe.Count > 0 Then
                While infoListe.Count <> 0
                    OleAdaptaterDelete = New OleDbDataAdapter("select * From WIT_COL where Libelle='" & Trim(Datagridaffiche.Columns(infoListe.Item(0)).Name) & "' And Libre=True", OleConnenection)
                    OleDeleteDataset = New DataSet
                    OleAdaptaterDelete.Fill(OleDeleteDataset)
                    OledatableDelete = OleDeleteDataset.Tables(0)
                    If OledatableDelete.Rows.Count <> 0 Then
                        'L'info Libre Parametrée par l'utilisateur existe dans Sage
                        If IsNothing(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) = False Then
                            If Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) <> "" Then
                                statistAdap = New OleDbDataAdapter("select * from cbSysLibre where CB_File='F_DOCENTETE' and CB_Name ='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "'", OleSocieteConnect)
                                statistDs = New DataSet
                                statistAdap.Fill(statistDs)
                                statistTab = statistDs.Tables(0)
                                If statistTab.Rows.Count <> 0 Then
                                    'Texte
                                    If statistTab.Rows(0).Item("CB_Type") = 9 Then
                                        OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                        OleRecherDataset = New DataSet
                                        OleRecherAdapter.Fill(OleRecherDataset)
                                        OleRechDatable = OleRecherDataset.Tables(0)
                                        If OleRechDatable.Rows.Count <> 0 Then
                                            If statistTab.Rows(0).Item("CB_Len") >= Strings.Len(Trim(OleRechDatable.Rows(0).Item("Correspond"))) Then

                                            Else
                                                ExisteLecture = False
                                                ErreurJrn.WriteLine("La Longueur de l'info libre : " & Datagridaffiche.Columns(infoListe.Item(0)).Name & " est inférieur au valeur adressée: " & Trim(OleRechDatable.Rows(0).Item("Correspond")) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
                                            End If
                                        Else
                                            If statistTab.Rows(0).Item("CB_Len") >= Strings.Len(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value)) Then

                                            Else
                                                ExisteLecture = False
                                                ErreurJrn.WriteLine("La Longueur de l'info libre : " & Datagridaffiche.Columns(infoListe.Item(0)).Name & " est inférieur au valeur adressée: " & Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
                                            End If
                                        End If
                                    End If
                                    'Table
                                    If statistTab.Rows(0).Item("CB_Type") = 22 Then
                                        OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                        OleRecherDataset = New DataSet
                                        OleRecherAdapter.Fill(OleRecherDataset)
                                        OleRechDatable = OleRecherDataset.Tables(0)
                                        If OleRechDatable.Rows.Count <> 0 Then
                                            If statistTab.Rows(0).Item("CB_Len") > Strings.Len(Trim(OleRechDatable.Rows(0).Item("Correspond"))) Then

                                            Else
                                                ExisteLecture = False
                                                ErreurJrn.WriteLine("La Longueur (21) de l'info libre de type table : " & Datagridaffiche.Columns(infoListe.Item(0)).Name & " est inférieur au valeur adressée: " & Trim(OleRechDatable.Rows(0).Item("Correspond")) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
                                            End If
                                        Else
                                            If statistTab.Rows(0).Item("CB_Len") > Strings.Len(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value)) Then

                                            Else
                                                ExisteLecture = False
                                                ErreurJrn.WriteLine("La Longueur (21) de l'info libre de type table : " & Datagridaffiche.Columns(infoListe.Item(0)).Name & " est inférieur au valeur adressée: " & Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
                                            End If
                                        End If
                                    End If
                                    'Montant
                                    If statistTab.Rows(0).Item("CB_Type") = 20 Then
                                        If Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) <> "" Then
                                            OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                            OleRecherDataset = New DataSet
                                            OleRecherAdapter.Fill(OleRecherDataset)
                                            OleRechDatable = OleRecherDataset.Tables(0)
                                            If OleRechDatable.Rows.Count <> 0 Then
                                                If EstNumeric(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone) = True Then

                                                Else
                                                    ExisteLecture = False
                                                    ErreurJrn.WriteLine("La valeur entrée : " & Trim(OleRechDatable.Rows(0).Item("Correspond")) & "doit être de type numérique pour l'info libre : " & Datagridaffiche.Columns(infoListe.Item(0)).Name & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
                                                End If
                                            Else
                                                If EstNumeric(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), DecimalNomb, DecimalMone) = True Then

                                                Else
                                                    ExisteLecture = False
                                                    ErreurJrn.WriteLine("La valeur entrée : " & Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) & "doit être de type numérique pour l'info libre : " & Datagridaffiche.Columns(infoListe.Item(0)).Name & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
                                                End If
                                            End If
                                        End If
                                    End If
                                    'Valeur
                                    If statistTab.Rows(0).Item("CB_Type") = 7 Then
                                        If Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) <> "" Then
                                            OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                            OleRecherDataset = New DataSet
                                            OleRecherAdapter.Fill(OleRecherDataset)
                                            OleRechDatable = OleRecherDataset.Tables(0)
                                            If OleRechDatable.Rows.Count <> 0 Then
                                                If EstNumeric(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone) = True Then

                                                Else
                                                    ExisteLecture = False
                                                    ErreurJrn.WriteLine("La valeur entrée : " & Trim(OleRechDatable.Rows(0).Item("Correspond")) & "doit être de type numérique pour l'info libre : " & Datagridaffiche.Columns(infoListe.Item(0)).Name & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
                                                End If
                                            Else
                                                If EstNumeric(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), DecimalNomb, DecimalMone) = True Then

                                                Else
                                                    ExisteLecture = False
                                                    ErreurJrn.WriteLine("La valeur entrée : " & Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) & "doit être de type numérique pour l'info libre : " & Datagridaffiche.Columns(infoListe.Item(0)).Name & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
                                                End If
                                            End If
                                        End If
                                    End If

                                    'Date Court
                                    If statistTab.Rows(0).Item("CB_Type") = 3 Then
                                        If Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) <> "" Then
                                            OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                            OleRecherDataset = New DataSet
                                            OleRecherAdapter.Fill(OleRecherDataset)
                                            OleRechDatable = OleRecherDataset.Tables(0)
                                            If OleRechDatable.Rows.Count <> 0 Then
                                                If Verificatdate(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier, Datagridaffiche.Columns(infoListe.Item(0)).Name) = True Then

                                                Else
                                                    ExisteLecture = False
                                                End If
                                            Else
                                                If Verificatdate(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), FormatDatefichier, Datagridaffiche.Columns(infoListe.Item(0)).Name) = True Then
                                                Else
                                                    ExisteLecture = False
                                                End If
                                            End If
                                        End If
                                    End If
                                    'Date Longue
                                    If statistTab.Rows(0).Item("CB_Type") = 14 Then
                                        If Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) <> "" Then
                                            OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Transfert' And Entete=True", OleConnenection)
                                            OleRecherDataset = New DataSet
                                            OleRecherAdapter.Fill(OleRecherDataset)
                                            OleRechDatable = OleRecherDataset.Tables(0)
                                            If OleRechDatable.Rows.Count <> 0 Then
                                                If Verificatdate(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier, Datagridaffiche.Columns(infoListe.Item(0)).Name) = True Then

                                                Else
                                                    ExisteLecture = False
                                                End If
                                            Else
                                                If Verificatdate(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), FormatDatefichier, Datagridaffiche.Columns(infoListe.Item(0)).Name) = True Then
                                                Else
                                                    ExisteLecture = False
                                                End If
                                            End If
                                        End If
                                    End If
                                Else
                                    ExisteLecture = False
                                    ErreurJrn.WriteLine(" Impossible de traiter l'information libre en Entête de Document :" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "  Il est inexistant dans la table CbsysLibre de Sage")
                                End If
                            End If
                        Else
                            'nothing
                        End If
                    Else
                        ExisteLecture = False
                        ErreurJrn.WriteLine(" Impossible de traiter l'information libre en Entête de Document :" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "  Il est inexistant dans la table de Paramétrage")
                    End If
                    'L'info Libre Parametrée par l'utilisateur n'existe pas dans Sage
                    infoListe.RemoveAt(0)
                End While
            End If
        Catch ex As Exception
            ExisteLecture = False
            exceptionTrouve = True
            ErreurJrn.WriteLine(" Erreur de Création de L'information Libre Entête de Document : " & ex.Message & ", vérifiez la longueur de la chaine avec celle paramétrée/ ou la cohérence des données")
        End Try
    End Sub
    Public Sub BT_integrer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_integrer.Click
        Dim i As Integer
        Dim ArtAdaptater As OleDbDataAdapter
        Dim ArtDataset As DataSet
        Dim Artdatatable As DataTable
        Dim CptaAdaptater As OleDbDataAdapter
        Dim CptaDataset As DataSet
        Dim Cptadatatable As DataTable
        Try
            If DataListeIntegrer.RowCount > 0 Then
                If Directory.Exists(Pathsfilejournal) Then
                    ErreurJrn = File.AppendText(Pathsfilejournal & "IMP_DOCUMENT_TRANSFERT" & Format(DateAndTime.Year(Now), "0000") & "" & Format(DateAndTime.Month(Now), "00") & "" & Format(DateAndTime.Day(Now), "00") & "_" & "" & Format(DateAndTime.Hour(Now), "00") & "_" & Format(DateAndTime.Minute(Now), "00") & "_" & Format(DateAndTime.Second(Now), "00") & ".txt")
                Else
                    Pathsfilejournal = "C:\"
                    ErreurJrn = File.AppendText(Pathsfilejournal & "IMP_DOCUMENT_TRANSFERT" & Format(DateAndTime.Year(Now), "0000") & "" & Format(DateAndTime.Month(Now), "00") & "" & Format(DateAndTime.Day(Now), "00") & "_" & "" & Format(DateAndTime.Hour(Now), "00") & "_" & Format(DateAndTime.Minute(Now), "00") & "_" & Format(DateAndTime.Second(Now), "00") & ".txt")
                End If
                For i = 0 To DataListeIntegrer.RowCount - 1
                    If DataListeIntegrer.Rows(i).Cells("Valider").Value = True Then
                        TraitementID = CInt(DataListeIntegrer.Rows(i).Cells("ID").Value)
                        Me.Cursor = Cursors.WaitCursor
                        exceptionTrouve = False
                        ExisteLecture = True
                        EntetePiecePrecedent = Nothing
                        ListePiece = New List(Of String)
                        If Trim(DataListeIntegrer.Rows(i).Cells("Mode").Value) = "Création" Then
                            'Format Creation 
                            If File.Exists(DataListeIntegrer.Rows(i).Cells("Chemin").Value) = True And File.Exists(DataListeIntegrer.Rows(i).Cells("CheminExport").Value) = True Then
                                If Trim(DataListeIntegrer.Rows(i).Cells("Comptable").Value) <> "" And Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) <> "" Then
                                    If Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value)) = "Délimité" Or Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value)) = "Tabulation" Or Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value)) = "Pipe" Then
                                        If Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value)) = "Tabulation" Then
                                            sColumnsSepar = ControlChars.Tab
                                        Else
                                            If Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value)) = "Pipe" Then
                                                sColumnsSepar = "|"
                                            Else
                                                sColumnsSepar = ";"
                                            End If
                                        End If
                                        ArtAdaptater = New OleDbDataAdapter("select * from PARAMETRE  Where  Societe='" & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & "' and nomtype='COMMERCIAL'", OleConnenection)
                                        ArtDataset = New DataSet
                                        ArtAdaptater.Fill(ArtDataset)
                                        Artdatatable = ArtDataset.Tables(0)
                                        If Artdatatable.Rows.Count <> 0 Then
                                            CptaAdaptater = New OleDbDataAdapter("select * from PARAMETRE  Where  Societe='" & Trim(DataListeIntegrer.Rows(i).Cells("Comptable").Value) & "' and nomtype='COMPTABILITE'", OleConnenection)
                                            CptaDataset = New DataSet
                                            CptaAdaptater.Fill(CptaDataset)
                                            Cptadatatable = CptaDataset.Tables(0)
                                            If Cptadatatable.Rows.Count <> 0 Then
                                                Dim Dataname() As String = Split(Trim(Cptadatatable.Rows(0).Item("Chemin1")), "\")
                                                If SocieteConnected(Strings.Left(Dataname(UBound(Dataname)), InStr(Dataname(UBound(Dataname)), ".") - 1), Trim(Cptadatatable.Rows(0).Item("MotPas").ToString), Trim(Cptadatatable.Rows(0).Item("NomUser")), LireChaine(Trim(Cptadatatable.Rows(0).Item("Chemin1")), "CBASE", "ServeurSQL")) = True Then
                                                    FermeBaseCial(BaseCial)
                                                    If OuvreBaseCial(BaseCial, BaseCpta, Trim(Artdatatable.Rows(0).Item("Chemin1")), Trim(Cptadatatable.Rows(0).Item("Chemin1")), Trim(Artdatatable.Rows(0).Item("UserSage")), Trim(Artdatatable.Rows(0).Item("PasseSage").ToString), Trim(Cptadatatable.Rows(0).Item("UserSage")), Trim(Cptadatatable.Rows(0).Item("PasseSage").ToString)) = True Then
                                                        ErreurJrn.WriteLine("Connexion à la Société " & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & " Reussie")
                                                        NomFichier = Trim(DataListeIntegrer.Rows(i).Cells("CheminExport").Value)
                                                        Do While InStr(Trim(NomFichier), "\") <> 0
                                                            NomFichier = Strings.Right(NomFichier, Strings.Len(Trim(NomFichier)) - InStr(Trim(NomFichier), "\"))
                                                        Loop
                                                        ErreurJrn.WriteLine("")
                                                        ErreurJrn.WriteLine("Début de traitement du fichier : " & NomFichier & " Date de traitement : " & DateTime.Today)
                                                        ErreurJrn.WriteLine("")
                                                        Label5.Refresh()
                                                        If Verification_Integration_Fichier(DataListeIntegrer.Rows(i).Cells("CheminExport").Value, DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value)), "", sColumnsSepar, RenvoieFormatDate(DataListeIntegrer.Rows(i).Cells("Chemin").Value), LirePieceCreation(DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value))), LirePieceAuto(DataListeIntegrer.Rows(i).Cells("Chemin").Value), LireTypeImport(DataListeIntegrer.Rows(i).Cells("Chemin").Value)) = True Then
                                                            Label5.Text = " Integration En Cours..."
                                                            Integration_Du_Fichier(DataListeIntegrer.Rows(i).Cells("CheminExport").Value, DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value)), "", sColumnsSepar, RenvoieFormatDate(DataListeIntegrer.Rows(i).Cells("Chemin").Value), LirePieceCreation(DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value))), LirePieceAuto(DataListeIntegrer.Rows(i).Cells("Chemin").Value), LireTypeImport(DataListeIntegrer.Rows(i).Cells("Chemin").Value))
                                                            RecuperationEnregistrement(DataListeIntegrer.Rows(i).Cells("CheminExport").Value, DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value)), "", sColumnsSepar, ListePiece, LirePieceCreation(DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value))), LirePieceAuto(DataListeIntegrer.Rows(i).Cells("Chemin").Value))
                                                            If DataListeIntegrer.Rows(i).Cells("Cible").Value = "FTP" Then
                                                                Dim FtpAdaptater As OleDbDataAdapter
                                                                Dim FtpDataset As DataSet
                                                                Dim Ftpdatatable As DataTable
                                                                FtpAdaptater = New OleDbDataAdapter("select * from  WIT_SCHEMA WHERE Cible='FTP' And IDDossier=" & CInt(DataListeIntegrer.Rows(i).Cells("ID").Value) & "", OleConnenection)
                                                                FtpDataset = New DataSet
                                                                FtpAdaptater.Fill(FtpDataset)
                                                                Ftpdatatable = FtpDataset.Tables(0)
                                                                If Ftpdatatable.Rows.Count <> 0 Then
                                                                    If DataListeIntegrer.Rows(i).Cells("Deplace").Value = True Then
                                                                        If exceptionTrouve = True Then
                                                                            'Il y'a une erreur au niveau de l'importation des articles à partir du fichier, on deplace pas le fichier
                                                                            File.Delete(DataListeIntegrer.Rows(i).Cells("CheminExport").Value)
                                                                            Label5.Refresh()
                                                                            Label5.Text = "Integration Terminée!"
                                                                        Else
                                                                            'Deplacement du fichier vers les repertoire de sauvegarde
                                                                            effaceFichier("FTP://" & RetourneServeurFtp(Ftpdatatable.Rows(0).Item("CheminFilexport")) & "/" & RetourneDirectoryFtp(Ftpdatatable.Rows(0).Item("CheminFilexport")) & "/" & System.IO.Path.GetFileName(DataListeIntegrer.Rows(i).Cells("CheminExport").Value), RetourneUserFtp(Ftpdatatable.Rows(0).Item("CheminFilexport")), RetournePassWordFtp(Ftpdatatable.Rows(0).Item("CheminFilexport")), ErreurJrn)
                                                                            File.Move(DataListeIntegrer.Rows(i).Cells("CheminExport").Value, PathsfileSave & "" & Strings.Right(DateAndTime.Year(Now), 2) & "" & Format(DateAndTime.Month(Now), "00") & "" & Format(DateAndTime.Day(Now), "00") & "[" & "" & Format(DateAndTime.Hour(Now), "00") & "-" & Format(DateAndTime.Minute(Now), "00") & "-" & Format(DateAndTime.Second(Now), "00") & "]" & NomFichier)
                                                                            Label5.Refresh()
                                                                            Label5.Text = "Integration Terminée! Suppression des Fichiers exécutée..."
                                                                        End If
                                                                    Else
                                                                        File.Delete(DataListeIntegrer.Rows(i).Cells("CheminExport").Value)
                                                                        Label5.Refresh()
                                                                        Label5.Text = "Integration Terminée!"
                                                                    End If
                                                                End If
                                                            Else
                                                                If DataListeIntegrer.Rows(i).Cells("Deplace").Value = True Then
                                                                    If exceptionTrouve = True Then
                                                                        'Il y'a une erreur au niveau de l'importation des articles à partir du fichier, on deplace pas le fichier
                                                                        Label5.Refresh()
                                                                        Label5.Text = "Integration Terminée!"
                                                                    Else
                                                                        'Deplacement du fichier vers les repertoire de sauvegarde
                                                                        File.Move(DataListeIntegrer.Rows(i).Cells("CheminExport").Value, PathsfileSave & "" & Strings.Right(DateAndTime.Year(Now), 2) & "" & Format(DateAndTime.Month(Now), "00") & "" & Format(DateAndTime.Day(Now), "00") & "[" & "" & Format(DateAndTime.Hour(Now), "00") & "-" & Format(DateAndTime.Minute(Now), "00") & "-" & Format(DateAndTime.Second(Now), "00") & "]" & NomFichier)
                                                                        Label5.Refresh()
                                                                        Label5.Text = "Integration Terminée! Suppression des Fichiers exécutée..."
                                                                    End If
                                                                Else
                                                                    Label5.Refresh()
                                                                    Label5.Text = "Integration Terminée!"
                                                                End If
                                                            End If
                                                        End If
                                                        DataListeIntegrer.Rows(i).Cells("Valider").Value = False
                                                    Else
                                                        ErreurJrn.WriteLine("Connexion à la Société - Base Commerciale :" & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & " -Base Comptable :" & Trim(DataListeIntegrer.Rows(i).Cells("Comptable").Value) & " Echec de traitement")
                                                        Label5.Text = "Connexion à la Société " & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & " : Echec"
                                                    End If
                                                Else
                                                    ErreurJrn.WriteLine("Echec de Connexion à SQL de base de données :" & Strings.Left(Dataname(UBound(Dataname)), InStr(Dataname(UBound(Dataname)), ".") - 1) & " Serveur : " & LireChaine(Trim(Cptadatatable.Rows(0).Item("Chemin1")), "CBASE", "ServeurSQL"))
                                                    Label5.Text = "Echec de Connexion SQL à la Société " & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & " : Echec de traitement"
                                                End If
                                            Else
                                                ErreurJrn.WriteLine("Aucune Base Comptable Correspondant à : " & Trim(DataListeIntegrer.Rows(i).Cells("Comptable").Value) & " Echec de traitement")
                                            End If
                                        Else
                                            ErreurJrn.WriteLine("Aucune Base Commerciale Correspondant à : " & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & " Echec de traitement")
                                        End If

                                    Else
                                        If Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value)) = "Excel" Then
                                            If Trim(DataListeIntegrer.Rows(i).Cells("FeuilleExcel").Value) <> "" Then
                                                If OleExcelConnected.State = ConnectionState.Open Then
                                                    OleExcelConnected.Close()
                                                End If
                                                If LoginAuFichierExcel(Trim(DataListeIntegrer.Rows(i).Cells("CheminExport").Value)) = True Then
                                                    ArtAdaptater = New OleDbDataAdapter("select * from PARAMETRE  Where  Societe='" & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & "' and nomtype='COMMERCIAL'", OleConnenection)
                                                    ArtDataset = New DataSet
                                                    ArtAdaptater.Fill(ArtDataset)
                                                    Artdatatable = ArtDataset.Tables(0)
                                                    If Artdatatable.Rows.Count <> 0 Then
                                                        CptaAdaptater = New OleDbDataAdapter("select * from PARAMETRE  Where  Societe='" & Trim(DataListeIntegrer.Rows(i).Cells("Comptable").Value) & "' and nomtype='COMPTABILITE'", OleConnenection)
                                                        CptaDataset = New DataSet
                                                        CptaAdaptater.Fill(CptaDataset)
                                                        Cptadatatable = CptaDataset.Tables(0)
                                                        If Cptadatatable.Rows.Count <> 0 Then
                                                            Dim Dataname() As String = Split(Trim(Cptadatatable.Rows(0).Item("Chemin1")), "\")
                                                            If SocieteConnected(Strings.Left(Dataname(UBound(Dataname)), InStr(Dataname(UBound(Dataname)), ".") - 1), Trim(Cptadatatable.Rows(0).Item("MotPas").ToString), Trim(Cptadatatable.Rows(0).Item("NomUser")), LireChaine(Trim(Cptadatatable.Rows(0).Item("Chemin1")), "CBASE", "ServeurSQL")) = True Then
                                                                FermeBaseCial(BaseCial)
                                                                If OuvreBaseCial(BaseCial, BaseCpta, Trim(Artdatatable.Rows(0).Item("Chemin1")), Trim(Cptadatatable.Rows(0).Item("Chemin1")), Trim(Artdatatable.Rows(0).Item("UserSage")), Trim(Artdatatable.Rows(0).Item("PasseSage").ToString), Trim(Cptadatatable.Rows(0).Item("UserSage")), Trim(Cptadatatable.Rows(0).Item("PasseSage").ToString)) = True Then
                                                                    ErreurJrn.WriteLine("Connexion à la Société " & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & " Reussie")
                                                                    NomFichier = Trim(DataListeIntegrer.Rows(i).Cells("CheminExport").Value)
                                                                    Do While InStr(Trim(NomFichier), "\") <> 0
                                                                        NomFichier = Strings.Right(NomFichier, Strings.Len(Trim(NomFichier)) - InStr(Trim(NomFichier), "\"))
                                                                    Loop
                                                                    ErreurJrn.WriteLine("")
                                                                    ErreurJrn.WriteLine("Début de traitement du fichier : " & NomFichier)
                                                                    ErreurJrn.WriteLine("")
                                                                    Label5.Refresh()
                                                                    If Verification_Integration_Fichier(DataListeIntegrer.Rows(i).Cells("CheminExport").Value, DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value)), Trim(DataListeIntegrer.Rows(i).Cells("FeuilleExcel").Value), "", RenvoieFormatDate(DataListeIntegrer.Rows(i).Cells("Chemin").Value), LirePieceCreation(DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value))), LirePieceAuto(DataListeIntegrer.Rows(i).Cells("Chemin").Value), LireTypeImport(DataListeIntegrer.Rows(i).Cells("Chemin").Value)) = True Then
                                                                        Label5.Text = "Integration En Cours..."
                                                                        Integration_Du_Fichier(DataListeIntegrer.Rows(i).Cells("CheminExport").Value, DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value)), Trim(DataListeIntegrer.Rows(i).Cells("FeuilleExcel").Value), "", RenvoieFormatDate(DataListeIntegrer.Rows(i).Cells("Chemin").Value), LirePieceCreation(DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value))), LirePieceAuto(DataListeIntegrer.Rows(i).Cells("Chemin").Value), LireTypeImport(DataListeIntegrer.Rows(i).Cells("Chemin").Value))
                                                                        RecuperationEnregistrement(DataListeIntegrer.Rows(i).Cells("CheminExport").Value, DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value)), Trim(DataListeIntegrer.Rows(i).Cells("FeuilleExcel").Value), "", ListePiece, LirePieceCreation(DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value))), LirePieceAuto(DataListeIntegrer.Rows(i).Cells("Chemin").Value))
                                                                        If DataListeIntegrer.Rows(i).Cells("Cible").Value = "FTP" Then
                                                                            Dim FtpAdaptater As OleDbDataAdapter
                                                                            Dim FtpDataset As DataSet
                                                                            Dim Ftpdatatable As DataTable
                                                                            FtpAdaptater = New OleDbDataAdapter("select * from  WIT_SCHEMA WHERE Cible='FTP' And IDDossier=" & CInt(DataListeIntegrer.Rows(i).Cells("ID").Value) & "", OleConnenection)
                                                                            FtpDataset = New DataSet
                                                                            FtpAdaptater.Fill(FtpDataset)
                                                                            Ftpdatatable = FtpDataset.Tables(0)
                                                                            If Ftpdatatable.Rows.Count <> 0 Then
                                                                                If DataListeIntegrer.Rows(i).Cells("Deplace").Value = True Then
                                                                                    If exceptionTrouve = True Then
                                                                                        'Il y'a une erreur au niveau de l'importation des articles à partir du fichier, on deplace pas le fichier
                                                                                        File.Delete(DataListeIntegrer.Rows(i).Cells("CheminExport").Value)
                                                                                        Label5.Refresh()
                                                                                        Label5.Text = "Integration Terminée!"
                                                                                    Else
                                                                                        'Deplacement du fichier vers les repertoire de sauvegarde
                                                                                        effaceFichier("FTP://" & RetourneServeurFtp(Ftpdatatable.Rows(0).Item("CheminFilexport")) & "/" & RetourneDirectoryFtp(Ftpdatatable.Rows(0).Item("CheminFilexport")) & "/" & System.IO.Path.GetFileName(DataListeIntegrer.Rows(i).Cells("CheminExport").Value), RetourneUserFtp(Ftpdatatable.Rows(0).Item("CheminFilexport")), RetournePassWordFtp(Ftpdatatable.Rows(0).Item("CheminFilexport")), ErreurJrn)
                                                                                        File.Move(DataListeIntegrer.Rows(i).Cells("CheminExport").Value, PathsfileSave & "" & Strings.Right(DateAndTime.Year(Now), 2) & "" & Format(DateAndTime.Month(Now), "00") & "" & Format(DateAndTime.Day(Now), "00") & "[" & "" & Format(DateAndTime.Hour(Now), "00") & "-" & Format(DateAndTime.Minute(Now), "00") & "-" & Format(DateAndTime.Second(Now), "00") & "]" & NomFichier)
                                                                                        Label5.Refresh()
                                                                                        Label5.Text = "Integration Terminée! Suppression des Fichiers exécutée..."
                                                                                    End If
                                                                                Else
                                                                                    File.Delete(DataListeIntegrer.Rows(i).Cells("CheminExport").Value)
                                                                                    Label5.Refresh()
                                                                                    Label5.Text = "Integration Terminée!"
                                                                                End If
                                                                            End If
                                                                        Else
                                                                            If DataListeIntegrer.Rows(i).Cells("Deplace").Value = True Then
                                                                                If exceptionTrouve = True Then
                                                                                    'Il y'a une erreur au niveau de l'importation des articles à partir du fichier, on deplace pas le fichier
                                                                                    Label5.Refresh()
                                                                                    Label5.Text = "Integration Terminée!"
                                                                                Else
                                                                                    'Deplacement du fichier vers les repertoire de sauvegarde
                                                                                    File.Move(DataListeIntegrer.Rows(i).Cells("CheminExport").Value, PathsfileSave & "" & Strings.Right(DateAndTime.Year(Now), 2) & "" & Format(DateAndTime.Month(Now), "00") & "" & Format(DateAndTime.Day(Now), "00") & "[" & "" & Format(DateAndTime.Hour(Now), "00") & "-" & Format(DateAndTime.Minute(Now), "00") & "-" & Format(DateAndTime.Second(Now), "00") & "]" & NomFichier)
                                                                                    Label5.Refresh()
                                                                                    Label5.Text = "Integration Terminée! Suppression des Fichiers exécutée..."
                                                                                End If
                                                                            Else
                                                                                Label5.Refresh()
                                                                                Label5.Text = "Integration Terminée!"
                                                                            End If
                                                                        End If
                                                                    End If
                                                                    If OleExcelConnected.State = ConnectionState.Open Then
                                                                        OleExcelConnected.Close()
                                                                    End If
                                                                    DataListeIntegrer.Rows(i).Cells("Valider").Value = False
                                                                Else
                                                                    ErreurJrn.WriteLine("Connexion à la Société - Base Commerciale :" & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & " -Base Comptable :" & Trim(DataListeIntegrer.Rows(i).Cells("Comptable").Value) & " Echec")
                                                                    Label5.Text = "Connexion à la Société " & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & " : Echec"
                                                                End If
                                                            Else
                                                                ErreurJrn.WriteLine("Echec de Connexion à SQL de base de données :" & Strings.Left(Dataname(UBound(Dataname)), InStr(Dataname(UBound(Dataname)), ".") - 1) & " Serveur : " & LireChaine(Trim(Cptadatatable.Rows(0).Item("Chemin1")), "CBASE", "ServeurSQL"))
                                                                Label5.Text = "Echec de Connexion SQL à la Société " & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & " : Echec de traitement"
                                                            End If
                                                        Else
                                                            ErreurJrn.WriteLine("Aucune Base Comptable Correspondant à : " & Trim(DataListeIntegrer.Rows(i).Cells("Comptable").Value) & " Echec de traitement")
                                                        End If
                                                    Else
                                                        ErreurJrn.WriteLine("Aucune Base Commerciale Correspondant à : " & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & " Echec de traitement")
                                                    End If

                                                Else
                                                    Label5.Text = "Echec de Connexion au fichier Excel :" & Trim(DataListeIntegrer.Rows(i).Cells("CheminExport").Value) & " : Echec de traitement"
                                                    ErreurJrn.WriteLine("Echec de Connexion au fichier Excel :" & Trim(DataListeIntegrer.Rows(i).Cells("CheminExport").Value) & " : Echec de traitement")
                                                End If
                                            Else
                                                Label5.Text = "Aucune Feuille Excel paramétrée pour le fichier :" & Trim(DataListeIntegrer.Rows(i).Cells("CheminExport").Value) & " : Echec de traitement"
                                                ErreurJrn.WriteLine("Aucune Feuille Excel paramétrée pour le fichier :" & Trim(DataListeIntegrer.Rows(i).Cells("CheminExport").Value) & " : Echec de traitement")
                                            End If
                                        Else
                                            If Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value)) = "Longueur Fixe" Then
                                                ArtAdaptater = New OleDbDataAdapter("select * from PARAMETRE  Where  Societe='" & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & "' and nomtype='COMMERCIAL'", OleConnenection)
                                                ArtDataset = New DataSet
                                                ArtAdaptater.Fill(ArtDataset)
                                                Artdatatable = ArtDataset.Tables(0)
                                                If Artdatatable.Rows.Count <> 0 Then
                                                    CptaAdaptater = New OleDbDataAdapter("select * from PARAMETRE  Where  Societe='" & Trim(DataListeIntegrer.Rows(i).Cells("Comptable").Value) & "' and nomtype='COMPTABILITE'", OleConnenection)
                                                    CptaDataset = New DataSet
                                                    CptaAdaptater.Fill(CptaDataset)
                                                    Cptadatatable = CptaDataset.Tables(0)
                                                    If Cptadatatable.Rows.Count <> 0 Then
                                                        Dim Dataname() As String = Split(Trim(Cptadatatable.Rows(0).Item("Chemin1")), "\")
                                                        If SocieteConnected(Strings.Left(Dataname(UBound(Dataname)), InStr(Dataname(UBound(Dataname)), ".") - 1), Trim(Cptadatatable.Rows(0).Item("MotPas").ToString), Trim(Cptadatatable.Rows(0).Item("NomUser")), LireChaine(Trim(Cptadatatable.Rows(0).Item("Chemin1")), "CBASE", "ServeurSQL")) = True Then
                                                            FermeBaseCial(BaseCial)
                                                            If OuvreBaseCial(BaseCial, BaseCpta, Trim(Artdatatable.Rows(0).Item("Chemin1")), Trim(Cptadatatable.Rows(0).Item("Chemin1")), Trim(Artdatatable.Rows(0).Item("UserSage")), Trim(Artdatatable.Rows(0).Item("PasseSage").ToString), Trim(Cptadatatable.Rows(0).Item("UserSage")), Trim(Cptadatatable.Rows(0).Item("PasseSage").ToString)) = True Then
                                                                ErreurJrn.WriteLine("Connexion à la Société " & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & " Reussie")
                                                                Label5.Refresh()
                                                                NomFichier = Trim(DataListeIntegrer.Rows(i).Cells("CheminExport").Value)
                                                                Do While InStr(Trim(NomFichier), "\") <> 0
                                                                    NomFichier = Strings.Right(NomFichier, Strings.Len(Trim(NomFichier)) - InStr(Trim(NomFichier), "\"))
                                                                Loop
                                                                ErreurJrn.WriteLine("")
                                                                ErreurJrn.WriteLine("Début de traitement du fichier : " & NomFichier)
                                                                ErreurJrn.WriteLine("")
                                                                Label5.Refresh()
                                                                If Verification_Integration_Fichier(DataListeIntegrer.Rows(i).Cells("CheminExport").Value, DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value)), "", "", RenvoieFormatDate(DataListeIntegrer.Rows(i).Cells("Chemin").Value), LirePieceCreation(DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value))), LirePieceAuto(DataListeIntegrer.Rows(i).Cells("Chemin").Value), LireTypeImport(DataListeIntegrer.Rows(i).Cells("Chemin").Value)) = True Then
                                                                    Label5.Text = "Integration En Cours..."
                                                                    Integration_Du_Fichier(DataListeIntegrer.Rows(i).Cells("CheminExport").Value, DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value)), "", "", RenvoieFormatDate(DataListeIntegrer.Rows(i).Cells("Chemin").Value), LirePieceCreation(DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value))), LirePieceAuto(DataListeIntegrer.Rows(i).Cells("Chemin").Value), LireTypeImport(DataListeIntegrer.Rows(i).Cells("Chemin").Value))
                                                                    RecuperationEnregistrement(DataListeIntegrer.Rows(i).Cells("CheminExport").Value, DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value)), "", "", ListePiece, LirePieceCreation(DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value))), LirePieceAuto(DataListeIntegrer.Rows(i).Cells("Chemin").Value))
                                                                    If DataListeIntegrer.Rows(i).Cells("Cible").Value = "FTP" Then
                                                                        Dim FtpAdaptater As OleDbDataAdapter
                                                                        Dim FtpDataset As DataSet
                                                                        Dim Ftpdatatable As DataTable
                                                                        FtpAdaptater = New OleDbDataAdapter("select * from  WIT_SCHEMA WHERE Cible='FTP' And IDDossier=" & CInt(DataListeIntegrer.Rows(i).Cells("ID").Value) & "", OleConnenection)
                                                                        FtpDataset = New DataSet
                                                                        FtpAdaptater.Fill(FtpDataset)
                                                                        Ftpdatatable = FtpDataset.Tables(0)
                                                                        If Ftpdatatable.Rows.Count <> 0 Then
                                                                            If DataListeIntegrer.Rows(i).Cells("Deplace").Value = True Then
                                                                                If exceptionTrouve = True Then
                                                                                    'Il y'a une erreur au niveau de l'importation des articles à partir du fichier, on deplace pas le fichier
                                                                                    File.Delete(DataListeIntegrer.Rows(i).Cells("CheminExport").Value)
                                                                                    Label5.Refresh()
                                                                                    Label5.Text = "Integration Terminée!"
                                                                                Else
                                                                                    'Deplacement du fichier vers les repertoire de sauvegarde
                                                                                    effaceFichier("FTP://" & RetourneServeurFtp(Ftpdatatable.Rows(0).Item("CheminFilexport")) & "/" & RetourneDirectoryFtp(Ftpdatatable.Rows(0).Item("CheminFilexport")) & "/" & System.IO.Path.GetFileName(DataListeIntegrer.Rows(i).Cells("CheminExport").Value), RetourneUserFtp(Ftpdatatable.Rows(0).Item("CheminFilexport")), RetournePassWordFtp(Ftpdatatable.Rows(0).Item("CheminFilexport")), ErreurJrn)
                                                                                    File.Move(DataListeIntegrer.Rows(i).Cells("CheminExport").Value, PathsfileSave & "" & Strings.Right(DateAndTime.Year(Now), 2) & "" & Format(DateAndTime.Month(Now), "00") & "" & Format(DateAndTime.Day(Now), "00") & "[" & "" & Format(DateAndTime.Hour(Now), "00") & "-" & Format(DateAndTime.Minute(Now), "00") & "-" & Format(DateAndTime.Second(Now), "00") & "]" & NomFichier)
                                                                                    Label5.Refresh()
                                                                                    Label5.Text = "Integration Terminée! Suppression des Fichiers exécutée..."
                                                                                End If
                                                                            Else
                                                                                File.Delete(DataListeIntegrer.Rows(i).Cells("CheminExport").Value)
                                                                                Label5.Refresh()
                                                                                Label5.Text = "Integration Terminée!"
                                                                            End If
                                                                        End If
                                                                    Else
                                                                        If DataListeIntegrer.Rows(i).Cells("Deplace").Value = True Then
                                                                            If exceptionTrouve = True Then
                                                                                'Il y'a une erreur au niveau de l'importation des articles à partir du fichier, on deplace pas le fichier
                                                                                Label5.Refresh()
                                                                                Label5.Text = "Integration Terminée!"
                                                                            Else
                                                                                'Deplacement du fichier vers les repertoire de sauvegarde
                                                                                File.Move(DataListeIntegrer.Rows(i).Cells("CheminExport").Value, PathsfileSave & "" & Strings.Right(DateAndTime.Year(Now), 2) & "" & Format(DateAndTime.Month(Now), "00") & "" & Format(DateAndTime.Day(Now), "00") & "[" & "" & Format(DateAndTime.Hour(Now), "00") & "-" & Format(DateAndTime.Minute(Now), "00") & "-" & Format(DateAndTime.Second(Now), "00") & "]" & NomFichier)
                                                                                Label5.Refresh()
                                                                                Label5.Text = "Integration Terminée! Suppression des Fichiers exécutée..."
                                                                            End If
                                                                        Else
                                                                            Label5.Refresh()
                                                                            Label5.Text = "Integration Terminée!"
                                                                        End If
                                                                    End If
                                                                End If
                                                                DataListeIntegrer.Rows(i).Cells("Valider").Value = False
                                                            Else
                                                                ErreurJrn.WriteLine("Connexion à la Société - Base Commerciale :" & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & " -Base Comptable :" & Trim(DataListeIntegrer.Rows(i).Cells("Comptable").Value) & " Echec")
                                                                Label5.Text = "Connexion à la Société " & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & " : Echec"
                                                            End If
                                                        Else
                                                            ErreurJrn.WriteLine("Echec de Connexion à SQL de base de données :" & Strings.Left(Dataname(UBound(Dataname)), InStr(Dataname(UBound(Dataname)), ".") - 1) & " Serveur : " & LireChaine(Trim(Cptadatatable.Rows(0).Item("Chemin1")), "CBASE", "ServeurSQL"))
                                                            Label5.Text = "Echec de Connexion SQL à la Société " & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & " : Echec de traitement"
                                                        End If
                                                    Else
                                                        ErreurJrn.WriteLine("Aucune Base Comptable Correspondant à : " & Trim(DataListeIntegrer.Rows(i).Cells("Comptable").Value) & " Echec de traitement")
                                                    End If
                                                Else
                                                    ErreurJrn.WriteLine("Aucune Base Commerciale Correspondant à : " & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & " Echec de traitement")
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                ErreurJrn.WriteLine("Chemin du fichier inexistant : " & DataListeIntegrer.Rows(i).Cells("Chemin").Value)
                            End If
                        Else
                            If Trim(DataListeIntegrer.Rows(i).Cells("Mode").Value) = "Modification" Then
                                'Format Modification
                                If File.Exists(DataListeIntegrer.Rows(i).Cells("Chemin").Value) = True And File.Exists(DataListeIntegrer.Rows(i).Cells("CheminExport").Value) = True Then
                                    If Trim(DataListeIntegrer.Rows(i).Cells("Comptable").Value) <> "" And Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) <> "" Then
                                        If Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value)) = "Délimité" Or Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value)) = "Tabulation" Or Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value)) = "Pipe" Then
                                            If Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value)) = "Tabulation" Then
                                                sColumnsSepar = ControlChars.Tab
                                            Else
                                                If Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value)) = "Pipe" Then
                                                    sColumnsSepar = "|"
                                                Else
                                                    sColumnsSepar = ";"
                                                End If
                                            End If
                                            ArtAdaptater = New OleDbDataAdapter("select * from PARAMETRE  Where  Societe='" & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & "' and nomtype='COMMERCIAL'", OleConnenection)
                                            ArtDataset = New DataSet
                                            ArtAdaptater.Fill(ArtDataset)
                                            Artdatatable = ArtDataset.Tables(0)
                                            If Artdatatable.Rows.Count <> 0 Then
                                                CptaAdaptater = New OleDbDataAdapter("select * from PARAMETRE  Where  Societe='" & Trim(DataListeIntegrer.Rows(i).Cells("Comptable").Value) & "' and nomtype='COMPTABILITE'", OleConnenection)
                                                CptaDataset = New DataSet
                                                CptaAdaptater.Fill(CptaDataset)
                                                Cptadatatable = CptaDataset.Tables(0)
                                                If Cptadatatable.Rows.Count <> 0 Then
                                                    Dim Dataname() As String = Split(Trim(Cptadatatable.Rows(0).Item("Chemin1")), "\")
                                                    If SocieteConnected(Strings.Left(Dataname(UBound(Dataname)), InStr(Dataname(UBound(Dataname)), ".") - 1), Trim(Cptadatatable.Rows(0).Item("MotPas").ToString), Trim(Cptadatatable.Rows(0).Item("NomUser")), LireChaine(Trim(Cptadatatable.Rows(0).Item("Chemin1")), "CBASE", "ServeurSQL")) = True Then
                                                        FermeBaseCial(BaseCial)
                                                        If OuvreBaseCial(BaseCial, BaseCpta, Trim(Artdatatable.Rows(0).Item("Chemin1")), Trim(Cptadatatable.Rows(0).Item("Chemin1")), Trim(Artdatatable.Rows(0).Item("UserSage")), Trim(Artdatatable.Rows(0).Item("PasseSage").ToString), Trim(Cptadatatable.Rows(0).Item("UserSage")), Trim(Cptadatatable.Rows(0).Item("PasseSage").ToString)) = True Then
                                                            ErreurJrn.WriteLine("Connexion à la Société " & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & " Reussie")
                                                            NomFichier = Trim(DataListeIntegrer.Rows(i).Cells("CheminExport").Value)
                                                            Do While InStr(Trim(NomFichier), "\") <> 0
                                                                NomFichier = Strings.Right(NomFichier, Strings.Len(Trim(NomFichier)) - InStr(Trim(NomFichier), "\"))
                                                            Loop
                                                            ErreurJrn.WriteLine("")
                                                            ErreurJrn.WriteLine("Début de traitement du fichier : " & NomFichier)
                                                            ErreurJrn.WriteLine("")
                                                            Label5.Refresh()
                                                            If Modification_Verification_Fichier(DataListeIntegrer.Rows(i).Cells("CheminExport").Value, DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value)), "", sColumnsSepar, LirePieceModification(DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value))), RenvoieFormatDate(DataListeIntegrer.Rows(i).Cells("Chemin").Value)) = True Then
                                                                Label5.Text = " Modification En Cours..."
                                                                Modification_Integration_Fichier(DataListeIntegrer.Rows(i).Cells("CheminExport").Value, DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value)), "", sColumnsSepar, LirePieceModification(DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value))), RenvoieFormatDate(DataListeIntegrer.Rows(i).Cells("Chemin").Value))
                                                                RecuperationEnregistrementModifié(DataListeIntegrer.Rows(i).Cells("CheminExport").Value, DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value)), "", sColumnsSepar, ListePiece, LirePieceModification(DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value))))
                                                                If DataListeIntegrer.Rows(i).Cells("Cible").Value = "FTP" Then
                                                                    Dim FtpAdaptater As OleDbDataAdapter
                                                                    Dim FtpDataset As DataSet
                                                                    Dim Ftpdatatable As DataTable
                                                                    FtpAdaptater = New OleDbDataAdapter("select * from  WIT_SCHEMA WHERE Cible='FTP' And IDDossier=" & CInt(DataListeIntegrer.Rows(i).Cells("ID").Value) & "", OleConnenection)
                                                                    FtpDataset = New DataSet
                                                                    FtpAdaptater.Fill(FtpDataset)
                                                                    Ftpdatatable = FtpDataset.Tables(0)
                                                                    If Ftpdatatable.Rows.Count <> 0 Then
                                                                        If DataListeIntegrer.Rows(i).Cells("Deplace").Value = True Then
                                                                            If exceptionTrouve = True Then
                                                                                'Il y'a une erreur au niveau de l'importation des articles à partir du fichier, on deplace pas le fichier
                                                                                File.Delete(DataListeIntegrer.Rows(i).Cells("CheminExport").Value)
                                                                                Label5.Refresh()
                                                                                Label5.Text = "Modification Terminée!"
                                                                            Else
                                                                                'Deplacement du fichier vers les repertoire de sauvegarde
                                                                                effaceFichier("FTP://" & RetourneServeurFtp(Ftpdatatable.Rows(0).Item("CheminFilexport")) & "/" & RetourneDirectoryFtp(Ftpdatatable.Rows(0).Item("CheminFilexport")) & "/" & System.IO.Path.GetFileName(DataListeIntegrer.Rows(i).Cells("CheminExport").Value), RetourneUserFtp(Ftpdatatable.Rows(0).Item("CheminFilexport")), RetournePassWordFtp(Ftpdatatable.Rows(0).Item("CheminFilexport")), ErreurJrn)
                                                                                File.Move(DataListeIntegrer.Rows(i).Cells("CheminExport").Value, PathsfileSave & "" & Strings.Right(DateAndTime.Year(Now), 2) & "" & Format(DateAndTime.Month(Now), "00") & "" & Format(DateAndTime.Day(Now), "00") & "[" & "" & Format(DateAndTime.Hour(Now), "00") & "-" & Format(DateAndTime.Minute(Now), "00") & "-" & Format(DateAndTime.Second(Now), "00") & "]" & NomFichier)
                                                                                Label5.Refresh()
                                                                                Label5.Text = "Modification Terminée! Suppression des Fichiers exécutée..."
                                                                            End If
                                                                        Else
                                                                            File.Delete(DataListeIntegrer.Rows(i).Cells("CheminExport").Value)
                                                                            Label5.Refresh()
                                                                            Label5.Text = "Modification Terminée!"
                                                                        End If
                                                                    End If
                                                                Else
                                                                    If DataListeIntegrer.Rows(i).Cells("Deplace").Value = True Then
                                                                        If exceptionTrouve = True Then
                                                                            'Il y'a une erreur au niveau de l'importation des articles à partir du fichier, on deplace pas le fichier
                                                                            Label5.Refresh()
                                                                            Label5.Text = "Modification Terminée!"
                                                                        Else
                                                                            'Deplacement du fichier vers les repertoire de sauvegarde
                                                                            File.Move(DataListeIntegrer.Rows(i).Cells("CheminExport").Value, PathsfileSave & "" & Strings.Right(DateAndTime.Year(Now), 2) & "" & Format(DateAndTime.Month(Now), "00") & "" & Format(DateAndTime.Day(Now), "00") & "[" & "" & Format(DateAndTime.Hour(Now), "00") & "-" & Format(DateAndTime.Minute(Now), "00") & "-" & Format(DateAndTime.Second(Now), "00") & "]" & NomFichier)
                                                                            Label5.Refresh()
                                                                            Label5.Text = "Modification Terminée! Suppression des Fichiers exécutée..."
                                                                        End If
                                                                    Else
                                                                        Label5.Refresh()
                                                                        Label5.Text = "Modification Terminée!"
                                                                    End If
                                                                End If
                                                            End If
                                                            DataListeIntegrer.Rows(i).Cells("Valider").Value = False
                                                        Else
                                                            ErreurJrn.WriteLine("Connexion à la Société - Base Commerciale :" & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & " -Base Comptable :" & Trim(DataListeIntegrer.Rows(i).Cells("Comptable").Value) & " Echec de traitement")
                                                            Label5.Text = "Connexion à la Société " & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & " : Echec"
                                                        End If
                                                    Else
                                                        ErreurJrn.WriteLine("Echec de Connexion à SQL de base de données :" & Strings.Left(Dataname(UBound(Dataname)), InStr(Dataname(UBound(Dataname)), ".") - 1) & " Serveur : " & LireChaine(Trim(Cptadatatable.Rows(0).Item("Chemin1")), "CBASE", "ServeurSQL"))
                                                        Label5.Text = "Echec de Connexion SQL à la Société " & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & " : Echec de traitement"
                                                    End If
                                                Else
                                                    ErreurJrn.WriteLine("Aucune Base Comptable Correspondant à : " & Trim(DataListeIntegrer.Rows(i).Cells("Comptable").Value) & " Echec de traitement")
                                                End If
                                            Else
                                                ErreurJrn.WriteLine("Aucune Base Commerciale Correspondant à : " & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & " Echec de traitement")
                                            End If

                                        Else
                                            If Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value)) = "Excel" Then
                                                If Trim(DataListeIntegrer.Rows(i).Cells("FeuilleExcel").Value) <> "" Then
                                                    If OleExcelConnected.State = ConnectionState.Open Then
                                                        OleExcelConnected.Close()
                                                    End If
                                                    If LoginAuFichierExcel(Trim(DataListeIntegrer.Rows(i).Cells("CheminExport").Value)) = True Then
                                                        ArtAdaptater = New OleDbDataAdapter("select * from PARAMETRE  Where  Societe='" & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & "' and nomtype='COMMERCIAL'", OleConnenection)
                                                        ArtDataset = New DataSet
                                                        ArtAdaptater.Fill(ArtDataset)
                                                        Artdatatable = ArtDataset.Tables(0)
                                                        If Artdatatable.Rows.Count <> 0 Then
                                                            CptaAdaptater = New OleDbDataAdapter("select * from PARAMETRE  Where  Societe='" & Trim(DataListeIntegrer.Rows(i).Cells("Comptable").Value) & "' and nomtype='COMPTABILITE'", OleConnenection)
                                                            CptaDataset = New DataSet
                                                            CptaAdaptater.Fill(CptaDataset)
                                                            Cptadatatable = CptaDataset.Tables(0)
                                                            If Cptadatatable.Rows.Count <> 0 Then
                                                                Dim Dataname() As String = Split(Trim(Cptadatatable.Rows(0).Item("Chemin1")), "\")
                                                                If SocieteConnected(Strings.Left(Dataname(UBound(Dataname)), InStr(Dataname(UBound(Dataname)), ".") - 1), Trim(Cptadatatable.Rows(0).Item("MotPas").ToString), Trim(Cptadatatable.Rows(0).Item("NomUser")), LireChaine(Trim(Cptadatatable.Rows(0).Item("Chemin1")), "CBASE", "ServeurSQL")) = True Then
                                                                    FermeBaseCial(BaseCial)
                                                                    If OuvreBaseCial(BaseCial, BaseCpta, Trim(Artdatatable.Rows(0).Item("Chemin1")), Trim(Cptadatatable.Rows(0).Item("Chemin1")), Trim(Artdatatable.Rows(0).Item("UserSage")), Trim(Artdatatable.Rows(0).Item("PasseSage").ToString), Trim(Cptadatatable.Rows(0).Item("UserSage")), Trim(Cptadatatable.Rows(0).Item("PasseSage").ToString)) = True Then
                                                                        ErreurJrn.WriteLine("Connexion à la Société " & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & " Reussie")
                                                                        NomFichier = Trim(DataListeIntegrer.Rows(i).Cells("CheminExport").Value)
                                                                        Do While InStr(Trim(NomFichier), "\") <> 0
                                                                            NomFichier = Strings.Right(NomFichier, Strings.Len(Trim(NomFichier)) - InStr(Trim(NomFichier), "\"))
                                                                        Loop
                                                                        ErreurJrn.WriteLine("")
                                                                        ErreurJrn.WriteLine("Début de traitement du fichier : " & NomFichier)
                                                                        ErreurJrn.WriteLine("")
                                                                        Label5.Refresh()
                                                                        If Modification_Verification_Fichier(DataListeIntegrer.Rows(i).Cells("CheminExport").Value, DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value)), Trim(DataListeIntegrer.Rows(i).Cells("FeuilleExcel").Value), "", LirePieceModification(DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value))), RenvoieFormatDate(DataListeIntegrer.Rows(i).Cells("Chemin").Value)) = True Then
                                                                            Label5.Text = "Modification En Cours..."
                                                                            Modification_Integration_Fichier(DataListeIntegrer.Rows(i).Cells("CheminExport").Value, DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value)), Trim(DataListeIntegrer.Rows(i).Cells("FeuilleExcel").Value), "", LirePieceModification(DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value))), RenvoieFormatDate(DataListeIntegrer.Rows(i).Cells("Chemin").Value))
                                                                            RecuperationEnregistrementModifié(DataListeIntegrer.Rows(i).Cells("CheminExport").Value, DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value)), Trim(DataListeIntegrer.Rows(i).Cells("FeuilleExcel").Value), "", ListePiece, LirePieceModification(DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value))))
                                                                            If DataListeIntegrer.Rows(i).Cells("Cible").Value = "FTP" Then
                                                                                Dim FtpAdaptater As OleDbDataAdapter
                                                                                Dim FtpDataset As DataSet
                                                                                Dim Ftpdatatable As DataTable
                                                                                FtpAdaptater = New OleDbDataAdapter("select * from  WIT_SCHEMA WHERE Cible='FTP' And IDDossier=" & CInt(DataListeIntegrer.Rows(i).Cells("ID").Value) & "", OleConnenection)
                                                                                FtpDataset = New DataSet
                                                                                FtpAdaptater.Fill(FtpDataset)
                                                                                Ftpdatatable = FtpDataset.Tables(0)
                                                                                If Ftpdatatable.Rows.Count <> 0 Then
                                                                                    If DataListeIntegrer.Rows(i).Cells("Deplace").Value = True Then
                                                                                        If exceptionTrouve = True Then
                                                                                            'Il y'a une erreur au niveau de l'importation des articles à partir du fichier, on deplace pas le fichier
                                                                                            File.Delete(DataListeIntegrer.Rows(i).Cells("CheminExport").Value)
                                                                                            Label5.Refresh()
                                                                                            Label5.Text = "Modification Terminée!"
                                                                                        Else
                                                                                            'Deplacement du fichier vers les repertoire de sauvegarde
                                                                                            effaceFichier("FTP://" & RetourneServeurFtp(Ftpdatatable.Rows(0).Item("CheminFilexport")) & "/" & RetourneDirectoryFtp(Ftpdatatable.Rows(0).Item("CheminFilexport")) & "/" & System.IO.Path.GetFileName(DataListeIntegrer.Rows(i).Cells("CheminExport").Value), RetourneUserFtp(Ftpdatatable.Rows(0).Item("CheminFilexport")), RetournePassWordFtp(Ftpdatatable.Rows(0).Item("CheminFilexport")), ErreurJrn)
                                                                                            File.Move(DataListeIntegrer.Rows(i).Cells("CheminExport").Value, PathsfileSave & "" & Strings.Right(DateAndTime.Year(Now), 2) & "" & Format(DateAndTime.Month(Now), "00") & "" & Format(DateAndTime.Day(Now), "00") & "[" & "" & Format(DateAndTime.Hour(Now), "00") & "-" & Format(DateAndTime.Minute(Now), "00") & "-" & Format(DateAndTime.Second(Now), "00") & "]" & NomFichier)
                                                                                            Label5.Refresh()
                                                                                            Label5.Text = "Modification Terminée! Suppression des Fichiers exécutée..."
                                                                                        End If
                                                                                    Else
                                                                                        File.Delete(DataListeIntegrer.Rows(i).Cells("CheminExport").Value)
                                                                                        Label5.Refresh()
                                                                                        Label5.Text = "Modification Terminée!"
                                                                                    End If
                                                                                End If
                                                                            Else
                                                                                If DataListeIntegrer.Rows(i).Cells("Deplace").Value = True Then
                                                                                    If exceptionTrouve = True Then
                                                                                        'Il y'a une erreur au niveau de l'importation des articles à partir du fichier, on deplace pas le fichier
                                                                                        Label5.Refresh()
                                                                                        Label5.Text = "Modification Terminée!"
                                                                                    Else
                                                                                        'Deplacement du fichier vers les repertoire de sauvegarde
                                                                                        File.Move(DataListeIntegrer.Rows(i).Cells("CheminExport").Value, PathsfileSave & "" & Strings.Right(DateAndTime.Year(Now), 2) & "" & Format(DateAndTime.Month(Now), "00") & "" & Format(DateAndTime.Day(Now), "00") & "[" & "" & Format(DateAndTime.Hour(Now), "00") & "-" & Format(DateAndTime.Minute(Now), "00") & "-" & Format(DateAndTime.Second(Now), "00") & "]" & NomFichier)
                                                                                        Label5.Refresh()
                                                                                        Label5.Text = "Modification Terminée! Suppression des Fichiers exécutée..."
                                                                                    End If
                                                                                Else
                                                                                    Label5.Refresh()
                                                                                    Label5.Text = "Modification Terminée!"
                                                                                End If
                                                                            End If
                                                                        End If
                                                                        If OleExcelConnected.State = ConnectionState.Open Then
                                                                            OleExcelConnected.Close()
                                                                        End If
                                                                        DataListeIntegrer.Rows(i).Cells("Valider").Value = False
                                                                    Else
                                                                        ErreurJrn.WriteLine("Connexion à la Société - Base Commerciale :" & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & " -Base Comptable :" & Trim(DataListeIntegrer.Rows(i).Cells("Comptable").Value) & " Echec")
                                                                        Label5.Text = "Connexion à la Société " & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & " : Echec"
                                                                    End If
                                                                Else
                                                                    ErreurJrn.WriteLine("Echec de Connexion à SQL de base de données :" & Strings.Left(Dataname(UBound(Dataname)), InStr(Dataname(UBound(Dataname)), ".") - 1) & " Serveur : " & LireChaine(Trim(Cptadatatable.Rows(0).Item("Chemin1")), "CBASE", "ServeurSQL"))
                                                                    Label5.Text = "Echec de Connexion SQL à la Société " & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & " : Echec de traitement"
                                                                End If
                                                            Else
                                                                ErreurJrn.WriteLine("Aucune Base Comptable Correspondant à : " & Trim(DataListeIntegrer.Rows(i).Cells("Comptable").Value) & " Echec de traitement")
                                                            End If
                                                        Else
                                                            ErreurJrn.WriteLine("Aucune Base Commerciale Correspondant à : " & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & " Echec de traitement")
                                                        End If

                                                    Else
                                                        Label5.Text = "Echec de Connexion au fichier Excel :" & Trim(DataListeIntegrer.Rows(i).Cells("CheminExport").Value) & " : Echec de traitement"
                                                        ErreurJrn.WriteLine("Echec de Connexion au fichier Excel :" & Trim(DataListeIntegrer.Rows(i).Cells("CheminExport").Value) & " : Echec de traitement")
                                                    End If
                                                Else
                                                    Label5.Text = "Aucune Feuille Excel paramétrée pour le fichier :" & Trim(DataListeIntegrer.Rows(i).Cells("CheminExport").Value) & " : Echec de traitement"
                                                    ErreurJrn.WriteLine("Aucune Feuille Excel paramétrée pour le fichier :" & Trim(DataListeIntegrer.Rows(i).Cells("CheminExport").Value) & " : Echec de traitement")
                                                End If
                                            Else
                                                If Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value)) = "Longueur Fixe" Then
                                                    ArtAdaptater = New OleDbDataAdapter("select * from PARAMETRE  Where  Societe='" & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & "' and nomtype='COMMERCIAL'", OleConnenection)
                                                    ArtDataset = New DataSet
                                                    ArtAdaptater.Fill(ArtDataset)
                                                    Artdatatable = ArtDataset.Tables(0)
                                                    If Artdatatable.Rows.Count <> 0 Then
                                                        CptaAdaptater = New OleDbDataAdapter("select * from PARAMETRE  Where  Societe='" & Trim(DataListeIntegrer.Rows(i).Cells("Comptable").Value) & "' and nomtype='COMPTABILITE'", OleConnenection)
                                                        CptaDataset = New DataSet
                                                        CptaAdaptater.Fill(CptaDataset)
                                                        Cptadatatable = CptaDataset.Tables(0)
                                                        If Cptadatatable.Rows.Count <> 0 Then
                                                            Dim Dataname() As String = Split(Trim(Cptadatatable.Rows(0).Item("Chemin1")), "\")
                                                            If SocieteConnected(Strings.Left(Dataname(UBound(Dataname)), InStr(Dataname(UBound(Dataname)), ".") - 1), Trim(Cptadatatable.Rows(0).Item("MotPas").ToString), Trim(Cptadatatable.Rows(0).Item("NomUser")), LireChaine(Trim(Cptadatatable.Rows(0).Item("Chemin1")), "CBASE", "ServeurSQL")) = True Then
                                                                FermeBaseCial(BaseCial)
                                                                If OuvreBaseCial(BaseCial, BaseCpta, Trim(Artdatatable.Rows(0).Item("Chemin1")), Trim(Cptadatatable.Rows(0).Item("Chemin1")), Trim(Artdatatable.Rows(0).Item("UserSage")), Trim(Artdatatable.Rows(0).Item("PasseSage").ToString), Trim(Cptadatatable.Rows(0).Item("UserSage")), Trim(Cptadatatable.Rows(0).Item("PasseSage").ToString)) = True Then
                                                                    ErreurJrn.WriteLine("Connexion à la Société " & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & " Reussie")
                                                                    Label5.Refresh()
                                                                    NomFichier = Trim(DataListeIntegrer.Rows(i).Cells("CheminExport").Value)
                                                                    Do While InStr(Trim(NomFichier), "\") <> 0
                                                                        NomFichier = Strings.Right(NomFichier, Strings.Len(Trim(NomFichier)) - InStr(Trim(NomFichier), "\"))
                                                                    Loop
                                                                    ErreurJrn.WriteLine("")
                                                                    ErreurJrn.WriteLine("Début de traitement du fichier : " & NomFichier)
                                                                    ErreurJrn.WriteLine("")
                                                                    Label5.Refresh()
                                                                    If Modification_Verification_Fichier(DataListeIntegrer.Rows(i).Cells("CheminExport").Value, DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value)), "", "", LirePieceModification(DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value))), RenvoieFormatDate(DataListeIntegrer.Rows(i).Cells("Chemin").Value)) = True Then
                                                                        Label5.Text = "Modification En Cours..."
                                                                        Modification_Integration_Fichier(DataListeIntegrer.Rows(i).Cells("CheminExport").Value, DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value)), "", "", LirePieceModification(DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value))), RenvoieFormatDate(DataListeIntegrer.Rows(i).Cells("Chemin").Value))
                                                                        RecuperationEnregistrementModifié(DataListeIntegrer.Rows(i).Cells("CheminExport").Value, DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value)), "", "", ListePiece, LirePieceModification(DataListeIntegrer.Rows(i).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(i).Cells("TypeFormat").Value))))
                                                                        If DataListeIntegrer.Rows(i).Cells("Cible").Value = "FTP" Then
                                                                            Dim FtpAdaptater As OleDbDataAdapter
                                                                            Dim FtpDataset As DataSet
                                                                            Dim Ftpdatatable As DataTable
                                                                            FtpAdaptater = New OleDbDataAdapter("select * from  WIT_SCHEMA WHERE Cible='FTP' And IDDossier=" & CInt(DataListeIntegrer.Rows(i).Cells("ID").Value) & "", OleConnenection)
                                                                            FtpDataset = New DataSet
                                                                            FtpAdaptater.Fill(FtpDataset)
                                                                            Ftpdatatable = FtpDataset.Tables(0)
                                                                            If Ftpdatatable.Rows.Count <> 0 Then
                                                                                If DataListeIntegrer.Rows(i).Cells("Deplace").Value = True Then
                                                                                    If exceptionTrouve = True Then
                                                                                        'Il y'a une erreur au niveau de l'importation des articles à partir du fichier, on deplace pas le fichier
                                                                                        File.Delete(DataListeIntegrer.Rows(i).Cells("CheminExport").Value)
                                                                                        Label5.Refresh()
                                                                                        Label5.Text = "Modification Terminée!"
                                                                                    Else
                                                                                        'Deplacement du fichier vers les repertoire de sauvegarde
                                                                                        effaceFichier("FTP://" & RetourneServeurFtp(Ftpdatatable.Rows(0).Item("CheminFilexport")) & "/" & RetourneDirectoryFtp(Ftpdatatable.Rows(0).Item("CheminFilexport")) & "/" & System.IO.Path.GetFileName(DataListeIntegrer.Rows(i).Cells("CheminExport").Value), RetourneUserFtp(Ftpdatatable.Rows(0).Item("CheminFilexport")), RetournePassWordFtp(Ftpdatatable.Rows(0).Item("CheminFilexport")), ErreurJrn)
                                                                                        File.Move(DataListeIntegrer.Rows(i).Cells("CheminExport").Value, PathsfileSave & "" & Strings.Right(DateAndTime.Year(Now), 2) & "" & Format(DateAndTime.Month(Now), "00") & "" & Format(DateAndTime.Day(Now), "00") & "[" & "" & Format(DateAndTime.Hour(Now), "00") & "-" & Format(DateAndTime.Minute(Now), "00") & "-" & Format(DateAndTime.Second(Now), "00") & "]" & NomFichier)
                                                                                        Label5.Refresh()
                                                                                        Label5.Text = "Modification Terminée! Suppression des Fichiers exécutée..."
                                                                                    End If
                                                                                Else
                                                                                    File.Delete(DataListeIntegrer.Rows(i).Cells("CheminExport").Value)
                                                                                    Label5.Refresh()
                                                                                    Label5.Text = "Modification Terminée!"
                                                                                End If
                                                                            End If
                                                                        Else
                                                                            If DataListeIntegrer.Rows(i).Cells("Deplace").Value = True Then
                                                                                If exceptionTrouve = True Then
                                                                                    'Il y'a une erreur au niveau de l'importation des articles à partir du fichier, on deplace pas le fichier
                                                                                    Label5.Refresh()
                                                                                    Label5.Text = "Modification Terminée!"
                                                                                Else
                                                                                    'Deplacement du fichier vers les repertoire de sauvegarde
                                                                                    File.Move(DataListeIntegrer.Rows(i).Cells("CheminExport").Value, PathsfileSave & "" & Strings.Right(DateAndTime.Year(Now), 2) & "" & Format(DateAndTime.Month(Now), "00") & "" & Format(DateAndTime.Day(Now), "00") & "[" & "" & Format(DateAndTime.Hour(Now), "00") & "-" & Format(DateAndTime.Minute(Now), "00") & "-" & Format(DateAndTime.Second(Now), "00") & "]" & NomFichier)
                                                                                    Label5.Refresh()
                                                                                    Label5.Text = "Modification Terminée! Suppression des Fichiers exécutée..."
                                                                                End If
                                                                            Else
                                                                                Label5.Refresh()
                                                                                Label5.Text = "Modification Terminée!"
                                                                            End If
                                                                        End If
                                                                    End If
                                                                    DataListeIntegrer.Rows(i).Cells("Valider").Value = False
                                                                Else
                                                                    ErreurJrn.WriteLine("Connexion à la Société - Base Commerciale :" & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & " -Base Comptable :" & Trim(DataListeIntegrer.Rows(i).Cells("Comptable").Value) & " Echec")
                                                                    Label5.Text = "Connexion à la Société " & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & " : Echec"
                                                                End If
                                                            Else
                                                                ErreurJrn.WriteLine("Echec de Connexion à SQL de base de données :" & Strings.Left(Dataname(UBound(Dataname)), InStr(Dataname(UBound(Dataname)), ".") - 1) & " Serveur : " & LireChaine(Trim(Cptadatatable.Rows(0).Item("Chemin1")), "CBASE", "ServeurSQL"))
                                                                Label5.Text = "Echec de Connexion SQL à la Société " & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & " : Echec de traitement"
                                                            End If
                                                        Else
                                                            ErreurJrn.WriteLine("Aucune Base Comptable Correspondant à : " & Trim(DataListeIntegrer.Rows(i).Cells("Comptable").Value) & " Echec de traitement")
                                                        End If
                                                    Else
                                                        ErreurJrn.WriteLine("Aucune Base Commerciale Correspondant à : " & Trim(DataListeIntegrer.Rows(i).Cells("Commercial").Value) & " Echec de traitement")
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Else
                                    ErreurJrn.WriteLine("Chemin du fichier inexistant : " & DataListeIntegrer.Rows(i).Cells("Chemin").Value)
                                End If
                            End If
                        End If
                    End If
                Next i
                ErreurJrn.Close()
            End If
            AfficheSchemasIntegrer()
            Affichagefichier()
            encours.Close()
            FermeBaseCial(BaseCial)
        Catch ex As Exception
            encours.Close()
            exceptionTrouve = True
            If IsNothing(ErreurJrn) = False Then
                ErreurJrn.WriteLine("Erreur système :" & ex.Message)
                ErreurJrn.Close()
            End If
        End Try
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub DataListeIntegrer_CellMouseLeave(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataListeIntegrer.CellMouseLeave
        Dim i As Integer
        Try
            If DataListeIntegrer.Columns(e.ColumnIndex).Name = "Valider" Then
                For i = 0 To DataListeIntegrer.RowCount - 1
                    If DataListeIntegrer.Rows(i).Cells("Valider").Value = True Then
                        IndexPrec = i
                        i = DataListeIntegrer.RowCount - 1
                    End If

                Next i
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub DataListeIntegrer_CellValidated(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataListeIntegrer.CellValidated
        Dim i As Integer
        Try
            If DataListeIntegrer.Columns(e.ColumnIndex).Name = "Valider" Then
                For i = 0 To DataListeIntegrer.RowCount - 1
                    If DataListeIntegrer.Rows(i).Cells("Valider").Value = True Then
                        IndexPrec = i
                        i = DataListeIntegrer.RowCount - 1
                    End If

                Next i
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub DataListeIntegrer_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataListeIntegrer.CellValueChanged
        Dim i As Integer
        Try
            If DataListeIntegrer.Columns(e.ColumnIndex).Name = "Valider" Then
                For i = 0 To DataListeIntegrer.RowCount - 1
                    If DataListeIntegrer.Rows(i).Cells("Valider").Value = True Then
                        IndexPrec = i
                        i = DataListeIntegrer.RowCount - 1
                    End If

                Next i
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub BT_Apercue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Apercue.Click
        Try
            If IndexPrec >= 0 Then
                If Renvoietypeformat(Trim(DataListeIntegrer.Rows(IndexPrec).Cells("TypeFormat").Value)) = "Délimité" Or Renvoietypeformat(Trim(DataListeIntegrer.Rows(IndexPrec).Cells("TypeFormat").Value)) = "Tabulation" Or Renvoietypeformat(Trim(DataListeIntegrer.Rows(IndexPrec).Cells("TypeFormat").Value)) = "Pipe" Then
                    If Renvoietypeformat(Trim(DataListeIntegrer.Rows(IndexPrec).Cells("TypeFormat").Value)) = "Tabulation" Then
                        sColumnsSepar = ControlChars.Tab
                    Else
                        If Renvoietypeformat(Trim(DataListeIntegrer.Rows(IndexPrec).Cells("TypeFormat").Value)) = "Pipe" Then
                            sColumnsSepar = "|"
                        Else
                            sColumnsSepar = ";"
                        End If
                    End If
                    If DataListeIntegrer.Rows(IndexPrec).Cells("Valider").Value = True Then
                        If File.Exists(DataListeIntegrer.Rows(IndexPrec).Cells("Chemin").Value) = True And File.Exists(DataListeIntegrer.Rows(IndexPrec).Cells("CheminExport").Value) Then
                            Lecture_Suivant_DuFichierExcel(DataListeIntegrer.Rows(IndexPrec).Cells("CheminExport").Value, DataListeIntegrer.Rows(IndexPrec).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(IndexPrec).Cells("TypeFormat").Value)), "", sColumnsSepar)
                        End If
                    End If
                Else
                    If Renvoietypeformat(Trim(DataListeIntegrer.Rows(IndexPrec).Cells("TypeFormat").Value)) = "Excel" Then
                        If Trim(DataListeIntegrer.Rows(IndexPrec).Cells("FeuilleExcel").Value) <> "" Then
                            If DataListeIntegrer.Rows(IndexPrec).Cells("Valider").Value = True Then
                                If File.Exists(DataListeIntegrer.Rows(IndexPrec).Cells("Chemin").Value) = True And File.Exists(DataListeIntegrer.Rows(IndexPrec).Cells("CheminExport").Value) Then
                                    If OleExcelConnected.State = ConnectionState.Open Then
                                        OleExcelConnected.Close()
                                    End If
                                    Lecture_Suivant_DuFichierExcel(DataListeIntegrer.Rows(IndexPrec).Cells("CheminExport").Value, DataListeIntegrer.Rows(IndexPrec).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(IndexPrec).Cells("TypeFormat").Value)), Trim(DataListeIntegrer.Rows(IndexPrec).Cells("FeuilleExcel").Value), "")
                                End If
                            End If
                        Else
                            MsgBox("La feuille Excel n'est pas renseignée", MsgBoxStyle.Information, "Import des Fichiers")
                        End If
                    Else
                        If Renvoietypeformat(Trim(DataListeIntegrer.Rows(IndexPrec).Cells("TypeFormat").Value)) = "Longueur Fixe" Then
                            If DataListeIntegrer.Rows(IndexPrec).Cells("Valider").Value = True Then
                                If File.Exists(DataListeIntegrer.Rows(IndexPrec).Cells("Chemin").Value) = True And File.Exists(DataListeIntegrer.Rows(IndexPrec).Cells("CheminExport").Value) Then
                                    Lecture_Suivant_DuFichierExcel(DataListeIntegrer.Rows(IndexPrec).Cells("CheminExport").Value, DataListeIntegrer.Rows(IndexPrec).Cells("Chemin").Value, Renvoietypeformat(Trim(DataListeIntegrer.Rows(IndexPrec).Cells("TypeFormat").Value)), "", "")
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception

        End Try
        If OleExcelConnected.State = ConnectionState.Open Then
            OleExcelConnected.Close()
        End If
    End Sub    
    Private Function ConvertionSQLDate(ByRef Valeur As Object, ByRef DateFormat As String) As String
        If DateFormat = "aa-mm-jj" Then
            If Strings.Len(Trim(Valeur)) = 8 Then
                If IsDate(Strings.Mid(Trim(Valeur), 7, 2) & "-" & Strings.Mid(Trim(Valeur), 4, 2) & "-" & Strings.Mid(Trim(Valeur), 1, 2)) = True Then
                    Valeur = "CONVERT(DATETIME, '" & Format(CDate(Strings.Mid(Trim(Valeur), 7, 2) & "-" & Strings.Mid(Trim(Valeur), 4, 2) & "-" & Strings.Mid(Trim(Valeur), 1, 2)), "yyyy/MM/dd") & "', 102)"
                End If
            End If
        Else
            If DateFormat = "aaaa-mm-jj" Then
                If Strings.Len(Trim(Valeur)) = 10 Then
                    If IsDate(Strings.Mid(Trim(Valeur), 9, 2) & "-" & Strings.Mid(Trim(Valeur), 6, 2) & "-" & Strings.Mid(Trim(Valeur), 1, 4)) = True Then
                        Valeur = "CONVERT(DATETIME, '" & Format(CDate(Strings.Mid(Trim(Valeur), 9, 2) & "-" & Strings.Mid(Trim(Valeur), 6, 2) & "-" & Strings.Mid(Trim(Valeur), 1, 4)), "yyyy/MM/dd") & "', 102)"
                    End If
                End If
            Else
                If DateFormat = "jj-mm-aa" Then
                    If Strings.Len(Trim(Valeur)) = 8 Then
                        If IsDate(Strings.Mid(Trim(Valeur), 1, 2) & "-" & Strings.Mid(Trim(Valeur), 4, 2) & "-" & Strings.Mid(Trim(Valeur), 7, 2)) = True Then
                            Valeur = "CONVERT(DATETIME, '" & Format(CDate(Strings.Mid(Trim(Valeur), 1, 2) & "-" & Strings.Mid(Trim(Valeur), 4, 2) & "-" & Strings.Mid(Trim(Valeur), 7, 2)), "yyyy/MM/dd") & "', 102)"
                        End If
                    End If
                Else
                    If DateFormat = "jj-mm-aaaa" Then
                        If Strings.Len(Trim(Valeur)) = 10 Then
                            If IsDate(Strings.Mid(Trim(Valeur), 1, 2) & "-" & Strings.Mid(Trim(Valeur), 4, 2) & "-" & Strings.Mid(Trim(Valeur), 7, 4)) = True Then
                                Valeur = "CONVERT(DATETIME, '" & Format(CDate(Strings.Mid(Trim(Valeur), 1, 2) & "-" & Strings.Mid(Trim(Valeur), 4, 2) & "-" & Strings.Mid(Trim(Valeur), 7, 4)), "yyyy/MM/dd") & "', 102)"
                            End If
                        End If
                    Else
                        If DateFormat = "jjmmaa" Then
                            If Strings.Len(Trim(Valeur)) = 6 Then
                                If IsDate(Strings.Mid(Trim(Valeur), 1, 2) & "/" & Strings.Mid(Trim(Valeur), 3, 2) & "/" & Strings.Mid(Trim(Valeur), 5, 2)) = True Then
                                    Valeur = "CONVERT(DATETIME, '" & Format(CDate(Strings.Mid(Trim(Valeur), 1, 2) & "/" & Strings.Mid(Trim(Valeur), 3, 2) & "/" & Strings.Mid(Trim(Valeur), 5, 2)), "yyyy/MM/dd") & "', 102)"
                                End If
                            Else
                                If Strings.Len(Trim(Valeur)) = 5 Then
                                    If IsNumeric(Trim(Valeur)) = True Then
                                        If IsDate(Strings.Mid(Format(CDbl(Trim(Valeur)), "000000"), 1, 2) & "/" & Strings.Mid(Format(CDbl(Trim(Valeur)), "000000"), 3, 2) & "/" & Strings.Mid(Format(CDbl(Trim(Valeur)), "000000"), 5, 2)) = True Then
                                            Valeur = "CONVERT(DATETIME, '" & Format(CDate(Strings.Mid(Format(CDbl(Trim(Valeur)), "000000"), 1, 2) & "/" & Strings.Mid(Format(CDbl(Trim(Valeur)), "000000"), 3, 2) & "/" & Strings.Mid(Format(CDbl(Trim(Valeur)), "000000"), 5, 2)), "yyyy/MM/dd") & "', 102)"
                                        End If
                                    End If
                                End If
                            End If
                        Else
                            If DateFormat = "jjmmaaaa" Then
                                If Strings.Len(Trim(Valeur)) = 8 Then
                                    If IsDate(Strings.Mid(Trim(Valeur), 1, 2) & "/" & Strings.Mid(Trim(Valeur), 3, 2) & "/" & Strings.Mid(Trim(Valeur), 5, 4)) = True Then
                                        Valeur = "CONVERT(DATETIME, '" & Format(CDate(Strings.Mid(Trim(Valeur), 1, 2) & "/" & Strings.Mid(Trim(Valeur), 3, 2) & "/" & Strings.Mid(Trim(Valeur), 5, 4)), "yyyy/MM/dd") & "', 102)"
                                    End If
                                Else
                                    If Strings.Len(Trim(Valeur)) = 7 Then
                                        If IsNumeric(Trim(Valeur)) = True Then
                                            If IsDate(Strings.Mid(Format(CDbl(Trim(Valeur)), "00000000"), 1, 2) & "/" & Strings.Mid(Format(CDbl(Trim(Valeur)), "00000000"), 3, 2) & "/" & Strings.Mid(Format(CDbl(Trim(Valeur)), "00000000"), 5, 4)) = True Then
                                                Valeur = "CONVERT(DATETIME, '" & Format(CDate(Strings.Mid(Format(CDbl(Trim(Valeur)), "00000000"), 1, 2) & "/" & Strings.Mid(Format(CDbl(Trim(Valeur)), "00000000"), 3, 2) & "/" & Strings.Mid(Format(CDbl(Trim(Valeur)), "00000000"), 5, 4)), "yyyy/MM/dd") & "', 102)"
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                If DateFormat = "aammjj" Then
                                    If Strings.Len(Trim(Valeur)) = 6 Then
                                        If IsDate(Strings.Mid(Trim(Valeur), 5, 2) & "/" & Strings.Mid(Trim(Valeur), 3, 2) & "/" & Strings.Mid(Trim(Valeur), 1, 2)) = True Then
                                            Valeur = "CONVERT(DATETIME, '" & Format(CDate(Strings.Mid(Trim(Valeur), 5, 2) & "/" & Strings.Mid(Trim(Valeur), 3, 2) & "/" & Strings.Mid(Trim(Valeur), 1, 2)), "yyyy/MM/dd") & "', 102)"
                                        End If
                                    Else
                                        If Strings.Len(Trim(Valeur)) = 5 Then
                                            If IsNumeric(Trim(Valeur)) = True Then
                                                If IsDate(Strings.Mid(Format(CDbl(Trim(Valeur)), "000000"), 5, 2) & "/" & Strings.Mid(Format(CDbl(Trim(Valeur)), "000000"), 3, 2) & "/" & Strings.Mid(Format(CDbl(Trim(Valeur)), "000000"), 1, 2)) = True Then
                                                    Valeur = "CONVERT(DATETIME, '" & Format(CDate(Strings.Mid(Format(CDbl(Trim(Valeur)), "000000"), 5, 2) & "/" & Strings.Mid(Format(CDbl(Trim(Valeur)), "000000"), 3, 2) & "/" & Strings.Mid(Format(CDbl(Trim(Valeur)), "000000"), 1, 2)), "yyyy/MM/dd") & "', 102)"
                                                End If
                                            End If
                                        End If
                                    End If
                                Else
                                    If DateFormat = "aaaammjj" Then
                                        If Strings.Len(Trim(Valeur)) = 8 Then
                                            If IsDate(Strings.Mid(Trim(Valeur), 7, 2) & "/" & Strings.Mid(Trim(Valeur), 5, 2) & "/" & Strings.Mid(Trim(Valeur), 1, 4)) = True Then
                                                Valeur = "CONVERT(DATETIME, '" & Format(CDate(Strings.Mid(Trim(Valeur), 7, 2) & "/" & Strings.Mid(Trim(Valeur), 5, 2) & "/" & Strings.Mid(Trim(Valeur), 1, 4)), "yyyy/MM/dd") & "', 102)"
                                            End If
                                        Else
                                            If Strings.Len(Trim(Valeur)) = 7 Then
                                                If IsNumeric(Trim(Valeur)) = True Then
                                                    If IsDate(Strings.Mid(Format(CDbl(Trim(Valeur)), "00000000"), 7, 2) & "/" & Strings.Mid(Format(CDbl(Trim(Valeur)), "00000000"), 5, 2) & "/" & Strings.Mid(Format(CDbl(Trim(Valeur)), "00000000"), 1, 4)) = True Then
                                                        Valeur = "CONVERT(DATETIME, '" & Format(CDate(Strings.Mid(Format(CDbl(Trim(Valeur)), "00000000"), 7, 2) & "/" & Strings.Mid(Format(CDbl(Trim(Valeur)), "00000000"), 5, 2) & "/" & Strings.Mid(Format(CDbl(Trim(Valeur)), "00000000"), 1, 4)), "yyyy/MM/dd") & "', 102)"
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If DateFormat = "jj/mm/aa" Then
                                            If Strings.Len(Trim(Valeur)) = 8 Then
                                                If IsDate(Strings.Mid(Trim(Valeur), 1, 2) & "/" & Strings.Mid(Trim(Valeur), 4, 2) & "/" & Strings.Mid(Trim(Valeur), 7, 2)) = True Then
                                                    Valeur = "CONVERT(DATETIME, '" & Format(CDate(Strings.Mid(Trim(Valeur), 1, 2) & "/" & Strings.Mid(Trim(Valeur), 4, 2) & "/" & Strings.Mid(Trim(Valeur), 7, 2)), "yyyy/MM/dd") & "', 102)"
                                                End If
                                            End If
                                        Else
                                            If DateFormat = "jj/mm/aaaa" Then
                                                If Strings.Len(Trim(Valeur)) = 10 Then
                                                    If IsDate(Strings.Mid(Trim(Valeur), 1, 2) & "/" & Strings.Mid(Trim(Valeur), 4, 2) & "/" & Strings.Mid(Trim(Valeur), 7, 4)) = True Then
                                                        Valeur = "CONVERT(DATETIME, '" & Format(CDate(Strings.Mid(Trim(Valeur), 1, 2) & "/" & Strings.Mid(Trim(Valeur), 4, 2) & "/" & Strings.Mid(Trim(Valeur), 7, 4)), "yyyy/MM/dd") & "', 102)"
                                                    End If
                                                End If
                                            Else
                                                If IsDate(Trim(Valeur)) = True Then
                                                    Valeur = "CONVERT(DATETIME, '" & Format(CDate(Valeur), "yyyy/MM/dd") & "', 102)"
                                                End If
                                            End If

                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        ConvertionSQLDate = Valeur
    End Function
    Private Function RenvoieDateValide(ByRef ValeurDate As Object, ByRef DateFormat As Object) As Date
        'Hermann
        Try
            Select Case DateFormat
                Case "mm/jj/aaaa"
                    If Strings.Len(Trim(ValeurDate)) = 10 Then
                        If IsDate(Strings.Mid(Trim(ValeurDate), 4, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 2) & "/" & Strings.Mid(Trim(ValeurDate), 7, 4)) = True Then
                            ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 4, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 2) & "/" & Strings.Mid(Trim(ValeurDate), 7, 4))
                        End If
                    End If
                    Exit Select
                Case "jj/mm/aa"
                    If Strings.Len(Trim(ValeurDate)) = 8 Then
                        If IsDate(Strings.Mid(Trim(ValeurDate), 1, 2) & "/" & Strings.Mid(Trim(ValeurDate), 4, 2) & "/" & Strings.Mid(Trim(ValeurDate), 7, 2)) = True Then
                            ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 1, 2) & "-" & Strings.Mid(Trim(ValeurDate), 4, 2) & "-" & Strings.Mid(Trim(ValeurDate), 7, 2))
                        End If
                    End If
                    Exit Select
                Case "aammjj"
                    If Strings.Len(Trim(ValeurDate)) = 6 Then
                        If IsDate(Strings.Mid(Trim(ValeurDate), 5, 2) & "/" & Strings.Mid(Trim(ValeurDate), 3, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 2)) = True Then
                            ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 5, 2) & "/" & Strings.Mid(Trim(ValeurDate), 3, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 2))
                        End If
                    Else
                        If Strings.Len(Trim(ValeurDate)) = 5 Then
                            If IsNumeric(Trim(ValeurDate)) = True Then
                                If IsDate(Strings.Mid(Trim(ValeurDate), 5, 1) & "/" & Strings.Mid(Trim(ValeurDate), 3, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 2)) = True Then
                                    ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 5, 1) & "/" & Strings.Mid(Trim(ValeurDate), 3, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 2))
                                End If
                            End If
                        End If
                    End If
                    Exit Select
                Case "jjmmaa"
                    If Strings.Len(Trim(ValeurDate)) = 6 Then
                        If IsDate(Strings.Mid(Trim(ValeurDate), 1, 2) & "/" & Strings.Mid(Trim(ValeurDate), 3, 2) & "/" & Strings.Mid(Trim(ValeurDate), 5, 2)) = True Then
                            ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 1, 2) & "/" & Strings.Mid(Trim(ValeurDate), 3, 2) & "/" & Strings.Mid(Trim(ValeurDate), 5, 2))
                        End If
                    Else
                        If Strings.Len(Trim(ValeurDate)) = 5 Then
                            If IsNumeric(Trim(ValeurDate)) = True Then
                                If IsDate(Strings.Mid(Trim(ValeurDate), 1, 1) & "/" & Strings.Mid(Trim(ValeurDate), 2, 2) & "/" & Strings.Mid(Trim(ValeurDate), 4, 2)) = True Then
                                    ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 1, 1) & "/" & Strings.Mid(Trim(ValeurDate), 2, 2) & "/" & Strings.Mid(Trim(ValeurDate), 4, 2))
                                End If
                            End If
                        End If
                    End If
                    Exit Select
                Case "aaaammjj"
                    If Strings.Len(Trim(ValeurDate)) = 8 Then
                        If IsDate(Strings.Mid(Trim(ValeurDate), 7, 2) & "/" & Strings.Mid(Trim(ValeurDate), 5, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 4)) = True Then
                            ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 7, 2) & "/" & Strings.Mid(Trim(ValeurDate), 5, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 4))
                        End If
                    Else
                        If Strings.Len(Trim(ValeurDate)) = 7 Then
                            If IsNumeric(Trim(ValeurDate)) = True Then
                                If IsDate(Strings.Mid(Trim(ValeurDate), 7, 1) & "/" & Strings.Mid(Trim(ValeurDate), 5, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 4)) = True Then
                                    ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 7, 1) & "/" & Strings.Mid(Trim(ValeurDate), 5, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 4))
                                End If
                            End If
                        End If
                    End If
                    Exit Select
                Case "jjmmaaaa"
                    If Strings.Len(Trim(ValeurDate)) = 8 Then
                        If IsDate(Strings.Mid(Trim(ValeurDate), 1, 2) & "/" & Strings.Mid(Trim(ValeurDate), 3, 2) & "/" & Strings.Mid(Trim(ValeurDate), 5, 4)) = True Then
                            ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 1, 2) & "/" & Strings.Mid(Trim(ValeurDate), 3, 2) & "/" & Strings.Mid(Trim(ValeurDate), 5, 4))
                        End If
                    Else
                        If Strings.Len(Trim(ValeurDate)) = 7 Then
                            If IsNumeric(Trim(ValeurDate)) = True Then
                                If IsDate(Strings.Mid(Trim(ValeurDate), 1, 1) & "/" & Strings.Mid(Trim(ValeurDate), 2, 2) & "/" & Strings.Mid(Trim(ValeurDate), 4, 4)) = True Then
                                    ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 1, 1) & "/" & Strings.Mid(Trim(ValeurDate), 2, 2) & "/" & Strings.Mid(Trim(ValeurDate), 4, 4))
                                End If
                            End If
                        End If
                    End If
                    Exit Select
                Case "aa-mm-jj"
                    If Strings.Len(Trim(ValeurDate)) = 8 Then
                        If IsDate(Strings.Mid(Trim(ValeurDate), 7, 2) & "-" & Strings.Mid(Trim(ValeurDate), 4, 2) & "-" & Strings.Mid(Trim(ValeurDate), 1, 2)) = True Then
                            ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 7, 2) & "-" & Strings.Mid(Trim(ValeurDate), 4, 2) & "-" & Strings.Mid(Trim(ValeurDate), 1, 2))
                        End If
                    End If
                    Exit Select
                Case "jj-mm-aa"
                    If Strings.Len(Trim(ValeurDate)) = 8 Then
                        If IsDate(Strings.Mid(Trim(ValeurDate), 1, 2) & "-" & Strings.Mid(Trim(ValeurDate), 4, 2) & "-" & Strings.Mid(Trim(ValeurDate), 7, 2)) = True Then
                            ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 1, 2) & "-" & Strings.Mid(Trim(ValeurDate), 4, 2) & "-" & Strings.Mid(Trim(ValeurDate), 7, 2))
                        End If
                    End If
                    Exit Select
                Case "aaaa-mm-jj"
                    If Strings.Len(Trim(ValeurDate)) = 10 Then
                        If IsDate(Strings.Mid(Trim(ValeurDate), 9, 2) & "-" & Strings.Mid(Trim(ValeurDate), 6, 2) & "-" & Strings.Mid(Trim(ValeurDate), 1, 4)) = True Then
                            ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 9, 2) & "-" & Strings.Mid(Trim(ValeurDate), 6, 2) & "-" & Strings.Mid(Trim(ValeurDate), 1, 4))
                        End If
                    End If
                    Exit Select
                Case "jj-mm-aaaa"
                    If Strings.Len(Trim(ValeurDate)) = 10 Then
                        If IsDate(Strings.Mid(Trim(ValeurDate), 1, 2) & "-" & Strings.Mid(Trim(ValeurDate), 4, 2) & "-" & Strings.Mid(Trim(ValeurDate), 7, 4)) = True Then
                            ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 1, 2) & "-" & Strings.Mid(Trim(ValeurDate), 4, 2) & "-" & Strings.Mid(Trim(ValeurDate), 7, 4))
                        End If
                    End If
                    Exit Select
                Case "mmaa"
                    If Strings.Len(Trim(ValeurDate)) = 6 Then
                        If IsDate(Strings.Mid(Trim(ValeurDate), 1, 2) & "/" & Strings.Mid(Trim(ValeurDate), 3, 2) & "/" & Strings.Mid(Trim(ValeurDate), 5, 2)) = True Then
                            ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 1, 2) & "/" & Strings.Mid(Trim(ValeurDate), 3, 2) & "/" & Strings.Mid(Trim(ValeurDate), 5, 2))
                        End If
                    Else
                        If Strings.Len(Trim(ValeurDate)) = 5 Then
                            If IsNumeric(Trim(ValeurDate)) = True Then
                                If IsDate(Strings.Mid(Trim(ValeurDate), 1, 1) & "/" & Strings.Mid(Trim(ValeurDate), 2, 2) & "/" & Strings.Mid(Trim(ValeurDate), 4, 2)) = True Then
                                    ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 1, 1) & "/" & Strings.Mid(Trim(ValeurDate), 2, 2) & "/" & Strings.Mid(Trim(ValeurDate), 4, 2))
                                End If
                            End If
                        End If
                    End If
                    Exit Select
                Case "aamm"
                    If Strings.Len(Trim(ValeurDate)) = 6 Then
                        If IsDate(Strings.Mid(Trim(ValeurDate), 5, 2) & "/" & Strings.Mid(Trim(ValeurDate), 3, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 2)) = True Then
                            ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 5, 2) & "/" & Strings.Mid(Trim(ValeurDate), 3, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 2))
                        End If
                    Else
                        If Strings.Len(Trim(ValeurDate)) = 5 Then
                            If IsNumeric(Trim(ValeurDate)) = True Then
                                If IsDate(Strings.Mid(Trim(ValeurDate), 5, 1) & "/" & Strings.Mid(Trim(ValeurDate), 3, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 2)) = True Then
                                    ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 5, 1) & "/" & Strings.Mid(Trim(ValeurDate), 3, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 2))
                                End If

                            End If
                        End If
                    End If
                    Exit Select
                Case "mmaaaa"
                    If Strings.Len(Trim(ValeurDate)) = 8 Then
                        If IsDate(Strings.Mid(Trim(ValeurDate), 1, 2) & "/" & Strings.Mid(Trim(ValeurDate), 3, 2) & "/" & Strings.Mid(Trim(ValeurDate), 5, 4)) = True Then
                            ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 1, 2) & "/" & Strings.Mid(Trim(ValeurDate), 3, 2) & "/" & Strings.Mid(Trim(ValeurDate), 5, 4))
                        End If
                    Else
                        If Strings.Len(Trim(ValeurDate)) = 7 Then
                            If IsNumeric(Trim(ValeurDate)) = True Then
                                If IsDate(Strings.Mid(Trim(ValeurDate), 1, 1) & "/" & Strings.Mid(Trim(ValeurDate), 2, 2) & "/" & Strings.Mid(Trim(ValeurDate), 4, 4)) = True Then
                                    ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 1, 1) & "/" & Strings.Mid(Trim(ValeurDate), 2, 2) & "/" & Strings.Mid(Trim(ValeurDate), 4, 4))
                                End If
                            End If
                        End If
                    End If
                    Exit Select
                Case "aaaamm"
                    If Strings.Len(Trim(ValeurDate)) = 8 Then
                        If IsDate(Strings.Mid(Trim(ValeurDate), 7, 2) & "/" & Strings.Mid(Trim(ValeurDate), 5, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 4)) = True Then
                            ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 7, 2) & "/" & Strings.Mid(Trim(ValeurDate), 5, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 4))
                        End If
                    Else
                        If Strings.Len(Trim(ValeurDate)) = 7 Then
                            If IsNumeric(Trim(ValeurDate)) = True Then
                                If IsDate(Strings.Mid(Trim(ValeurDate), 7, 1) & "/" & Strings.Mid(Trim(ValeurDate), 5, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 4)) = True Then
                                    ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 7, 1) & "/" & Strings.Mid(Trim(ValeurDate), 5, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 4))
                                End If
                            End If
                        End If
                    End If
                    Exit Select
                Case "mmjjaa"
                    If Strings.Len(Trim(ValeurDate)) = 6 Then
                        If IsDate(Strings.Mid(Trim(ValeurDate), 3, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 2) & "/" & Strings.Mid(Trim(ValeurDate), 5, 2)) = True Then
                            ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 3, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 2) & "/" & Strings.Mid(Trim(ValeurDate), 5, 2))
                        End If
                    Else
                        If Strings.Len(Trim(ValeurDate)) = 5 Then
                            If IsNumeric(Trim(ValeurDate)) = True Then
                                If IsDate(Strings.Mid(Trim(ValeurDate), 2, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 1) & "/" & Strings.Mid(Trim(ValeurDate), 4, 2)) = True Then
                                    ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 2, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 1) & "/" & Strings.Mid(Trim(ValeurDate), 4, 2))
                                End If
                            End If
                        End If
                    End If
                    Exit Select
                Case "mmjjaaaa"
                    If Strings.Len(Trim(ValeurDate)) = 8 Then
                        If IsDate(Strings.Mid(Trim(ValeurDate), 3, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 2) & "/" & Strings.Mid(Trim(ValeurDate), 5, 4)) = True Then
                            ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 3, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 2) & "/" & Strings.Mid(Trim(ValeurDate), 5, 4))
                        End If
                    Else
                        If Strings.Len(Trim(ValeurDate)) = 7 Then
                            If IsNumeric(Trim(ValeurDate)) = True Then
                                If IsDate(Strings.Mid(Trim(ValeurDate), 2, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 1) & "/" & Strings.Mid(Trim(ValeurDate), 4, 4)) = True Then
                                    ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 2, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 1) & "/" & Strings.Mid(Trim(ValeurDate), 4, 4))
                                End If
                            End If
                        End If
                    End If
                    Exit Select
                Case "aajjmm"
                    If Strings.Len(Trim(ValeurDate)) = 6 Then
                        If IsDate(Strings.Mid(Trim(ValeurDate), 5, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 2) & "/" & Strings.Mid(Trim(ValeurDate), 3, 2)) = True Then
                            ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 5, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 2) & "/" & Strings.Mid(Trim(ValeurDate), 3, 2))
                        End If
                    Else
                        If Strings.Len(Trim(ValeurDate)) = 5 Then
                            If IsNumeric(Trim(ValeurDate)) = True Then
                                If IsDate(Strings.Mid(Trim(ValeurDate), 4, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 1) & "/" & Strings.Mid(Trim(ValeurDate), 2, 2)) = True Then
                                    ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 4, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 1) & "/" & Strings.Mid(Trim(ValeurDate), 2, 2))
                                End If
                            End If
                        End If
                    End If
                    Exit Select
                Case "aaaajjmm"
                    If Strings.Len(Trim(ValeurDate)) = 8 Then
                        If IsDate(Strings.Mid(Trim(ValeurDate), 5, 4) & "/" & Strings.Mid(Trim(ValeurDate), 1, 2) & "/" & Strings.Mid(Trim(ValeurDate), 3, 2)) = True Then
                            ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 5, 4) & "/" & Strings.Mid(Trim(ValeurDate), 1, 2) & "/" & Strings.Mid(Trim(ValeurDate), 3, 2))
                        End If
                    Else
                        If Strings.Len(Trim(ValeurDate)) = 7 Then
                            If IsNumeric(Trim(ValeurDate)) = True Then
                                If IsDate(Strings.Mid(Trim(ValeurDate), 4, 4) & "/" & Strings.Mid(Trim(ValeurDate), 1, 1) & "/" & Strings.Mid(Trim(ValeurDate), 2, 2)) = True Then
                                    ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 4, 4) & "/" & Strings.Mid(Trim(ValeurDate), 1, 1) & "/" & Strings.Mid(Trim(ValeurDate), 2, 2))
                                End If
                            End If
                        End If
                    End If
                    Exit Select
                Case "mm/jj/aa"
                    If Strings.Len(Trim(ValeurDate)) = 8 Then
                        If IsDate(Strings.Mid(Trim(ValeurDate), 4, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 2) & "/" & Strings.Mid(Trim(ValeurDate), 7, 2)) = True Then
                            ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 4, 2) & "-" & Strings.Mid(Trim(ValeurDate), 1, 2) & "-" & Strings.Mid(Trim(ValeurDate), 7, 2))
                        End If
                    End If
                    Exit Select
                Case "mm-jj-aa"
                    If Strings.Len(Trim(ValeurDate)) = 8 Then
                        If IsDate(Strings.Mid(Trim(ValeurDate), 4, 2) & "-" & Strings.Mid(Trim(ValeurDate), 1, 2) & "-" & Strings.Mid(Trim(ValeurDate), 7, 2)) = True Then
                            ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 4, 2) & "-" & Strings.Mid(Trim(ValeurDate), 1, 2) & "-" & Strings.Mid(Trim(ValeurDate), 7, 2))
                        End If
                    End If
                    Exit Select
                Case "mm-jj-aaaa"
                    If Strings.Len(Trim(ValeurDate)) = 10 Then
                        If IsDate(Strings.Mid(Trim(ValeurDate), 4, 2) & "-" & Strings.Mid(Trim(ValeurDate), 1, 2) & "-" & Strings.Mid(Trim(ValeurDate), 7, 4)) = True Then
                            ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 4, 2) & "-" & Strings.Mid(Trim(ValeurDate), 1, 2) & "-" & Strings.Mid(Trim(ValeurDate), 7, 4))
                        End If
                    End If
                    Exit Select
                Case "aa-jj-mm"
                    If Strings.Len(Trim(ValeurDate)) = 8 Then
                        If IsDate(Strings.Mid(Trim(ValeurDate), 7, 2) & "-" & Strings.Mid(Trim(ValeurDate), 1, 2) & "-" & Strings.Mid(Trim(ValeurDate), 4, 2)) = True Then
                            ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 7, 2) & "-" & Strings.Mid(Trim(ValeurDate), 1, 2) & "-" & Strings.Mid(Trim(ValeurDate), 4, 2))
                        End If
                    End If
                    Exit Select
                Case "aaaa-jj-mm"
                    If Strings.Len(Trim(ValeurDate)) = 10 Then
                        If IsDate(Strings.Mid(Trim(ValeurDate), 7, 4) & "-" & Strings.Mid(Trim(ValeurDate), 1, 2) & "-" & Strings.Mid(Trim(ValeurDate), 4, 2)) = True Then
                            ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 7, 4) & "-" & Strings.Mid(Trim(ValeurDate), 1, 2) & "-" & Strings.Mid(Trim(ValeurDate), 4, 2))
                        End If
                    End If
                    Exit Select
                Case "mm/aa"
                    If Strings.Len(Trim(ValeurDate)) = 8 Then
                        If IsDate(Strings.Mid(Trim(ValeurDate), 1, 2) & "/" & Strings.Mid(Trim(ValeurDate), 4, 2) & "/" & Strings.Mid(Trim(ValeurDate), 7, 2)) = True Then
                            ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 1, 2) & "/" & Strings.Mid(Trim(ValeurDate), 4, 2) & "/" & Strings.Mid(Trim(ValeurDate), 7, 2))
                        End If
                    End If
                    Exit Select
                Case "aa/mm"
                    If Strings.Len(Trim(ValeurDate)) = 8 Then
                        If IsDate(Strings.Mid(Trim(ValeurDate), 7, 2) & "/" & Strings.Mid(Trim(ValeurDate), 4, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 2)) = True Then
                            ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 7, 2) & "/" & Strings.Mid(Trim(ValeurDate), 4, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 2))
                        End If
                    End If
                    Exit Select
                Case "mm/aaaa"
                    If Strings.Len(Trim(ValeurDate)) = 10 Then
                        If IsDate(Strings.Mid(Trim(ValeurDate), 1, 2) & "/" & Strings.Mid(Trim(ValeurDate), 4, 2) & "/" & Strings.Mid(Trim(ValeurDate), 7, 4)) = True Then
                            ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 1, 2) & "/" & Strings.Mid(Trim(ValeurDate), 4, 2) & "/" & Strings.Mid(Trim(ValeurDate), 7, 4))
                        End If
                    End If
                    Exit Select
                Case "aaaa/mm"
                    If Strings.Len(Trim(ValeurDate)) = 10 Then
                        If IsDate(Strings.Mid(Trim(ValeurDate), 9, 2) & "/" & Strings.Mid(Trim(ValeurDate), 6, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 4)) = True Then
                            ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 9, 2) & "/" & Strings.Mid(Trim(ValeurDate), 6, 2) & "/" & Strings.Mid(Trim(ValeurDate), 1, 4))
                        End If
                    End If
                    Exit Select
                Case "mm-aa"
                    If Strings.Len(Trim(ValeurDate)) = 8 Then
                        If IsDate(Strings.Mid(Trim(ValeurDate), 1, 2) & "-" & Strings.Mid(Trim(ValeurDate), 4, 2) & "-" & Strings.Mid(Trim(ValeurDate), 7, 2)) = True Then
                            ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 1, 2) & "-" & Strings.Mid(Trim(ValeurDate), 4, 2) & "-" & Strings.Mid(Trim(ValeurDate), 7, 2))
                        End If
                    End If
                    Exit Select
                Case "aa-mm"
                    If Strings.Len(Trim(ValeurDate)) = 8 Then
                        If IsDate(Strings.Mid(Trim(ValeurDate), 7, 2) & "-" & Strings.Mid(Trim(ValeurDate), 4, 2) & "-" & Strings.Mid(Trim(ValeurDate), 1, 2)) = True Then
                            ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 7, 2) & "-" & Strings.Mid(Trim(ValeurDate), 4, 2) & "-" & Strings.Mid(Trim(ValeurDate), 1, 2))
                        End If
                    End If
                    Exit Select
                Case "mm-aaaa"
                    If Strings.Len(Trim(ValeurDate)) = 10 Then
                        If IsDate(Strings.Mid(Trim(ValeurDate), 1, 2) & "-" & Strings.Mid(Trim(ValeurDate), 4, 2) & "-" & Strings.Mid(Trim(ValeurDate), 7, 4)) = True Then
                            ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 1, 2) & "-" & Strings.Mid(Trim(ValeurDate), 4, 2) & "-" & Strings.Mid(Trim(ValeurDate), 7, 4))
                        End If
                    End If
                    Exit Select
                Case "aaaa-mm"
                    If Strings.Len(Trim(ValeurDate)) = 10 Then
                        If IsDate(Strings.Mid(Trim(ValeurDate), 9, 2) & "-" & Strings.Mid(Trim(ValeurDate), 6, 2) & "-" & Strings.Mid(Trim(ValeurDate), 1, 4)) = True Then
                            ValeurDate = CDate(Strings.Mid(Trim(ValeurDate), 9, 2) & "-" & Strings.Mid(Trim(ValeurDate), 6, 2) & "-" & Strings.Mid(Trim(ValeurDate), 1, 4))
                        End If
                    End If
                    Exit Select
                Case Else
                    If IsDate(Trim(ValeurDate)) = True Then
                        ValeurDate = CDate(Trim(ValeurDate))
                    End If
                    Exit Select
            End Select
        Catch ex As Exception
        End Try
        RenvoieDateValide = ValeurDate
    End Function
    Private Function Verificatdate(ByRef Valeur As Object, ByRef DateFormat As String, ByRef Champ As String) As Boolean
        'hermann
        Dim Estsimuller As Boolean = True
        Select Case DateFormat
            Case "aa-mm-jj"
                If Strings.Len(Trim(Valeur)) = 8 Then
                    If IsDate(Strings.Mid(Trim(Valeur), 7, 2) & "-" & Strings.Mid(Trim(Valeur), 4, 2) & "-" & Strings.Mid(Trim(Valeur), 1, 2)) = True Then
                    Else
                        Estsimuller = False
                        ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                    End If
                Else
                    Estsimuller = False
                    ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                End If
            Case "aa-jj-mm"
                If Strings.Len(Trim(Valeur)) = 8 Then
                    If IsDate(Strings.Mid(Trim(Valeur), 7, 2) & "-" & Strings.Mid(Trim(Valeur), 1, 2) & "-" & Strings.Mid(Trim(Valeur), 4, 2)) = True Then
                    Else
                        Estsimuller = False
                        ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                    End If
                Else
                    Estsimuller = False
                    ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                End If
            Case "aaaa-mm-jj"
                If Strings.Len(Trim(Valeur)) = 10 Then
                    If IsDate(Strings.Mid(Trim(Valeur), 9, 2) & "-" & Strings.Mid(Trim(Valeur), 6, 2) & "-" & Strings.Mid(Trim(Valeur), 1, 4)) = True Then
                    Else
                        Estsimuller = False
                        ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                    End If
                Else
                    Estsimuller = False
                    ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                End If
            Case "aaaa-jj-mm"
                If Strings.Len(Trim(Valeur)) = 10 Then
                    If IsDate(Strings.Mid(Trim(Valeur), 9, 2) & "-" & Strings.Mid(Trim(Valeur), 1, 4) & "-" & Strings.Mid(Trim(Valeur), 6, 2)) = True Then
                    Else
                        Estsimuller = False
                        ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                    End If
                Else
                    Estsimuller = False
                    ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                End If
            Case "jj-mm-aa"
                If Strings.Len(Trim(Valeur)) = 8 Then
                    If IsDate(Strings.Mid(Trim(Valeur), 1, 2) & "-" & Strings.Mid(Trim(Valeur), 4, 2) & "-" & Strings.Mid(Trim(Valeur), 7, 2)) = True Then
                    Else
                        Estsimuller = False
                        ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                    End If
                Else
                    Estsimuller = False
                    ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                End If
            Case "mm-jj-aa"
                If Strings.Len(Trim(Valeur)) = 8 Then
                    If IsDate(Strings.Mid(Trim(Valeur), 4, 2) & "-" & Strings.Mid(Trim(Valeur), 1, 2) & "-" & Strings.Mid(Trim(Valeur), 7, 2)) = True Then
                    Else
                        Estsimuller = False
                        ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                    End If
                Else
                    Estsimuller = False
                    ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                End If
            Case "jj-mm-aaaa"
                If Strings.Len(Trim(Valeur)) = 10 Then
                    If IsDate(Strings.Mid(Trim(Valeur), 1, 2) & "-" & Strings.Mid(Trim(Valeur), 4, 2) & "-" & Strings.Mid(Trim(Valeur), 7, 4)) = True Then
                    Else
                        Estsimuller = False
                        ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                    End If
                Else
                    Estsimuller = False
                    ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                End If
            Case "mm-jj-aaaa"
                If Strings.Len(Trim(Valeur)) = 10 Then
                    If IsDate(Strings.Mid(Trim(Valeur), 4, 2) & "-" & Strings.Mid(Trim(Valeur), 1, 2) & "-" & Strings.Mid(Trim(Valeur), 7, 4)) = True Then
                    Else
                        Estsimuller = False
                        ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                    End If
                Else
                    Estsimuller = False
                    ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                End If
            Case "jjmmaa"
                If Strings.Len(Trim(Valeur)) = 6 Then
                    If IsDate(Strings.Mid(Trim(Valeur), 1, 2) & "/" & Strings.Mid(Trim(Valeur), 3, 2) & "/" & Strings.Mid(Trim(Valeur), 5, 2)) = True Then
                    Else
                        Estsimuller = False
                        ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                    End If
                Else
                    If Strings.Len(Trim(Valeur)) = 5 Then
                        If IsNumeric(Trim(Valeur)) = True Then
                            If IsDate(Strings.Mid(Format(CDbl(Trim(Valeur)), "000000"), 1, 2) & "/" & Strings.Mid(Format(CDbl(Trim(Valeur)), "000000"), 3, 2) & "/" & Strings.Mid(Format(CDbl(Trim(Valeur)), "000000"), 5, 2)) = True Then

                            Else
                                Estsimuller = False
                                ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                            End If
                        Else
                            Estsimuller = False
                            ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                        End If
                    Else
                        Estsimuller = False
                        ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                    End If
                End If
            Case "mmjjaa"
                If Strings.Len(Trim(Valeur)) = 6 Then
                    If IsDate(Strings.Mid(Trim(Valeur), 3, 2) & "/" & Strings.Mid(Trim(Valeur), 1, 2) & "/" & Strings.Mid(Trim(Valeur), 5, 2)) = True Then
                    Else
                        Estsimuller = False
                        ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                    End If
                Else
                    If Strings.Len(Trim(Valeur)) = 5 Then
                        If IsNumeric(Trim(Valeur)) = True Then
                            If IsDate(Strings.Mid(Format(CDbl(Trim(Valeur)), "000000"), 3, 2) & "/" & Strings.Mid(Format(CDbl(Trim(Valeur)), "000000"), 1, 2) & "/" & Strings.Mid(Format(CDbl(Trim(Valeur)), "000000"), 5, 2)) = True Then
                            Else
                                Estsimuller = False
                                ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                            End If
                        Else
                            Estsimuller = False
                            ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                        End If
                    Else
                        Estsimuller = False
                        ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                    End If
                End If
            Case "jjmmaaaa"
                If Strings.Len(Trim(Valeur)) = 8 Then
                    If IsDate(Strings.Mid(Trim(Valeur), 1, 2) & "/" & Strings.Mid(Trim(Valeur), 3, 2) & "/" & Strings.Mid(Trim(Valeur), 5, 4)) = True Then
                    Else
                        Estsimuller = False
                        ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                    End If
                Else
                    If Strings.Len(Trim(Valeur)) = 7 Then
                        If IsNumeric(Trim(Valeur)) = True Then
                            If IsDate(Strings.Mid(Format(CDbl(Trim(Valeur)), "00000000"), 1, 2) & "/" & Strings.Mid(Format(CDbl(Trim(Valeur)), "00000000"), 3, 2) & "/" & Strings.Mid(Format(CDbl(Trim(Valeur)), "00000000"), 5, 4)) = True Then
                            Else
                                Estsimuller = False
                                ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                            End If
                        Else
                            Estsimuller = False
                            ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                        End If
                    Else
                        Estsimuller = False
                        ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                    End If
                End If
            Case "mmjjaaaa"
                If Strings.Len(Trim(Valeur)) = 8 Then
                    If IsDate(Strings.Mid(Trim(Valeur), 3, 2) & "/" & Strings.Mid(Trim(Valeur), 1, 2) & "/" & Strings.Mid(Trim(Valeur), 5, 4)) = True Then
                    Else
                        Estsimuller = False
                        ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                    End If
                Else
                    If Strings.Len(Trim(Valeur)) = 7 Then
                        If IsNumeric(Trim(Valeur)) = True Then
                            If IsDate(Strings.Mid(Format(CDbl(Trim(Valeur)), "00000000"), 3, 2) & "/" & Strings.Mid(Format(CDbl(Trim(Valeur)), "00000000"), 1, 2) & "/" & Strings.Mid(Format(CDbl(Trim(Valeur)), "00000000"), 5, 4)) = True Then
                            Else
                                Estsimuller = False
                                ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                            End If
                        Else
                            Estsimuller = False
                            ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                        End If
                    Else
                        Estsimuller = False
                        ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                    End If
                End If
            Case "aammjj"
                If Strings.Len(Trim(Valeur)) = 6 Then
                    If IsDate(Strings.Mid(Trim(Valeur), 5, 2) & "/" & Strings.Mid(Trim(Valeur), 3, 2) & "/" & Strings.Mid(Trim(Valeur), 1, 2)) = True Then
                    Else
                        Estsimuller = False
                        ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                    End If
                Else
                    If Strings.Len(Trim(Valeur)) = 5 Then
                        If IsNumeric(Trim(Valeur)) = True Then
                            If IsDate(Strings.Mid(Format(CDbl(Trim(Valeur)), "000000"), 5, 2) & "/" & Strings.Mid(Format(CDbl(Trim(Valeur)), "000000"), 3, 2) & "/" & Strings.Mid(Format(CDbl(Trim(Valeur)), "000000"), 1, 2)) = True Then
                            Else
                                Estsimuller = False
                                ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                            End If
                        Else
                            Estsimuller = False
                            ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                        End If
                    Else
                        Estsimuller = False
                        ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                    End If
                End If
            Case "aajjmm"
                If Strings.Len(Trim(Valeur)) = 6 Then
                    If IsDate(Strings.Mid(Trim(Valeur), 5, 2) & "/" & Strings.Mid(Trim(Valeur), 1, 2) & "/" & Strings.Mid(Trim(Valeur), 3, 2)) = True Then
                    Else
                        Estsimuller = False
                        ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                    End If
                Else
                    If Strings.Len(Trim(Valeur)) = 5 Then
                        If IsNumeric(Trim(Valeur)) = True Then
                            If IsDate(Strings.Mid(Format(CDbl(Trim(Valeur)), "000000"), 5, 2) & "/" & Strings.Mid(Format(CDbl(Trim(Valeur)), "000000"), 1, 2) & "/" & Strings.Mid(Format(CDbl(Trim(Valeur)), "000000"), 3, 2)) = True Then
                            Else
                                Estsimuller = False
                                ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                            End If
                        Else
                            Estsimuller = False
                            ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                        End If
                    Else
                        Estsimuller = False
                        ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                    End If
                End If
            Case "aaaammjj"
                If Strings.Len(Trim(Valeur)) = 8 Then
                    If IsDate(Strings.Mid(Trim(Valeur), 7, 2) & "/" & Strings.Mid(Trim(Valeur), 5, 2) & "/" & Strings.Mid(Trim(Valeur), 1, 4)) = True Then
                    Else
                        Estsimuller = False
                        ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                    End If
                Else
                    If Strings.Len(Trim(Valeur)) = 7 Then
                        If IsNumeric(Trim(Valeur)) = True Then
                            If IsDate(Strings.Mid(Format(CDbl(Trim(Valeur)), "00000000"), 7, 2) & "/" & Strings.Mid(Format(CDbl(Trim(Valeur)), "00000000"), 5, 2) & "/" & Strings.Mid(Format(CDbl(Trim(Valeur)), "00000000"), 1, 4)) = True Then
                            Else
                                Estsimuller = False
                                ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                            End If
                        Else
                            Estsimuller = False
                            ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                        End If
                    Else
                        Estsimuller = False
                        ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                    End If
                End If
            Case "aaaajjmm"
                If Strings.Len(Trim(Valeur)) = 8 Then
                    If IsDate(Strings.Mid(Trim(Valeur), 7, 2) & "/" & Strings.Mid(Trim(Valeur), 1, 4) & "/" & Strings.Mid(Trim(Valeur), 5, 2)) = True Then
                    Else
                        Estsimuller = False
                        ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                    End If
                Else
                    If Strings.Len(Trim(Valeur)) = 7 Then
                        If IsNumeric(Trim(Valeur)) = True Then
                            If IsDate(Strings.Mid(Format(CDbl(Trim(Valeur)), "00000000"), 7, 2) & "/" & Strings.Mid(Format(CDbl(Trim(Valeur)), "00000000"), 1, 4) & "/" & Strings.Mid(Format(CDbl(Trim(Valeur)), "00000000"), 5, 2)) = True Then
                            Else
                                Estsimuller = False
                                ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                            End If
                        Else
                            Estsimuller = False
                            ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                        End If
                    Else
                        Estsimuller = False
                        ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                    End If
                End If
            Case "jj/mm/aa"
                If Strings.Len(Trim(Valeur)) = 8 Then
                    If IsDate(Strings.Mid(Trim(Valeur), 1, 2) & "/" & Strings.Mid(Trim(Valeur), 4, 2) & "/" & Strings.Mid(Trim(Valeur), 7, 2)) = True Then
                    Else
                        Estsimuller = False
                        ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                    End If
                Else
                    Estsimuller = False
                    ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                End If
            Case "mm/jj/aa"
                If Strings.Len(Trim(Valeur)) = 8 Then
                    If IsDate(Strings.Mid(Trim(Valeur), 4, 2) & "/" & Strings.Mid(Trim(Valeur), 1, 2) & "/" & Strings.Mid(Trim(Valeur), 7, 2)) = True Then
                    Else
                        Estsimuller = False
                        ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                    End If
                Else
                    Estsimuller = False
                    ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                End If
            Case "jj/mm/aaaa"
                If Strings.Len(Trim(Valeur)) = 10 Then
                    If IsDate(Strings.Mid(Trim(Valeur), 1, 2) & "/" & Strings.Mid(Trim(Valeur), 4, 2) & "/" & Strings.Mid(Trim(Valeur), 7, 4)) = True Then
                    Else
                        Estsimuller = False
                        ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                    End If
                Else
                    Estsimuller = False
                    ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                End If
            Case "mm/jj/aaaa"
                If Strings.Len(Trim(Valeur)) = 10 Then
                    If IsDate(Strings.Mid(Trim(Valeur), 4, 2) & "/" & Strings.Mid(Trim(Valeur), 1, 2) & "/" & Strings.Mid(Trim(Valeur), 7, 4)) = True Then
                    Else
                        Estsimuller = False
                        ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                    End If
                Else
                    Estsimuller = False
                    ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                End If
            Case Else
                If IsDate(Trim(Valeur)) = True Then
                Else
                    Estsimuller = False
                    ErreurJrn.WriteLine("La valeur entrée :" & Trim(Valeur) & " pour le Champ : " & Champ & "  doit être au format date :" & DateFormat)
                End If
                Exit Select
        End Select
        Verificatdate = Estsimuller
    End Function
    Private Function LirePieceModification(ByRef ScheminFileFormat As String, ByRef Lireformatype As String) As Object
        Dim NomColonne As String
        Dim NomEntete As String
        Dim PosLeft As Integer
        Dim poslongueur, LigneArticle As String
        Dim Defaut, Piece, SageFichier As String
        Dim ValeurDefaut, Infolibre As String
        Dim typedeformat As Object = Nothing
        Dim typeImport As Object = Nothing
        Dim ModeFormat As Object = Nothing
        Dim DateFormat As Object = Nothing
        Dim PieceAuto, Punitaire As Object
        Try
            If Trim(ScheminFileFormat) <> "" Then
                If File.Exists(ScheminFileFormat) = True Then
                    Dim FileXml As New XmlTextReader(Trim(ScheminFileFormat))
                    While (FileXml.Read())
                        If FileXml.LocalName = "ColUse" Then
                            NomColonne = FileXml.ReadString
                            FileXml.Read()
                            NomEntete = FileXml.ReadString

                            FileXml.Read()
                            PosLeft = FileXml.ReadString

                            If Trim(Lireformatype) = "Excel" Then
                            Else
                                If Trim(Lireformatype) = "Longueur Fixe" Then
                                    FileXml.Read()
                                    poslongueur = FileXml.ReadString
                                End If
                            End If

                            FileXml.Read()
                            Infolibre = FileXml.ReadString

                            FileXml.Read()
                            SageFichier = FileXml.ReadString

                            FileXml.Read()
                            Piece = FileXml.ReadString

                            FileXml.Read()
                            LigneArticle = FileXml.ReadString

                            FileXml.Read()
                            Defaut = FileXml.ReadString

                            FileXml.Read()
                            ValeurDefaut = FileXml.ReadString

                            FileXml.Read()
                            DecFormat = FileXml.ReadString

                            FileXml.Read()
                            typedeformat = FileXml.ReadString

                            FileXml.Read()
                            ModeFormat = FileXml.ReadString

                            FileXml.Read()
                            DateFormat = FileXml.ReadString

                            FileXml.Read()
                            PieceAuto = FileXml.ReadString

                            FileXml.Read()
                            Punitaire = FileXml.ReadString

                            FileXml.Read()
                            typeImport = FileXml.ReadString

                            If Trim(Piece) = "oui" Then
                                Result = NomEntete
                                Exit While
                            End If
                        End If
                    End While
                    FileXml.Close()
                End If
            End If
        Catch ex As Exception
        End Try
        Return Result
    End Function
    Private Function LirePieceCreation(ByRef ScheminFileFormat As String, ByRef Lireformatype As String) As Object
        Dim NomColonne As String
        Dim NomEntete As String
        Dim PosLeft As Integer
        Dim poslongueur, LigneArticle As String
        Dim Defaut, Piece, SageFichier As String
        Dim ValeurDefaut, Infolibre As String
        Dim PieceAuto, Punitaire As Object
        Dim typedeformat As Object = Nothing
        Dim typeImport As Object = Nothing
        Dim ModeFormat As Object = Nothing
        Dim DateFormat As Object = Nothing
        Result = ""
        Try
            If Trim(ScheminFileFormat) <> "" Then
                If File.Exists(ScheminFileFormat) = True Then
                    Dim FileXml As New XmlTextReader(Trim(ScheminFileFormat))
                    While (FileXml.Read())
                        If FileXml.LocalName = "ColUse" Then
                            NomColonne = FileXml.ReadString

                            FileXml.Read()
                            NomEntete = FileXml.ReadString

                            FileXml.Read()
                            PosLeft = FileXml.ReadString

                            If Trim(Lireformatype) = "Excel" Then
                            Else
                                If Trim(Lireformatype) = "Longueur Fixe" Then
                                    FileXml.Read()
                                    poslongueur = FileXml.ReadString
                                End If
                            End If

                            FileXml.Read()
                            Infolibre = FileXml.ReadString

                            FileXml.Read()
                            SageFichier = FileXml.ReadString

                            FileXml.Read()
                            Piece = FileXml.ReadString

                            FileXml.Read()
                            LigneArticle = FileXml.ReadString

                            FileXml.Read()
                            Defaut = FileXml.ReadString

                            FileXml.Read()
                            ValeurDefaut = FileXml.ReadString

                            FileXml.Read()
                            DecFormat = FileXml.ReadString

                            FileXml.Read()
                            typedeformat = FileXml.ReadString

                            FileXml.Read()
                            ModeFormat = FileXml.ReadString

                            FileXml.Read()
                            DateFormat = FileXml.ReadString

                            FileXml.Read()
                            PieceAuto = FileXml.ReadString

                            FileXml.Read()
                            Punitaire = FileXml.ReadString

                            FileXml.Read()
                            typeImport = FileXml.ReadString

                            If Trim(Piece) = "oui" Then
                                Result = NomEntete
                                Exit While
                            End If
                        End If
                    End While
                    FileXml.Close()
                End If
            End If
        Catch ex As Exception
        End Try
        Return Result
    End Function
    Private Function LirePUDefaut(ByRef ScheminFileFormat As String, ByRef Lireformatype As String) As Object
        Dim NomColonne As String
        Dim NomEntete As String
        Dim PosLeft As Integer
        Dim typedeformat As Object = Nothing
        Dim typeImport As Object = Nothing
        Dim ModeFormat As Object = Nothing
        Dim DateFormat As Object = Nothing
        Dim poslongueur, Punitaire, LigneArticle As String
        Dim Defaut, Piece, SageFichier As String
        Dim Decal, PieceAuto, ValeurDefaut, Infolibre As String
        Result = ""
        Try
            If Trim(ScheminFileFormat) <> "" Then
                If File.Exists(ScheminFileFormat) = True Then
                    Dim FileXml As New XmlTextReader(Trim(ScheminFileFormat))
                    While (FileXml.Read())
                        If FileXml.LocalName = "ColUse" Then
                            NomColonne = FileXml.ReadString

                            FileXml.Read()
                            NomEntete = FileXml.ReadString

                            FileXml.Read()
                            PosLeft = FileXml.ReadString

                            If Trim(Lireformatype) = "Excel" Then
                            Else
                                If Trim(Lireformatype) = "Longueur Fixe" Then
                                    FileXml.Read()
                                    poslongueur = FileXml.ReadString
                                End If
                            End If

                            FileXml.Read()
                            Infolibre = FileXml.ReadString

                            FileXml.Read()
                            SageFichier = FileXml.ReadString

                            FileXml.Read()
                            Piece = FileXml.ReadString

                            FileXml.Read()
                            LigneArticle = FileXml.ReadString

                            FileXml.Read()
                            Defaut = FileXml.ReadString

                            FileXml.Read()
                            ValeurDefaut = FileXml.ReadString

                            FileXml.Read()
                            Decal = FileXml.ReadString

                            FileXml.Read()
                            typedeformat = FileXml.ReadString

                            FileXml.Read()
                            ModeFormat = FileXml.ReadString

                            FileXml.Read()
                            DateFormat = FileXml.ReadString

                            FileXml.Read()
                            PieceAuto = FileXml.ReadString

                            FileXml.Read()
                            Punitaire = FileXml.ReadString

                            FileXml.Read()
                            typeImport = FileXml.ReadString

                            If Trim(Punitaire) <> "" Then
                                Result = Trim(Punitaire)
                                Exit While
                            End If
                        End If
                    End While
                    FileXml.Close()
                End If
            End If
        Catch ex As Exception
        End Try
        Return Result
    End Function
    Private Function LireLigneArticle(ByRef ScheminFileFormat As String, ByRef Lireformatype As String) As Object
        Dim NomColonne As String
        Dim NomEntete As String
        Dim PosLeft As Integer
        Dim poslongueur, LigneArticle As String
        Dim Defaut, Piece, SageFichier As String
        Dim ValeurDefaut, Infolibre As String
        Dim typedeformat As Object = Nothing
        Dim typeImport As Object = Nothing
        Dim ModeFormat As Object = Nothing
        Dim DateFormat As Object = Nothing
        Dim PieceAuto, Punitaire As Object
        Result = ""
        Try
            If Trim(ScheminFileFormat) <> "" Then
                If File.Exists(ScheminFileFormat) = True Then
                    Dim FileXml As New XmlTextReader(Trim(ScheminFileFormat))
                    While (FileXml.Read())
                        If FileXml.LocalName = "ColUse" Then
                            NomColonne = FileXml.ReadString

                            FileXml.Read()
                            NomEntete = FileXml.ReadString

                            FileXml.Read()
                            PosLeft = FileXml.ReadString

                            If Trim(Lireformatype) = "Excel" Then
                            Else
                                If Trim(Lireformatype) = "Longueur Fixe" Then
                                    FileXml.Read()
                                    poslongueur = FileXml.ReadString
                                End If
                            End If

                            FileXml.Read()
                            Infolibre = FileXml.ReadString

                            FileXml.Read()
                            SageFichier = FileXml.ReadString

                            FileXml.Read()
                            Piece = FileXml.ReadString

                            FileXml.Read()
                            LigneArticle = FileXml.ReadString

                            FileXml.Read()
                            Defaut = FileXml.ReadString

                            FileXml.Read()
                            ValeurDefaut = FileXml.ReadString

                            FileXml.Read()
                            DecFormat = FileXml.ReadString

                            FileXml.Read()
                            typedeformat = FileXml.ReadString

                            FileXml.Read()
                            ModeFormat = FileXml.ReadString

                            FileXml.Read()
                            DateFormat = FileXml.ReadString

                            FileXml.Read()
                            PieceAuto = FileXml.ReadString

                            FileXml.Read()
                            Punitaire = FileXml.ReadString

                            FileXml.Read()
                            typeImport = FileXml.ReadString

                            If Trim(LigneArticle) = "oui" Then
                                Result = NomEntete
                                Exit While
                            End If
                        End If
                    End While
                    FileXml.Close()
                End If
            End If
        Catch ex As Exception
        End Try
        Return Result
    End Function
    Private Function RenvoieFormatDate(ByVal CheminFichier) As String
        Dim xdoc As XmlDocument
        Dim racine As XmlElement
        Dim nodelist As XmlNodeList
        Dim FormatDatelu As Object = Nothing
        Dim i As Integer
        Try
            If File.Exists(CheminFichier) = True Then
                Dim FileXml As New XmlTextReader(Trim(CheminFichier))
                xdoc = New XmlDocument
                xdoc.Load(Trim(CheminFichier))
                racine = xdoc.DocumentElement
                nodelist = racine.ChildNodes
                For i = 0 To nodelist.Count - 1
                    If Trim(nodelist.ItemOf(i).Name) = "DateFormat" Then
                        FormatDatelu = nodelist.ItemOf(i).InnerText
                        Exit For
                    End If
                Next i
                FileXml.Close()
            Else
                MsgBox("Nom du Format inexistant!", MsgBoxStyle.Information, "Format d'integration")
            End If
        Catch ex As Exception

        End Try
        RenvoieFormatDate = FormatDatelu
    End Function
    Private Function LirePieceAuto(ByVal CheminFichier) As String
        Dim xdoc As XmlDocument
        Dim racine As XmlElement
        Dim nodelist As XmlNodeList
        Dim PieceAuto As Object = Nothing
        Dim i As Integer
        Try
            If File.Exists(CheminFichier) = True Then
                Dim FileXml As New XmlTextReader(Trim(CheminFichier))
                xdoc = New XmlDocument
                xdoc.Load(Trim(CheminFichier))
                racine = xdoc.DocumentElement
                nodelist = racine.ChildNodes
                For i = 0 To nodelist.Count - 1
                    If Trim(nodelist.ItemOf(i).Name) = "PieceAuto" Then
                        PieceAuto = nodelist.ItemOf(i).InnerText
                        Exit For
                    End If
                Next i
                FileXml.Close()
            Else
                MsgBox("Nom du Format inexistant!", MsgBoxStyle.Information, "Format d'integration")
            End If
        Catch ex As Exception

        End Try
        LirePieceAuto = PieceAuto
    End Function

    Private Function LireTypeImport(ByVal CheminFichier) As String
        Dim xdoc As XmlDocument
        Dim racine As XmlElement
        Dim nodelist As XmlNodeList
        Dim TypeImport As Object = Nothing
        Dim i As Integer
        Try
            If File.Exists(CheminFichier) = True Then
                Dim FileXml As New XmlTextReader(Trim(CheminFichier))
                xdoc = New XmlDocument
                xdoc.Load(Trim(CheminFichier))
                racine = xdoc.DocumentElement
                nodelist = racine.ChildNodes
                For i = 0 To nodelist.Count - 1
                    If Trim(nodelist.ItemOf(i).Name) = "IMPORT" Then
                        TypeImport = nodelist.ItemOf(i).InnerText
                        Exit For
                    End If
                Next i
                FileXml.Close()
            Else
                MsgBox("Nom du Format inexistant!", MsgBoxStyle.Information, "Format d'integration")
            End If
        Catch ex As Exception

        End Try
        LireTypeImport = TypeImport
    End Function

    Private Sub BT_SelAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_SelAll.Click
        Dim i As Integer
        For i = 0 To DataListeIntegrer.RowCount - 1
            DataListeIntegrer.Rows(i).Cells("Valider").Value = True
        Next i
    End Sub

    Private Sub BT_DelAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_DelAll.Click
        Dim i As Integer
        For i = 0 To DataListeIntegrer.RowCount - 1
            DataListeIntegrer.Rows(i).Cells("Valider").Value = False
        Next i
    End Sub
End Class
