Imports System.IO
Imports Objets100Lib
Imports System.Data.OleDb
Public Class Frm_MvtE_S
#Region "Declaration des variable"
    Public Critere As String = ""
    Public CountChecked As Integer = 0
    Public NbInfosLibre As Integer
    Public NbInfosLibreVue As Integer
    Public PositionG As Integer = 0
    Public Longueur As Integer = 0
    Public DefaultValeur As String = ""
    Public OleAdaptaterschemaSage, OleAdaptaterschemaSageDetails, OleAdaptaterschema, OleAdaptaterschemaLigne, OleAdaptaterschemaFourssAR As OleDbDataAdapter
    Public OleSchemaDatasetSage, OleSchemaDatasetSageDetails, OleSchemaDataset, OleSchemaDatasetLigne, OleSchemaDatasetFourssAR As DataSet
    Public OleAdaptaterschemaDétailLigne As OleDbDataAdapter
    Public OleSchemaDatasetDétailLigne As DataSet
    Public OledatableSchema, OledatableSchemaLigne, OledatableSchemaDétailLigne As DataTable
    Public DefaultValue As String = ""
    'Public ListeStock As List(Of String)
    'Public ExisteLecture As Boolean = True

    'Public Documents As IBODocumentVente3 = Nothing
    'Public LigneDocument As IBODocumentVenteLigne3 = Nothing
    'Public DocumentInfolibre As IBODocumentVente3 = Nothing
    'Public DocumentReliquat As IBODocumentVente3 = Nothing
    'Public LigneReliquat As IBODocumentVenteLigne3 = Nothing
#End Region
#Region "Variable import Mvt"
    Public EnteteIntituleDepot, EntetePieceInterne, EntetePiecePrecedent, EnteteReference, EntetePlanAnalytique As Object
    Public EnteteSoucheDocument, EnteteTyPeDocument, LigneDatedeFabrication, LigneDatedeLivraison, LigneDatedePeremption As Object
    Public LigneDesignationArticle, IDDepotEntete, LigneNSerieLot, LigneCodeArticle, PieceArticle, EnteteDateDocument As Object
    Public LignePoidsBrut, LignePoidsNet, LignePrixUnitaire, LigneQuantite, LigneReference, CodeSociete, EnteteCodeAffaire As Object
    Public OleSocieteConnect As OleDbConnection
    'Variable d'exception du deplacement de fichier
    Public exceptionTrouve As Boolean = False
    Public ExisteLecture As Boolean = True
    Public NomFichier As String
    Public infoListe As List(Of Integer)
    Public infoLigne As List(Of Integer)
    Public ListePiece As List(Of String)
    Public ListeStock As List(Of String)
    Public Document As IBODocumentStock3 = Nothing
    Public LigneDocument As IBODocumentStockLigne3 = Nothing
    Public DocumentInfolibre As IBODocumentStock3 = Nothing
    Public PlanAna As IBPAnalytique3
#End Region
    Private Sub OuvreLaListedeFichier(ByRef Directpath As String)
        Dim i As Integer
        Dim NomFichier As String
        Dim aLines() As String
        Dim jRow As Integer
        Datagridaffiche.Rows.Clear()
        Try
            If Directory.Exists(Directpath) = True Then
                aLines = Directory.GetFiles(Directpath)
                For i = 0 To UBound(aLines) ' - 1
                    NomFichier = Trim(aLines(i))
                    Do While InStr(Trim(NomFichier), "\") <> 0
                        NomFichier = Strings.Right(NomFichier, Strings.Len(Trim(NomFichier)) - InStr(Trim(NomFichier), "\"))
                    Loop
                    If Critere = NomFichier.Substring(0, 3) Then
                        Select Case NomFichier.Substring(0, 3)
                            Case "VST"
                                Datagridaffiche.RowCount = jRow + 1
                                Datagridaffiche.Rows(jRow).Cells("C1").Value = "Mouvement Entrée/Sortie"
                                Datagridaffiche.Rows(jRow).Cells("C2").Value = "Exportation"
                                Datagridaffiche.Rows(jRow).Cells("C3").Value = NomFichier.Substring(0, 3)
                                Datagridaffiche.Rows(jRow).Cells("C4").Value = NomFichier.Substring(3, 2)
                                Datagridaffiche.Rows(jRow).Cells("C5").Value = NomFichier
                                Datagridaffiche.Rows(jRow).Cells("C6").Value = True
                                Datagridaffiche.Rows(jRow).Cells("C7").Value = My.Resources.btFermer22
                                Datagridaffiche.Rows(jRow).Cells("C8").Value = aLines(i)
                                jRow = jRow + 1
                                btnview.Enabled = True
                                BT_integrer.Enabled = True
                        End Select
                    End If
                Next i
                aLines = Nothing
            Else
                MsgBox("Ce Repertoire n'est pas valide! " & Directpath, MsgBoxStyle.Information, "Repertoire des Fichiers à Traiter")
            End If
        Catch ex As Exception
        End Try
    End Sub
    Public Vsate As Boolean = True
    Public Function IsChecked() As Integer
        Try
            Dim CountCheked As Integer
            For i As Integer = 0 To Datagridaffiche.RowCount - 1
                If Datagridaffiche.Rows(i).Cells("C6").Value = True Then
                    CountCheked += 1
                End If
            Next
            Return CountCheked
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Public StatutCreationEnteteDoc As Boolean = False
    Public StatutCreationLigneDoc As Boolean = False
    Private Sub Creation_Entete_Document(ByRef typedoc As String, Optional ByRef FormatDatefichier As String = "", Optional ByRef CreationPieceDocument As Object = "", Optional ByRef PieceInterne As Object = "", Optional ByRef PieceAutomatique As Object = "")
        infosExport.Text = "Création de l'Entête du Mvt"
        Dim statistAdap As OleDbDataAdapter
        Dim statistDs As DataSet
        Dim statistTab As DataTable
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
            If Trim(EntetePlanAnalytique) <> "" Then
                If Trim(EnteteCodeAffaire) <> "" Then
                    If BaseCpta.FactoryAnalytique.ExistIntitule(Trim(EntetePlanAnalytique)) = True Then
                        PlanAna = BaseCpta.FactoryAnalytique.ReadIntitule(Trim(EntetePlanAnalytique))
                        If BaseCpta.FactoryCompteA.ExistNumero(PlanAna, Trim(EnteteCodeAffaire)) = True Then
                            .CompteA = BaseCpta.FactoryCompteA.ReadNumero(PlanAna, Trim(EnteteCodeAffaire))
                        End If
                    End If
                End If
            End If
            If Trim(IDDepotEntete) <> "" Then
                If IsNumeric(Trim(IDDepotEntete)) = True Then
                    statistAdap = New OleDbDataAdapter("select * from F_DEPOT where DE_No ='" & CInt(Trim(IDDepotEntete)) & "'", OleSocieteConnect)
                    statistDs = New DataSet
                    statistAdap.Fill(statistDs)
                    statistTab = statistDs.Tables(0)
                    If statistTab.Rows.Count <> 0 Then
                        .DepotStockage = BaseCial.FactoryDepot.ReadIntitule(Trim(statistTab.Rows(0).Item("DE_Intitule")))
                    End If
                End If
            End If
            If Trim(EnteteIntituleDepot) <> "" Then
                If BaseCial.FactoryDepot.ExistIntitule(Trim(EnteteIntituleDepot)) = True Then
                    .DepotStockage = BaseCial.FactoryDepot.ReadIntitule(Trim(EnteteIntituleDepot))
                Else
                    If BaseCial.FactoryDepot.ExistIntitule(Trim(EnteteIntituleDepot)) = True Then
                        .DepotStockage = BaseCial.FactoryDepot.ReadIntitule(Trim(EnteteIntituleDepot))
                    End If
                End If
            End If

            If Trim(EnteteDateDocument) <> "" Then
                If Trim(EnteteDateDocument) <> "" Then
                    .DO_Date = RenvoieDateValide(Trim(EnteteDateDocument), FormatDatefichier)
                End If
            End If

            If ChkPieceAuto.Checked = False Then
                If EntetePieceInterne = "" Then
                    .DO_Piece = 1
                Else
                    .DO_Piece = EntetePieceInterne
                End If
            Else
                If Trim(EnteteSoucheDocument) <> "" Then
                    If Trim(EnteteSoucheDocument) <> "" Then
                        If EstNumeric(Trim(EnteteSoucheDocument), DecimalNomb, DecimalMone) = True Then

                        Else
                            If BaseCial.FactorySoucheStock.ExistIntitule(Trim(EnteteSoucheDocument)) = True Then
                                If BaseCial.FactorySoucheStock.ReadIntitule(Trim(EnteteSoucheDocument)).IsValide = True Then
                                    If typedoc = "20" Then
                                        .Souche = BaseCial.FactorySoucheStock.ReadIntitule(Trim(EnteteSoucheDocument))
                                        .DO_Piece = BaseCial.FactorySoucheStock.ReadIntitule(Trim(EnteteSoucheDocument)).ReadCurrentPiece(DocumentType.DocumentTypeStockMouvIn, DocumentProvenanceType.DocProvenanceNormale)
                                    Else
                                        If typedoc = "21" Then
                                            .Souche = BaseCial.FactorySoucheStock.ReadIntitule(Trim(EnteteSoucheDocument))
                                            .DO_Piece = BaseCial.FactorySoucheStock.ReadIntitule(Trim(EnteteSoucheDocument)).ReadCurrentPiece(DocumentType.DocumentTypeStockMouvOut, DocumentProvenanceType.DocProvenanceNormale)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            If Trim(EnteteReference) <> "" Then
                .DO_Ref = EnteteReference
            End If
            .Write()
            ErreurJrn.WriteLine("-----------------------------------------------------------------------------------------------------")
            ErreurJrn.WriteLine("")
            If typedoc = "20" Then
                ErreurJrn.WriteLine("Mouvement d'entrée N° : " & Trim(Document.DO_Piece) & " Créé Pour la pièce N° :" & Trim(EntetePieceInterne))
            Else
                If typedoc = "21" Then
                    ErreurJrn.WriteLine("Mouvement de sortie N° : " & Trim(Document.DO_Piece) & " Créé Pour la pièce N° :" & Trim(EntetePieceInterne))
                End If
            End If
            'Traitement des Infos Libres
            Try
                'If infoListe.Count > 0 Then
                '    While infoListe.Count <> 0
                '        OleAdaptaterDelete = New OleDbDataAdapter("select * From WIS_COL where Libelle='" & Trim(Datagridaffiche.Columns(infoListe.Item(0)).Name) & "' And Libre=True", OleConnenection)
                '        OleDeleteDataset = New DataSet
                '        OleAdaptaterDelete.Fill(OleDeleteDataset)
                '        OledatableDelete = OleDeleteDataset.Tables(0)
                '        If OledatableDelete.Rows.Count <> 0 Then
                '            'L'info Libre Parametrée par l'utilisateur existe dans Sage
                '            If Document.InfoLibre.Count <> 0 Then
                '                If IsNothing(infoListe.Item(0)) = False Then
                '                    If Trim(infoListe.Item(0)) <> "" Then
                '                        statistAdap = New OleDbDataAdapter("select * from cbSysLibre where CB_File='F_DOCENTETE' and CB_Name ='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "'", OleSocieteConnect)
                '                        statistDs = New DataSet
                '                        statistAdap.Fill(statistDs)
                '                        statistTab = statistDs.Tables(0)
                '                        If statistTab.Rows.Count <> 0 Then
                '                            'Texte
                '                            If statistTab.Rows(0).Item("CB_Type") = 9 Then
                '                                If statistTab.Rows(0).Item("CB_Len") >= Strings.Len(Trim(infoListe.Item(0))) Then
                '                                    Document.InfoLibre.Item("" & infoListe.Item(0) & "") = Trim(infoListe.Item(0))
                '                                    Document.Write()
                '                                End If
                '                            End If
                '                            'Table
                '                            If statistTab.Rows(0).Item("CB_Type") = 22 Then
                '                                If statistTab.Rows(0).Item("CB_Len") > Strings.Len(Trim(infoListe.Item(0))) Then
                '                                    Document.InfoLibre.Item("" & infoListe.Item(0) & "") = Trim(infoListe.Item(0))
                '                                    Document.Write()
                '                                End If
                '                            End If
                '                            'Montant
                '                            If statistTab.Rows(0).Item("CB_Type") = 20 Then
                '                                If Trim(infoListe.Item(0)) <> "" Then
                '                                    If EstNumeric(Trim(infoListe.Item(0)), DecimalNomb, DecimalMone) = True Then
                '                                        Document.InfoLibre.Item("" & infoListe.Item(0) & "") = CDbl(RenvoiTaux(Trim(infoListe.Item(0)), DecimalNomb, DecimalMone))
                '                                        Document.Write()
                '                                    End If
                '                                End If
                '                            End If
                '                            'Valeur
                '                            If statistTab.Rows(0).Item("CB_Type") = 7 Then
                '                                If Trim(infoListe.Item(0)) <> "" Then
                '                                    If EstNumeric(Trim(infoListe.Item(0)), DecimalNomb, DecimalMone) = True Then
                '                                        Document.InfoLibre.Item("" & infoListe.Item(0) & "") = CDbl(RenvoiTaux(Trim(infoListe.Item(0)), DecimalNomb, DecimalMone))
                '                                        Document.Write()
                '                                    End If
                '                                End If
                '                            End If

                '                            'Date Court
                '                            If statistTab.Rows(0).Item("CB_Type") = 3 Then
                '                                If Trim(infoListe.Item(0)) <> "" Then
                '                                    Document.InfoLibre.Item("" & infoListe.Item(0) & "") = RenvoieDateValide(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier)
                '                                    Document.Write()
                '                                End If
                '                            End If
                '                            'Date Longue
                '                            If statistTab.Rows(0).Item("CB_Type") = 14 Then
                '                                If Trim(infoListe.Item(0)) <> "" Then
                '                                    Document.InfoLibre.Item("" & infoListe.Item(0) & "") = RenvoieDateValide(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier)
                '                                    Document.Write()
                '                                End If
                '                            End If
                '                        End If
                '                    End If
                '                Else
                '                    'nothing
                '                End If
                '            End If
                '        End If
                '        'L'info Libre Parametrée par l'utilisateur n'existe pas dans Sage
                '        infoListe.RemoveAt(0)
                '    End While
                'End If
            Catch ex As Exception
                exceptionTrouve = True
                If typedoc = "20" Then
                    ErreurJrn.WriteLine("Mouvement d'entrée N° : " & Trim(Document.DO_Piece) & " Erreur de Modification de L'information Libre , vérifiez la longueur de la chaine avec celle paramétrée/ ou la cohérence des données " & ex.Message)
                Else
                    If typedoc = "21" Then
                        ErreurJrn.WriteLine("Mouvement de sortie N° : " & Trim(Document.DO_Piece) & " Erreur de Modification de L'information Libre , vérifiez la longueur de la chaine avec celle paramétrée/ ou la cohérence des données " & ex.Message)
                    End If
                End If
            End Try
        End With
    End Sub

    Private Sub PriUnitaireEntrer(ByRef Lignedocument As IBODocumentStockLigne3, ByRef IDdepotStoc As Integer, ByRef RefArticle As String)
        Dim DossierAdap As OleDbDataAdapter
        Dim DossierDs As DataSet
        Dim DossierTab As DataTable
        Dim statistAdap As OleDbDataAdapter
        Dim statistDs As DataSet
        Dim statistTab As DataTable
        With Lignedocument
            DossierAdap = New OleDbDataAdapter("select * from F_ARTSTOCK where DE_No ='" & CInt(Trim(IDdepotStoc)) & "' And AR_Ref='" & Trim(RefArticle) & "' and AS_QteSto <> 0 and AS_MontSto <> 0", OleSocieteConnect)
            DossierDs = New DataSet
            DossierAdap.Fill(DossierDs)
            DossierTab = DossierDs.Tables(0)
            If DossierTab.Rows.Count <> 0 Then
                .DL_PrixUnitaire = DossierTab.Rows(0).Item("AS_MontSto") / DossierTab.Rows(0).Item("AS_QteSto")
            Else
                DossierAdap = New OleDbDataAdapter("select * from F_ARTSTOCK where DE_No ='" & CInt(Trim(IDdepotStoc)) & "' And AR_Ref='" & Trim(RefArticle) & "' and AS_Principal=1 and AS_QteSto <> 0 and AS_MontSto <> 0", OleSocieteConnect)
                DossierDs = New DataSet
                DossierAdap.Fill(DossierDs)
                DossierTab = DossierDs.Tables(0)
                If DossierTab.Rows.Count <> 0 Then
                    .DL_PrixUnitaire = DossierTab.Rows(0).Item("AS_MontSto") / DossierTab.Rows(0).Item("AS_QteSto")
                Else
                    If .Article.AR_PUNet <> 0 Then
                        .DL_PrixUnitaire = .Article.AR_PUNet
                    Else
                        DossierAdap = New OleDbDataAdapter("select * from F_ARTFOURNISS where  AR_Ref='" & Trim(RefArticle) & "' and AF_Principal=1 And (AF_PrixDev<> 0 OR AF_PrixAch<> 0)", OleSocieteConnect)
                        DossierDs = New DataSet
                        DossierAdap.Fill(DossierDs)
                        DossierTab = DossierDs.Tables(0)
                        If DossierTab.Rows.Count <> 0 Then
                            Dim OleDocAdapter As OleDbDataAdapter
                            Dim OleDocDataset As DataSet
                            Dim OleDocDatable As DataTable
                            OleDocAdapter = New OleDbDataAdapter("Select  * From P_DOSSIERCIAL WHERE N_DeviseCompte =" & DossierTab.Rows(0).Item("AF_Devise") & "", OleSocieteConnect)
                            OleDocDataset = New DataSet
                            OleDocAdapter.Fill(OleDocDataset)
                            OleDocDatable = OleDocDataset.Tables(0)
                            If OleDocDatable.Rows.Count <> 0 Then
                                If DossierTab.Rows(0).Item("AF_PrixDev") <> 0 Then
                                    If Convert.IsDBNull(DossierTab.Rows(0).Item("AF_Remise")) = False Then
                                        If DossierTab.Rows(0).Item("AF_Remise") <> 0 Then
                                            .DL_PrixUnitaire = DossierTab.Rows(0).Item("AF_PrixDev") - ((DossierTab.Rows(0).Item("AF_PrixDev") * DossierTab.Rows(0).Item("AF_Remise")) / 100)
                                        Else
                                            .DL_PrixUnitaire = DossierTab.Rows(0).Item("AF_PrixDev")
                                        End If
                                    Else
                                        .DL_PrixUnitaire = DossierTab.Rows(0).Item("AF_PrixDev")
                                    End If
                                End If
                                If DossierTab.Rows(0).Item("AF_PrixAch") <> 0 Then
                                    If Convert.IsDBNull(DossierTab.Rows(0).Item("AF_Remise")) = False Then
                                        If DossierTab.Rows(0).Item("AF_Remise") <> 0 Then
                                            .DL_PrixUnitaire = DossierTab.Rows(0).Item("AF_PrixAch") - ((DossierTab.Rows(0).Item("AF_PrixAch") * DossierTab.Rows(0).Item("AF_Remise")) / 100)
                                        Else
                                            .DL_PrixUnitaire = DossierTab.Rows(0).Item("AF_PrixAch")
                                        End If
                                    Else
                                        .DL_PrixUnitaire = DossierTab.Rows(0).Item("AF_PrixAch")
                                    End If
                                End If
                            Else
                                If DossierTab.Rows(0).Item("AF_Devise") <> 0 Then
                                    statistAdap = New OleDbDataAdapter("select * from P_DEVISE where cbIndice=" & DossierTab.Rows(0).Item("AF_Devise") & " And N_DeviseCot =(Select  N_DeviseCompte From P_DOSSIERCIAL) And D_Cours<> 0", OleSocieteConnect)
                                    statistDs = New DataSet
                                    statistAdap.Fill(statistDs)
                                    statistTab = statistDs.Tables(0)
                                    If statistTab.Rows.Count <> 0 Then
                                        If Convert.IsDBNull(DossierTab.Rows(0).Item("AF_Remise")) = False Then
                                            If DossierTab.Rows(0).Item("AF_Remise") <> 0 Then
                                                .DL_PrixUnitaire = (DossierTab.Rows(0).Item("AF_PrixDev") - ((DossierTab.Rows(0).Item("AF_PrixDev") * DossierTab.Rows(0).Item("AF_Remise")) / 100)) / statistTab.Rows(0).Item("D_Cours")
                                            Else
                                                .DL_PrixUnitaire = DossierTab.Rows(0).Item("AF_PrixDev") / statistTab.Rows(0).Item("D_Cours")
                                            End If
                                        Else
                                            .DL_PrixUnitaire = DossierTab.Rows(0).Item("AF_PrixDev") / statistTab.Rows(0).Item("D_Cours")
                                        End If
                                    Else
                                        If .Article.AR_PrixAchat <> 0 Then
                                            .DL_PrixUnitaire = .Article.AR_PrixAchat
                                        End If
                                    End If
                                Else
                                    If .Article.AR_PrixAchat <> 0 Then
                                        .DL_PrixUnitaire = .Article.AR_PrixAchat
                                    End If
                                End If
                            End If
                        Else
                            If .Article.AR_PrixAchat <> 0 Then
                                .DL_PrixUnitaire = .Article.AR_PrixAchat
                            End If
                        End If
                    End If
                End If
            End If
        End With
    End Sub
    Private Sub Creation_Ligne_Article(ByRef FormatDatefichier As String, Optional ByRef PieceArticle As String = "", Optional ByRef Punitaire As String = "", Optional ByRef IdentifiantArticle As String = "", Optional ByRef EnteteTyPeDocument As String = "")
        infosExport.Text = "Création de la Ligne du Mvt"
        Dim statistAdap As OleDbDataAdapter
        Dim statistDs As DataSet
        Dim statistTab As DataTable

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
                If Trim(LigneDesignationArticle) <> "" Then
                    .DL_Design = LigneDesignationArticle
                End If
                If Trim(LignePoidsNet) <> "" Then
                    If EstNumeric(Trim(LignePoidsNet), DecimalNomb, DecimalMone) = True Then
                        .DL_PoidsNet = CDbl(RenvoiMontant(Trim(LignePoidsNet), FormatQte, DecimalNomb, DecimalMone))
                    End If
                End If
                If Trim(LignePoidsBrut) <> "" Then
                    If EstNumeric(Trim(LignePoidsBrut), DecimalNomb, DecimalMone) = True Then
                        .DL_PoidsBrut = CDbl(RenvoiMontant(Trim(LignePoidsBrut), FormatQte, DecimalNomb, DecimalMone))
                    End If
                End If
                If Trim(LigneReference) <> "" Then
                    .DO_Ref = Trim(LigneReference)
                End If
                If Trim(LigneDatedeFabrication) <> "" Then
                    If Trim(LigneDatedeFabrication) <> "" Then
                        .LS_Fabrication = RenvoieDateValide(Trim(LigneDatedeFabrication), FormatDatefichier)
                    End If
                End If
                If Trim(LigneDatedePeremption) <> "" Then
                    If Trim(LigneDatedePeremption) <> "" Then
                        .LS_Peremption = RenvoieDateValide(Trim(LigneDatedePeremption), FormatDatefichier)
                    End If
                End If
                If Trim(LigneNSerieLot) <> "" Then
                    .LS_NoSerie = LigneNSerieLot
                End If
                .Valorisee = True
                If Trim(LigneCodeArticle) <> "" Then
                    If BaseCial.FactoryArticle.ExistReference(Trim(LigneCodeArticle)) = True Then
                        If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                            .SetDefaultArticle(BaseCial.FactoryArticle.ReadReference(Trim(LigneCodeArticle)), CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)))
                        End If
                    End If
                End If
                If EnteteTyPeDocument = "20" Then
                    If Trim(LigneCodeArticle) <> "" Then
                        If BaseCial.FactoryArticle.ExistReference(Trim(LigneCodeArticle)) = True Then
                            If Trim(IDDepotEntete) <> "" Then
                                If IsNumeric(Trim(IDDepotEntete)) = True Then
                                    statistAdap = New OleDbDataAdapter("select * from F_DEPOT where DE_No ='" & CInt(Trim(IDDepotEntete)) & "'", OleSocieteConnect)
                                    statistDs = New DataSet
                                    statistAdap.Fill(statistDs)
                                    statistTab = statistDs.Tables(0)
                                    If statistTab.Rows.Count <> 0 Then
                                        PriUnitaireEntrer(LigneDocument, CInt(Trim(IDDepotEntete)), Trim(LigneCodeArticle))
                                    End If
                                End If
                            End If
                            If Trim(EnteteIntituleDepot) <> "" Then
                                If BaseCial.FactoryDepot.ExistIntitule(Trim(EnteteIntituleDepot)) = True Then
                                    DossierAdap = New OleDbDataAdapter("select * from F_DEPOT where DE_Intitule ='" & Join(Split(Trim(EnteteIntituleDepot), "'"), "''") & "'", OleSocieteConnect)
                                    DossierDs = New DataSet
                                    DossierAdap.Fill(DossierDs)
                                    DossierTab = DossierDs.Tables(0)
                                    If DossierTab.Rows.Count <> 0 Then
                                        PriUnitaireEntrer(LigneDocument, DossierTab.Rows(0).Item("DE_No"), Trim(LigneCodeArticle))
                                    End If
                                ElseIf BaseCial.FactoryDepot.ExistIntitule(Trim(EnteteIntituleDepot)) = True Then
                                    DossierAdap = New OleDbDataAdapter("select * from F_DEPOT where DE_Intitule ='" & Join(Split(Trim(EnteteIntituleDepot), "'"), "''") & "'", OleSocieteConnect)
                                    DossierDs = New DataSet
                                    DossierAdap.Fill(DossierDs)
                                    DossierTab = DossierDs.Tables(0)
                                    If DossierTab.Rows.Count <> 0 Then
                                        PriUnitaireEntrer(LigneDocument, DossierTab.Rows(0).Item("DE_No"), Trim(LigneCodeArticle))
                                    End If
                                End If
                            End If
                        End If
                    End If
                    .Write()
                Else
                    If Punitaire = "oui" Then
                        .WriteDefault()
                    Else
                        .Write()
                    End If
                    If Trim(LignePrixUnitaire) <> "" Then
                        If EstNumeric(Trim(LignePrixUnitaire), DecimalNomb, DecimalMone) = True Then
                            .DL_PrixUnitaire = CDbl(RenvoiMontant(Trim(LignePrixUnitaire), FormatMnt, DecimalNomb, DecimalMone))
                        End If
                    End If
                End If
                If Trim(LignePoidsNet) <> "" Then
                    If EstNumeric(Trim(LignePoidsNet), DecimalNomb, DecimalMone) = True Then
                        .DL_PoidsNet = CDbl(RenvoiMontant(Trim(LignePoidsNet), FormatQte, DecimalNomb, DecimalMone))
                        .Write()
                    End If
                End If
                If Trim(LignePoidsBrut) <> "" Then
                    If EstNumeric(Trim(LignePoidsBrut), DecimalNomb, DecimalMone) = True Then
                        .DL_PoidsBrut = CDbl(RenvoiMontant(Trim(LignePoidsBrut), FormatQte, DecimalNomb, DecimalMone))
                        .Write()
                    End If
                End If
                If IsNothing(LigneDocument.Article) = False Then
                    ErreurJrn.WriteLine("Code article : " & Trim(LigneDocument.Article.AR_Ref) & " Créé Pour la pièce du fichier N° :" & Trim(EntetePieceInterne))
                Else
                    ErreurJrn.WriteLine("Code article : " & Trim("Vide") & " Créé Pour la pièce du fichier N° :" & Trim(EntetePieceInterne))
                End If
                'Traitement des Infos Libres
                Try
                    'If infoLigne.Count > 0 Then
                    '    While infoLigne.Count <> 0
                    '        OleAdaptaterDelete = New OleDbDataAdapter("select * From WIS_COL where Libelle='" & Trim(Datagridaffiche.Columns(infoLigne.Item(0)).Name) & "' And InfoLigne=True", OleConnenection)
                    '        OleDeleteDataset = New DataSet
                    '        OleAdaptaterDelete.Fill(OleDeleteDataset)
                    '        OledatableDelete = OleDeleteDataset.Tables(0)
                    '        If OledatableDelete.Rows.Count <> 0 Then
                    '            'L'info Libre Parametrée par l'utilisateur existe dans Sage
                    '            If LigneDocument.InfoLibre.Count <> 0 Then
                    '                If IsNothing(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) = False Then
                    '                    If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                    '                        statistAdap = New OleDbDataAdapter("select * from cbSysLibre where CB_File='F_DOCLIGNE' and CB_Name ='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "'", OleSocieteConnect)
                    '                        statistDs = New DataSet
                    '                        statistAdap.Fill(statistDs)
                    '                        statistTab = statistDs.Tables(0)
                    '                        If statistTab.Rows.Count <> 0 Then
                    '                            'Texte
                    '                            If statistTab.Rows(0).Item("CB_Type") = 9 Then
                    '                                OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Stock' And Ligne=True", OleConnenection)
                    '                                OleRecherDataset = New DataSet
                    '                                OleRecherAdapter.Fill(OleRecherDataset)
                    '                                OleRechDatable = OleRecherDataset.Tables(0)
                    '                                If OleRechDatable.Rows.Count <> 0 Then
                    '                                    If statistTab.Rows(0).Item("CB_Len") >= Strings.Len(Trim(OleRechDatable.Rows(0).Item("Correspond"))) Then
                    '                                        LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = Trim(OleRechDatable.Rows(0).Item("Correspond"))
                    '                                        LigneDocument.Write()
                    '                                    End If
                    '                                Else
                    '                                    If statistTab.Rows(0).Item("CB_Len") >= Strings.Len(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value)) Then
                    '                                        LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value)
                    '                                        LigneDocument.Write()
                    '                                    End If
                    '                                End If
                    '                            End If
                    '                            'Table
                    '                            If statistTab.Rows(0).Item("CB_Type") = 22 Then
                    '                                OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Stock' And Ligne=True", OleConnenection)
                    '                                OleRecherDataset = New DataSet
                    '                                OleRecherAdapter.Fill(OleRecherDataset)
                    '                                OleRechDatable = OleRecherDataset.Tables(0)
                    '                                If OleRechDatable.Rows.Count <> 0 Then
                    '                                    If statistTab.Rows(0).Item("CB_Len") > Strings.Len(Trim(OleRechDatable.Rows(0).Item("Correspond"))) Then
                    '                                        LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = Trim(OleRechDatable.Rows(0).Item("Correspond"))
                    '                                        LigneDocument.Write()
                    '                                    End If
                    '                                Else
                    '                                    If statistTab.Rows(0).Item("CB_Len") > Strings.Len(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value)) Then
                    '                                        LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value)
                    '                                        LigneDocument.Write()
                    '                                    End If
                    '                                End If
                    '                            End If
                    '                            'Montant
                    '                            If statistTab.Rows(0).Item("CB_Type") = 20 Then
                    '                                If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                    '                                    OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Stock' And Ligne=True", OleConnenection)
                    '                                    OleRecherDataset = New DataSet
                    '                                    OleRecherAdapter.Fill(OleRecherDataset)
                    '                                    OleRechDatable = OleRecherDataset.Tables(0)
                    '                                    If OleRechDatable.Rows.Count <> 0 Then
                    '                                        If EstNumeric(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone) = True Then
                    '                                            LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone))
                    '                                            LigneDocument.Write()
                    '                                        End If
                    '                                    Else
                    '                                        If EstNumeric(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), DecimalNomb, DecimalMone) = True Then
                    '                                            LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), DecimalNomb, DecimalMone))
                    '                                            LigneDocument.Write()
                    '                                        End If
                    '                                    End If
                    '                                End If
                    '                            End If
                    '                            'Valeur
                    '                            If statistTab.Rows(0).Item("CB_Type") = 7 Then
                    '                                If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                    '                                    OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Stock' And Ligne=True", OleConnenection)
                    '                                    OleRecherDataset = New DataSet
                    '                                    OleRecherAdapter.Fill(OleRecherDataset)
                    '                                    OleRechDatable = OleRecherDataset.Tables(0)
                    '                                    If OleRechDatable.Rows.Count <> 0 Then
                    '                                        If EstNumeric(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone) = True Then
                    '                                            LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone))
                    '                                            LigneDocument.Write()
                    '                                        End If
                    '                                    Else
                    '                                        If EstNumeric(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), DecimalNomb, DecimalMone) = True Then
                    '                                            LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), DecimalNomb, DecimalMone))
                    '                                            LigneDocument.Write()
                    '                                        End If
                    '                                    End If
                    '                                End If
                    '                            End If

                    '                            'Date Court
                    '                            If statistTab.Rows(0).Item("CB_Type") = 3 Then
                    '                                If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                    '                                    OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Stock' And Ligne=True", OleConnenection)
                    '                                    OleRecherDataset = New DataSet
                    '                                    OleRecherAdapter.Fill(OleRecherDataset)
                    '                                    OleRechDatable = OleRecherDataset.Tables(0)
                    '                                    If OleRechDatable.Rows.Count <> 0 Then
                    '                                        If Verificatdate(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then
                    '                                            LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = RenvoieDateValide(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier)
                    '                                            LigneDocument.Write()
                    '                                        End If
                    '                                    Else
                    '                                        If Verificatdate(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then
                    '                                            LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = RenvoieDateValide(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), FormatDatefichier)
                    '                                            LigneDocument.Write()
                    '                                        End If
                    '                                    End If
                    '                                End If
                    '                            End If
                    '                            'Date Longue
                    '                            If statistTab.Rows(0).Item("CB_Type") = 14 Then
                    '                                If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                    '                                    OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Stock' And Ligne=True", OleConnenection)
                    '                                    OleRecherDataset = New DataSet
                    '                                    OleRecherAdapter.Fill(OleRecherDataset)
                    '                                    OleRechDatable = OleRecherDataset.Tables(0)
                    '                                    If OleRechDatable.Rows.Count <> 0 Then
                    '                                        If Verificatdate(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then
                    '                                            LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = RenvoieDateValide(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier)
                    '                                            LigneDocument.Write()
                    '                                        End If
                    '                                    Else
                    '                                        If Verificatdate(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then
                    '                                            LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = RenvoieDateValide(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), FormatDatefichier)
                    '                                            LigneDocument.Write()
                    '                                        End If
                    '                                    End If
                    '                                End If
                    '                            End If
                    '                        Else
                    '                            If IsNothing(LigneDocument.Article) = False Then
                    '                                ErreurJrn.WriteLine("Code Article : " & Trim(LigneDocument.Article.AR_Ref) & " Impossible de traiter l'information libre :" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "  Il est inexistant dans la table CbsysLibre de Sage")
                    '                            Else
                    '                                ErreurJrn.WriteLine("Code Article : " & Trim("Vide") & " Impossible de traiter l'information libre :" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "  Il est inexistant dans la table CbsysLibre de Sage")
                    '                            End If
                    '                        End If
                    '                    End If
                    '                Else
                    '                    'nothing
                    '                End If
                    '            End If
                    '        End If
                    '        'L'info Libre Parametrée par l'utilisateur n'existe pas dans Sage
                    '        infoLigne.RemoveAt(0)
                    '    End While
                    'End If
                Catch ex As Exception
                    exceptionTrouve = True
                    If IsNothing(LigneDocument.Article) = False Then
                        ErreurJrn.WriteLine("Code Article : " & Trim(LigneDocument.Article.AR_Ref) & " Erreur de Création de L'information Libre , vérifiez la longueur de la chaine avec celle paramétrée/ ou la cohérence des données " & ex.Message)
                    Else
                        ErreurJrn.WriteLine("Code Article : " & Trim("Vide") & " Erreur de Création de L'information Libre , vérifiez la longueur de la chaine avec celle paramétrée/ ou la cohérence des données " & ex.Message)
                    End If
                End Try
                .WriteDefault()
            End With
        Catch ex As Exception
            exceptionTrouve = True
            ErreurJrn.WriteLine("Code Article : " & Trim(LigneCodeArticle) & " N°Pièce : " & EntetePieceInterne & " Erreur système de Création de l'article : " & ex.Message)
            ListePiece.Add(EntetePieceInterne)
        End Try
    End Sub
    Public RegardeStatut As Boolean = True
    Private Sub Verification_Parametrage(ByRef EnteteIntituleDepot As String, ByRef EntetePieceInterne As String, ByRef EnteteTyPeDocument As String, ByRef Document As IBODocumentStock3, ByRef infoListe As List(Of Integer), ByRef FormatDatefichier As String, ByRef PieceCreationDocument As Object, ByRef PieceAutomtique As Object, ByRef PieceArticle As String, ByRef Punitaire As String, ByRef IdentifiantArticle As String)
        Dim statistAdap As OleDbDataAdapter
        Dim statistDs As DataSet
        Dim statistTab As DataTable

        Dim DossierAdap As OleDbDataAdapter
        Dim DossierDs As DataSet
        Dim DossierTab As DataTable
        Dim FormatQte As Integer = 0
        Dim FormatMnt As Integer = 0
        infosExport.Refresh()
        infosExport.Text = "Vérification des Integrations!"
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
        If ChkPieceAuto.Checked = False Then
            If EnteteTyPeDocument = "20" Then
                If Trim(LigneQuantite) <> "" Then
                    If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                        If CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)) < 0 Then
                            ExisteLecture = False
                            ErreurJrn.WriteLine("< La Quantité : " & Trim(LigneQuantite) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " ne doit pas être négative >")
                        End If
                    End If
                End If
                If EntetePieceInterne <> "" Then
                    If BaseCial.FactoryDocumentStock.ExistPiece(DocumentType.DocumentTypeStockMouvIn, EntetePieceInterne) = True Then
                        ErreurJrn.WriteLine("Mouvement d'entrée  N° : " & EntetePieceInterne & " Existe déja ")
                        ExisteLecture = False
                    End If
                End If
            Else
                If EnteteTyPeDocument = "21" Then
                    If Trim(LigneQuantite) <> "" Then
                        If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                            If CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)) < 0 Then
                                ExisteLecture = False
                                ErreurJrn.WriteLine("< La Quantité : " & Trim(LigneQuantite) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " ne doit pas être négative >")
                            End If
                        End If
                    End If
                    If EntetePieceInterne <> "" Then
                        If BaseCial.FactoryDocumentStock.ExistPiece(DocumentType.DocumentTypeStockMouvIn, EntetePieceInterne) = True Then
                            ErreurJrn.WriteLine("Mouvement d'entrée  N° : " & EntetePieceInterne & " Existe déja ")
                            ExisteLecture = False
                        End If
                    End If
                End If
            End If
        End If

        If Trim(EnteteTyPeDocument) <> "" Then
            If Trim(EnteteTyPeDocument) <> "" Then
                If Trim(EnteteTyPeDocument) = "20" Then
                    If Trim(LigneQuantite) <> "" Then
                        If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                            If CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)) < 0 Then
                                ExisteLecture = False
                                ErreurJrn.WriteLine("< La Quantité : " & Trim(LigneQuantite) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " ne doit pas être négative >")
                            End If
                        End If
                    End If
                Else
                    If Trim(EnteteTyPeDocument) = "21" Then
                        If Trim(LigneQuantite) <> "" Then
                            If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                                If CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)) < 0 Then
                                    ExisteLecture = False
                                    ErreurJrn.WriteLine("< La Quantité : " & Trim(LigneQuantite) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " ne doit pas être négative >")
                                End If
                            End If
                        End If
                    Else
                        ErreurJrn.WriteLine("Le statut du document " & EnteteTyPeDocument & " dois être égal à 20:Entrée,21:Sortie : " & EntetePieceInterne & " le statut par défaut va être utilisé")
                    End If
                End If
            End If
        End If

        If Trim(EntetePieceInterne) <> "" And ChkPieceAuto.Checked = False Then
            ErreurJrn.WriteLine("Le N°Pièce du fichier  : " & EntetePieceInterne & " ne doit pas être vide ")
            ExisteLecture = False
        End If
        If Trim(EntetePlanAnalytique) <> "" Then
            If Trim(EnteteCodeAffaire) <> "" Then
                If BaseCpta.FactoryAnalytique.ExistIntitule(Trim(EntetePlanAnalytique)) = True Then
                    PlanAna = BaseCpta.FactoryAnalytique.ReadIntitule(Trim(EntetePlanAnalytique))
                    If BaseCpta.FactoryCompteA.ExistNumero(PlanAna, Trim(EnteteCodeAffaire)) = False Then
                        ExisteLecture = False
                        ErreurJrn.WriteLine("< Le Code de Section analytique : " & Trim(EnteteCodeAffaire) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'existe pas dans la table de paramétrage et dans Sage >")
                    End If
                Else
                    ExisteLecture = False
                    ErreurJrn.WriteLine("< Le Code du plan analytique : " & Trim(EntetePlanAnalytique) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'existe pas dans la table de paramétrage et dans Sage >")
                End If
            End If
        End If
        If EnteteTyPeDocument = "21" Then
            If Trim(IDDepotEntete) <> "" Then
                If Trim(LigneCodeArticle) <> "" Then
                    If BaseCial.FactoryArticle.ExistReference(Trim(LigneCodeArticle)) = True Then
                        If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                            DossierAdap = New OleDbDataAdapter("select * from F_ARTSTOCK where DE_No ='" & CInt(Trim(IDDepotEntete)) & "' And AR_Ref='" & Trim(LigneCodeArticle) & "'", OleSocieteConnect)
                            DossierDs = New DataSet
                            DossierAdap.Fill(DossierDs)
                            DossierTab = DossierDs.Tables(0)
                            If DossierTab.Rows.Count <> 0 Then
                                If Convert.IsDBNull(DossierTab.Rows(0).Item("AS_QteSto")) = False Then
                                    If Convert.IsDBNull(DossierTab.Rows(0).Item("AS_QtePrepa")) = False Then
                                        If (DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) < CDbl(Trim(LigneQuantite)) Then
                                            ExisteLecture = False
                                            ErreurJrn.WriteLine("Le Stock : " & (DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) & " dans le dépôt de Sortie : " & Trim(IDDepotEntete) & "   ne permet pas de sortir l'article : " & LigneCodeArticle & " de Quantité : " & CDbl(Trim(LigneQuantite)) & " ,  N° :" & Trim(EntetePieceInterne))
                                        Else
                                            If ListeStock.Count <> 0 Then
                                                Dim ExisteArtStock As Boolean = False
                                                For i As Integer = 0 To ListeStock.Count - 1
                                                    Dim StockListe() As String = Split(ListeStock.Item(i), ControlChars.Tab)
                                                    If StockListe(0) = CInt(Trim(IDDepotEntete)) And Trim(StockListe(1)) = Trim(LigneCodeArticle) Then
                                                        ExisteArtStock = True
                                                        If StockListe(2) < CDbl(Trim(LigneQuantite)) Then
                                                            ExisteLecture = False
                                                            ErreurJrn.WriteLine("Le Stock Calculé " & StockListe(2) & " dans le dépôt de Sortie : " & StockListe(0) & "   ne permet pas de sortir l'article : " & LigneCodeArticle & " de Quantité : " & CDbl(Trim(LigneQuantite)) & " ,  N° :" & Trim(EntetePieceInterne))
                                                        Else
                                                            ListeStock.RemoveAt(i)
                                                            ListeStock.Add(CInt(Trim(IDDepotEntete)) & ControlChars.Tab & Trim(LigneCodeArticle) & ControlChars.Tab & (StockListe(2) - CDbl(Trim(LigneQuantite))))
                                                            Exit For
                                                        End If
                                                    End If
                                                Next i
                                                If ExisteArtStock = False Then
                                                    ListeStock.Add(CInt(Trim(IDDepotEntete)) & ControlChars.Tab & Trim(LigneCodeArticle) & ControlChars.Tab & ((DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) - CDbl(Trim(LigneQuantite))))
                                                End If
                                            Else
                                                ListeStock.Add(CInt(Trim(IDDepotEntete)) & ControlChars.Tab & Trim(LigneCodeArticle) & ControlChars.Tab & ((DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) - CDbl(Trim(LigneQuantite))))
                                            End If
                                        End If
                                    Else
                                        If DossierTab.Rows(0).Item("AS_QteSto") < CDbl(Trim(LigneQuantite)) Then
                                            ExisteLecture = False
                                            ErreurJrn.WriteLine("Le Stock : " & DossierTab.Rows(0).Item("AS_QteSto") & " dans le dépôt de Sortie : " & Trim(IDDepotEntete) & "   ne permet pas de sortir l'article : " & LigneCodeArticle & " de Quantité : " & CDbl(Trim(LigneQuantite)) & " ,  N° :" & Trim(EntetePieceInterne))
                                        Else
                                            If ListeStock.Count <> 0 Then
                                                Dim ExisteArtStock As Boolean = False
                                                For i As Integer = 0 To ListeStock.Count - 1
                                                    Dim StockListe() As String = Split(ListeStock.Item(i), ControlChars.Tab)
                                                    If StockListe(0) = CInt(Trim(IDDepotEntete)) And Trim(StockListe(1)) = Trim(LigneCodeArticle) Then
                                                        ExisteArtStock = True
                                                        If StockListe(2) < CDbl(Trim(LigneQuantite)) Then
                                                            ExisteLecture = False
                                                            ErreurJrn.WriteLine("Le Stock Calculé " & StockListe(2) & " dans le dépôt de Sortie : " & StockListe(0) & "   ne permet pas de sortir l'article : " & LigneCodeArticle & " de Quantité : " & CDbl(Trim(LigneQuantite)) & " ,  N° :" & Trim(EntetePieceInterne))
                                                        Else
                                                            ListeStock.RemoveAt(i)
                                                            ListeStock.Add(CInt(Trim(IDDepotEntete)) & ControlChars.Tab & Trim(LigneCodeArticle) & ControlChars.Tab & (StockListe(2) - CDbl(Trim(LigneQuantite))))
                                                            Exit For
                                                        End If
                                                    End If
                                                Next i
                                                If ExisteArtStock = False Then
                                                    ListeStock.Add(CInt(Trim(IDDepotEntete)) & ControlChars.Tab & Trim(LigneCodeArticle) & ControlChars.Tab & (DossierTab.Rows(0).Item("AS_QteSto") - CDbl(Trim(LigneQuantite))))
                                                End If
                                            Else
                                                ListeStock.Add(CInt(Trim(IDDepotEntete)) & ControlChars.Tab & Trim(LigneCodeArticle) & ControlChars.Tab & (DossierTab.Rows(0).Item("AS_QteSto") - CDbl(Trim(LigneQuantite))))
                                            End If
                                        End If
                                    End If
                                Else
                                    ExisteLecture = False
                                    ErreurJrn.WriteLine("Le Stock dans le dépôt de Sortie : " & Trim(IDDepotEntete) & "  est NULL et ne permet pas de sortir l'article : " & LigneCodeArticle & " ,  N° :" & Trim(EntetePieceInterne))
                                End If
                            Else
                                ExisteLecture = False
                                ErreurJrn.WriteLine("Le Stock dans le dépôt de Sortie : " & Trim(IDDepotEntete) & "   ne permet pas de sortir l'article : " & LigneCodeArticle & " ,  N° :" & Trim(EntetePieceInterne))
                            End If
                        End If
                    End If
                End If
            End If

            If Trim(LigneCodeArticle) <> "" Then
                If BaseCial.FactoryArticle.ExistReference(Trim(LigneCodeArticle)) = True Then
                    If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                        If Trim(EnteteIntituleDepot) <> "" Then
                            If BaseCial.FactoryDepot.ExistIntitule(Trim(EnteteIntituleDepot)) = False Then
                                If BaseCial.FactoryDepot.ExistIntitule(Trim(EnteteIntituleDepot)) = True Then
                                    EnteteIntituleDepot = Trim(EnteteIntituleDepot)
                                End If
                            End If
                            If BaseCial.FactoryDepot.ExistIntitule(Trim(EnteteIntituleDepot)) = True Then
                                DossierAdap = New OleDbDataAdapter("select * from F_ARTSTOCK where DE_No =(Select DE_No from F_DEPOT WHERE DE_Intitule='" & Trim(EnteteIntituleDepot) & "') And AR_Ref='" & Trim(LigneCodeArticle) & "'", OleSocieteConnect)
                                DossierDs = New DataSet
                                DossierAdap.Fill(DossierDs)
                                DossierTab = DossierDs.Tables(0)
                                If DossierTab.Rows.Count <> 0 Then
                                    If Convert.IsDBNull(DossierTab.Rows(0).Item("AS_QteSto")) = False Then
                                        If Convert.IsDBNull(DossierTab.Rows(0).Item("AS_QtePrepa")) = False Then
                                            If (DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) < CDbl(Trim(LigneQuantite)) Then
                                                ExisteLecture = False
                                                ErreurJrn.WriteLine("Le Stock : " & (DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) & " dans le dépôt de Sortie : " & Trim(EnteteIntituleDepot) & "   ne permet pas de sortir l'article : " & LigneCodeArticle & " de Quantité : " & CDbl(Trim(LigneQuantite)) & " ,  N° :" & Trim(EntetePieceInterne))
                                            Else
                                                If ListeStock.Count <> 0 Then
                                                    Dim ExisteArtStock As Boolean = False
                                                    For i As Integer = 0 To ListeStock.Count - 1
                                                        Dim StockListe() As String = Split(ListeStock.Item(i), ControlChars.Tab)
                                                        If StockListe(0) = Trim(EnteteIntituleDepot) And Trim(StockListe(1)) = Trim(LigneCodeArticle) Then
                                                            ExisteArtStock = True
                                                            If StockListe(2) < CDbl(Trim(LigneQuantite)) Then
                                                                ExisteLecture = False
                                                                ErreurJrn.WriteLine("Le Stock Calculé " & StockListe(2) & " dans le dépôt de Sortie : " & StockListe(0) & "   ne permet pas de sortir l'article : " & LigneCodeArticle & " de Quantité : " & CDbl(Trim(LigneQuantite)) & " ,  N° :" & Trim(EntetePieceInterne))
                                                            Else
                                                                ListeStock.RemoveAt(i)
                                                                ListeStock.Add(Trim(EnteteIntituleDepot) & ControlChars.Tab & Trim(LigneCodeArticle) & ControlChars.Tab & (StockListe(2) - CDbl(Trim(LigneQuantite))))
                                                                Exit For
                                                            End If
                                                        End If
                                                    Next i
                                                    If ExisteArtStock = False Then
                                                        ListeStock.Add(Trim(EnteteIntituleDepot) & ControlChars.Tab & Trim(LigneCodeArticle) & ControlChars.Tab & ((DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) - CDbl(Trim(LigneQuantite))))
                                                    End If
                                                Else
                                                    ListeStock.Add(Trim(EnteteIntituleDepot) & ControlChars.Tab & Trim(LigneCodeArticle) & ControlChars.Tab & ((DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) - CDbl(Trim(LigneQuantite))))
                                                End If
                                            End If
                                        Else
                                            If DossierTab.Rows(0).Item("AS_QteSto") < CDbl(Trim(LigneQuantite)) Then
                                                ExisteLecture = False
                                                ErreurJrn.WriteLine("Le Stock : " & DossierTab.Rows(0).Item("AS_QteSto") & " dans le dépôt de Sortie : " & Trim(EnteteIntituleDepot) & "   ne permet pas de sortir l'article : " & LigneCodeArticle & " de Quantité : " & CDbl(Trim(LigneQuantite)) & " ,  N° :" & Trim(EntetePieceInterne))
                                            Else
                                                If ListeStock.Count <> 0 Then
                                                    Dim ExisteArtStock As Boolean = False
                                                    For i As Integer = 0 To ListeStock.Count - 1
                                                        Dim StockListe() As String = Split(ListeStock.Item(i), ControlChars.Tab)
                                                        If StockListe(0) = Trim(EnteteIntituleDepot) And Trim(StockListe(1)) = Trim(LigneCodeArticle) Then
                                                            ExisteArtStock = True
                                                            If StockListe(2) < CDbl(Trim(LigneQuantite)) Then
                                                                ExisteLecture = False
                                                                ErreurJrn.WriteLine("Le Stock Calculé " & StockListe(2) & " dans le dépôt de Sortie : " & StockListe(0) & "   ne permet pas de sortir l'article : " & LigneCodeArticle & " de Quantité : " & CDbl(Trim(LigneQuantite)) & " ,  N° :" & Trim(EntetePieceInterne))
                                                            Else
                                                                ListeStock.RemoveAt(i)
                                                                ListeStock.Add(Trim(EnteteIntituleDepot) & ControlChars.Tab & Trim(LigneCodeArticle) & ControlChars.Tab & (StockListe(2) - CDbl(Trim(LigneQuantite))))
                                                                Exit For
                                                            End If
                                                        End If
                                                    Next i
                                                    If ExisteArtStock = False Then
                                                        ListeStock.Add(Trim(EnteteIntituleDepot) & ControlChars.Tab & Trim(LigneCodeArticle) & ControlChars.Tab & (DossierTab.Rows(0).Item("AS_QteSto") - CDbl(Trim(LigneQuantite))))
                                                    End If
                                                Else
                                                    ListeStock.Add(Trim(EnteteIntituleDepot) & ControlChars.Tab & Trim(LigneCodeArticle) & ControlChars.Tab & (DossierTab.Rows(0).Item("AS_QteSto") - CDbl(Trim(LigneQuantite))))
                                                End If
                                            End If
                                        End If
                                    Else
                                        ExisteLecture = False
                                        ErreurJrn.WriteLine("Le Stock dans le dépôt de Sortie : " & Trim(EnteteIntituleDepot) & "  est NULL et ne permet pas de sortir l'article : " & LigneCodeArticle & " ,  N° :" & Trim(EntetePieceInterne))
                                    End If
                                Else
                                    ExisteLecture = False
                                    ErreurJrn.WriteLine("Le Stock dans le dépôt de Sortie : " & Trim(EnteteIntituleDepot) & "   ne permet pas de sortir l'article : " & LigneCodeArticle & " ,  N° :" & Trim(EntetePieceInterne))
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        If Trim(LigneQuantite) <> "" Then
            If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                If CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)) <> 0 Then
                    If Trim(EnteteIntituleDepot) <> "" Then
                        If BaseCial.FactoryDepot.ExistIntitule(Trim(EnteteIntituleDepot)) = False Then
                        Else
                            If BaseCial.FactoryDepot.ExistIntitule(Trim(EnteteIntituleDepot)) = False Then
                                ExisteLecture = False
                                ErreurJrn.WriteLine("< L'Intitulé dépôt : " & Trim(EnteteIntituleDepot) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'existe pas dans la table de paramétrage et dans Sage >")
                            End If
                        End If
                    End If
                    If IsNumeric(Trim(IDDepotEntete)) = True Then
                        ExisteLecture = False
                        ErreurJrn.WriteLine("< L'ID dépôt : " & Trim(IDDepotEntete) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'existe pas dans la table de paramétrage et dans Sage >")
                    End If
                End If
            End If
        Else
            If BaseCial.FactoryDepot.ExistIntitule(Trim(EnteteIntituleDepot)) = False Then
            Else
                If BaseCial.FactoryDepot.ExistIntitule(Trim(EnteteIntituleDepot)) = False Then
                    ExisteLecture = False
                    ErreurJrn.WriteLine("< L'Intitulé dépôt : " & Trim(EnteteIntituleDepot) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'existe pas dans la table de paramétrage et dans Sage >")
                End If
            End If
            If IsNumeric(Trim(IDDepotEntete)) = True Then
                ExisteLecture = False
                ErreurJrn.WriteLine("< L'ID dépôt : " & IDDepotEntete & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " existant dans le fichier n'est pas numérique>")
            End If
        End If

        If Trim(EnteteDateDocument) <> "" Then
            'If Verificatdate(Trim(EnteteDateDocument), FormatDatefichier, "Date de Document") = True Then
            'Else
            '    ExisteLecture = False
            'End If
        End If
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
                If BaseCial.FactorySoucheStock.ExistIntitule(Trim(EnteteSoucheDocument)) = False Then
                    ExisteLecture = False
                    ErreurJrn.WriteLine("< L'Intitulé de la  Souche du Document : " & Trim(EnteteSoucheDocument) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'existe pas dans la table de paramétrage et dans Sage >")
                End If
            End If
        End If
        If Trim(LignePoidsNet) <> "" Then
            If EstNumeric(Trim(LignePoidsNet), DecimalNomb, DecimalMone) = True Then
            Else
                ExisteLecture = False
                ErreurJrn.WriteLine("< Le poids Net : " & Trim(LignePoidsNet) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'est pas numérique >")
            End If
        End If
        If Trim(LignePoidsBrut) <> "" Then
            If EstNumeric(Trim(LignePoidsBrut), DecimalNomb, DecimalMone) = True Then
            Else
                ExisteLecture = False
                ErreurJrn.WriteLine("< Le poids brut : " & Trim(LignePoidsBrut) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'est pas numérique >")
            End If
        End If
        If Trim(LignePrixUnitaire) <> "" Then
            If EstNumeric(Trim(LignePrixUnitaire), DecimalNomb, DecimalMone) = True Then
            Else
                ExisteLecture = False
                ErreurJrn.WriteLine("< Le prix unitaire : " & Trim(LignePrixUnitaire) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'est pas numérique >")
            End If
        End If
        If Trim(LigneDatedeFabrication) <> "" Then
            If IsDate(Trim(LigneDatedeFabrication)) = True Then
            Else
                ExisteLecture = False
                ErreurJrn.WriteLine("< La date de Fabrication : " & Trim(LigneDatedeFabrication) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'est pas au format date >")
            End If
        End If
        If Trim(LigneDatedePeremption) <> "" Then
            If IsDate(Trim(LigneDatedePeremption)) = True Then
            Else
                ExisteLecture = False
                ErreurJrn.WriteLine("< La date de Peremption : " & Trim(LigneDatedePeremption) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'est pas au format date >")
            End If
        End If
        If Trim(LigneQuantite) <> "" Then
            If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
            Else
                ExisteLecture = False
                ErreurJrn.WriteLine("< La Quantité : " & Trim(LigneQuantite) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'est pas numérique >")
            End If
        End If
        If Trim(LigneCodeArticle) <> "" Then
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
                End If
            Else
                If Trim(LigneQuantite) <> "" Then
                    ExisteLecture = False
                    ErreurJrn.WriteLine("< La quantité :" & Trim(LigneQuantite) & " ne doit pas être renseignée pour La Référence Article : " & Trim(LigneCodeArticle) & "vide - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " >")
                End If
                If Trim(LignePrixUnitaire) <> "" Then
                    ExisteLecture = False
                    ErreurJrn.WriteLine("< Le prix unitaire :" & Trim(LignePrixUnitaire) & " ne doit pas être renseignée pour La Référence Article : " & Trim(LigneCodeArticle) & "vide - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " >")
                End If
            End If
        Else
            If Trim(LigneQuantite) <> "" Then
                ExisteLecture = False
                ErreurJrn.WriteLine("< La quantité :" & Trim(LigneQuantite) & " ne doit pas être renseignée pour La Référence Article : " & Trim(LigneCodeArticle) & "vide - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " >")
            End If
            If Trim(LignePrixUnitaire) <> "" Then
                ExisteLecture = False
                ErreurJrn.WriteLine("< Le prix unitaire :" & Trim(LignePrixUnitaire) & " ne doit pas être renseignée pour La Référence Article : " & Trim(LigneCodeArticle) & "vide - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " >")
            End If
        End If

        'Traitement des Infos Libres
        Try
            'If infoLigne.Count > 0 Then
            '    While infoLigne.Count <> 0
            '        OleAdaptaterDelete = New OleDbDataAdapter("select * From WIS_COL where Libelle='" & Trim(Datagridaffiche.Columns(infoLigne.Item(0)).Name) & "' And InfoLigne=True", OleConnenection)
            '        OleDeleteDataset = New DataSet
            '        OleAdaptaterDelete.Fill(OleDeleteDataset)
            '        OledatableDelete = OleDeleteDataset.Tables(0)
            '        If OledatableDelete.Rows.Count <> 0 Then
            '            'L'info Libre Parametrée par l'utilisateur existe dans Sage
            '            If IsNothing(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) = False Then
            '                If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
            '                    statistAdap = New OleDbDataAdapter("select * from cbSysLibre where CB_File='F_DOCLIGNE' and CB_Name ='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "'", OleSocieteConnect)
            '                    statistDs = New DataSet
            '                    statistAdap.Fill(statistDs)
            '                    statistTab = statistDs.Tables(0)
            '                    If statistTab.Rows.Count <> 0 Then
            '                        'Texte
            '                        If statistTab.Rows(0).Item("CB_Type") = 9 Then
            '                            OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Stock' And Ligne=True", OleConnenection)
            '                            OleRecherDataset = New DataSet
            '                            OleRecherAdapter.Fill(OleRecherDataset)
            '                            OleRechDatable = OleRecherDataset.Tables(0)
            '                            If OleRechDatable.Rows.Count <> 0 Then
            '                                If statistTab.Rows(0).Item("CB_Len") >= Strings.Len(Trim(OleRechDatable.Rows(0).Item("Correspond"))) Then

            '                                Else
            '                                    ExisteLecture = False
            '                                    ErreurJrn.WriteLine("La Longueur de l'info libre : " & Datagridaffiche.Columns(infoLigne.Item(0)).Name & " est inférieur au valeur adressée: " & Trim(OleRechDatable.Rows(0).Item("Correspond")) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
            '                                End If
            '                            Else
            '                                If statistTab.Rows(0).Item("CB_Len") >= Strings.Len(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value)) Then

            '                                Else
            '                                    ExisteLecture = False
            '                                    ErreurJrn.WriteLine("La Longueur de l'info libre : " & Datagridaffiche.Columns(infoLigne.Item(0)).Name & " est inférieur au valeur adressée: " & Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
            '                                End If
            '                            End If
            '                        End If
            '                        'Table
            '                        If statistTab.Rows(0).Item("CB_Type") = 22 Then
            '                            OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Stock' And Ligne=True", OleConnenection)
            '                            OleRecherDataset = New DataSet
            '                            OleRecherAdapter.Fill(OleRecherDataset)
            '                            OleRechDatable = OleRecherDataset.Tables(0)
            '                            If OleRechDatable.Rows.Count <> 0 Then
            '                                If statistTab.Rows(0).Item("CB_Len") > Strings.Len(Trim(OleRechDatable.Rows(0).Item("Correspond"))) Then

            '                                Else
            '                                    ExisteLecture = False
            '                                    ErreurJrn.WriteLine("La Longueur (21) de l'info libre de type table : " & Datagridaffiche.Columns(infoLigne.Item(0)).Name & " est inférieur au valeur adressée: " & Trim(OleRechDatable.Rows(0).Item("Correspond")) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
            '                                End If
            '                            Else
            '                                If statistTab.Rows(0).Item("CB_Len") > Strings.Len(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value)) Then

            '                                Else
            '                                    ExisteLecture = False
            '                                    ErreurJrn.WriteLine("La Longueur (21) de l'info libre de type table : " & Datagridaffiche.Columns(infoLigne.Item(0)).Name & " est inférieur au valeur adressée: " & Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
            '                                End If
            '                            End If
            '                        End If
            '                        'Montant
            '                        If statistTab.Rows(0).Item("CB_Type") = 20 Then
            '                            If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
            '                                OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Stock' And Ligne=True", OleConnenection)
            '                                OleRecherDataset = New DataSet
            '                                OleRecherAdapter.Fill(OleRecherDataset)
            '                                OleRechDatable = OleRecherDataset.Tables(0)
            '                                If OleRechDatable.Rows.Count <> 0 Then
            '                                    If EstNumeric(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone) = True Then

            '                                    Else
            '                                        ExisteLecture = False
            '                                        ErreurJrn.WriteLine("La valeur entrée : " & Trim(OleRechDatable.Rows(0).Item("Correspond")) & "doit être de type numérique pour l'info libre : " & Datagridaffiche.Columns(infoLigne.Item(0)).Name & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
            '                                    End If
            '                                Else
            '                                    If EstNumeric(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), DecimalNomb, DecimalMone) = True Then

            '                                    Else
            '                                        ExisteLecture = False
            '                                        ErreurJrn.WriteLine("La valeur entrée : " & Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) & "doit être de type numérique pour l'info libre : " & Datagridaffiche.Columns(infoLigne.Item(0)).Name & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
            '                                    End If
            '                                End If
            '                            End If
            '                        End If
            '                        'Valeur
            '                        If statistTab.Rows(0).Item("CB_Type") = 7 Then
            '                            If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
            '                                OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Stock' And Ligne=True", OleConnenection)
            '                                OleRecherDataset = New DataSet
            '                                OleRecherAdapter.Fill(OleRecherDataset)
            '                                OleRechDatable = OleRecherDataset.Tables(0)
            '                                If OleRechDatable.Rows.Count <> 0 Then
            '                                    If EstNumeric(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone) = True Then

            '                                    Else
            '                                        ExisteLecture = False
            '                                        ErreurJrn.WriteLine("La valeur entrée : " & Trim(OleRechDatable.Rows(0).Item("Correspond")) & "doit être de type numérique pour l'info libre : " & Datagridaffiche.Columns(infoLigne.Item(0)).Name & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
            '                                    End If
            '                                Else
            '                                    If EstNumeric(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), DecimalNomb, DecimalMone) = True Then

            '                                    Else
            '                                        ExisteLecture = False
            '                                        ErreurJrn.WriteLine("La valeur entrée : " & Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) & "doit être de type numérique pour l'info libre : " & Datagridaffiche.Columns(infoLigne.Item(0)).Name & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
            '                                    End If
            '                                End If
            '                            End If
            '                        End If

            '                        'Date Court
            '                        If statistTab.Rows(0).Item("CB_Type") = 3 Then
            '                            If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
            '                                OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Stock' And Ligne=True", OleConnenection)
            '                                OleRecherDataset = New DataSet
            '                                OleRecherAdapter.Fill(OleRecherDataset)
            '                                OleRechDatable = OleRecherDataset.Tables(0)
            '                                If OleRechDatable.Rows.Count <> 0 Then
            '                                    If Verificatdate(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then

            '                                    Else
            '                                        ExisteLecture = False
            '                                    End If
            '                                Else
            '                                    If Verificatdate(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then
            '                                    Else
            '                                        ExisteLecture = False
            '                                    End If
            '                                End If
            '                            End If
            '                        End If
            '                        'Date Longue
            '                        If statistTab.Rows(0).Item("CB_Type") = 14 Then
            '                            If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
            '                                OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Stock' And Ligne=True", OleConnenection)
            '                                OleRecherDataset = New DataSet
            '                                OleRecherAdapter.Fill(OleRecherDataset)
            '                                OleRechDatable = OleRecherDataset.Tables(0)
            '                                If OleRechDatable.Rows.Count <> 0 Then
            '                                    If Verificatdate(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then

            '                                    Else
            '                                        ExisteLecture = False
            '                                    End If
            '                                Else
            '                                    If Verificatdate(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), FormatDatefichier, Datagridaffiche.Columns(infoLigne.Item(0)).Name) = True Then
            '                                    Else
            '                                        ExisteLecture = False
            '                                    End If
            '                                End If
            '                            End If
            '                        End If
            '                    Else
            '                        ExisteLecture = False
            '                        ErreurJrn.WriteLine(" Impossible de traiter l'information libre en Ligne de Document :" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "  Il est inexistant dans la table CbsysLibre de Sage")
            '                    End If
            '                End If
            '            Else
            '                'nothing
            '            End If
            '        Else
            '            ExisteLecture = False
            '            ErreurJrn.WriteLine(" Impossible de traiter l'information libre en Ligne de Document :" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "  Il est inexistant dans la table de Paramétrage")
            '        End If
            '        'L'info Libre Parametrée par l'utilisateur n'existe pas dans Sage
            '        infoLigne.RemoveAt(0)
            '    End While
            'End If
        Catch ex As Exception
            exceptionTrouve = True
            ExisteLecture = False
            ErreurJrn.WriteLine(" Erreur de Création de L'information Libre Ligne Document " & ex.Message & ", vérifiez la longueur de la chaine avec celle paramétrée/ ou la cohérence des données")
        End Try
        'Traitement des Infos Libres
        Try
            'If infoListe.Count > 0 Then
            '    While infoListe.Count <> 0
            '        OleAdaptaterDelete = New OleDbDataAdapter("select * From WIS_COL where Libelle='" & Trim(Datagridaffiche.Columns(infoListe.Item(0)).Name) & "' And Libre=True", OleConnenection)
            '        OleDeleteDataset = New DataSet
            '        OleAdaptaterDelete.Fill(OleDeleteDataset)
            '        OledatableDelete = OleDeleteDataset.Tables(0)
            '        If OledatableDelete.Rows.Count <> 0 Then
            '            'L'info Libre Parametrée par l'utilisateur existe dans Sage
            '            If IsNothing(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) = False Then
            '                If Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) <> "" Then
            '                    statistAdap = New OleDbDataAdapter("select * from cbSysLibre where CB_File='F_DOCENTETE' and CB_Name ='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "'", OleSocieteConnect)
            '                    statistDs = New DataSet
            '                    statistAdap.Fill(statistDs)
            '                    statistTab = statistDs.Tables(0)
            '                    If statistTab.Rows.Count <> 0 Then
            '                        'Texte
            '                        If statistTab.Rows(0).Item("CB_Type") = 9 Then
            '                            OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Stock' And Entete=True", OleConnenection)
            '                            OleRecherDataset = New DataSet
            '                            OleRecherAdapter.Fill(OleRecherDataset)
            '                            OleRechDatable = OleRecherDataset.Tables(0)
            '                            If OleRechDatable.Rows.Count <> 0 Then
            '                                If statistTab.Rows(0).Item("CB_Len") >= Strings.Len(Trim(OleRechDatable.Rows(0).Item("Correspond"))) Then

            '                                Else
            '                                    ExisteLecture = False
            '                                    ErreurJrn.WriteLine("La Longueur de l'info libre : " & Datagridaffiche.Columns(infoListe.Item(0)).Name & " est inférieur au valeur adressée: " & Trim(OleRechDatable.Rows(0).Item("Correspond")) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
            '                                End If
            '                            Else
            '                                If statistTab.Rows(0).Item("CB_Len") >= Strings.Len(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value)) Then

            '                                Else
            '                                    ExisteLecture = False
            '                                    ErreurJrn.WriteLine("La Longueur de l'info libre : " & Datagridaffiche.Columns(infoListe.Item(0)).Name & " est inférieur au valeur adressée: " & Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
            '                                End If
            '                            End If
            '                        End If
            '                        'Table
            '                        If statistTab.Rows(0).Item("CB_Type") = 22 Then
            '                            OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Stock' And Entete=True", OleConnenection)
            '                            OleRecherDataset = New DataSet
            '                            OleRecherAdapter.Fill(OleRecherDataset)
            '                            OleRechDatable = OleRecherDataset.Tables(0)
            '                            If OleRechDatable.Rows.Count <> 0 Then
            '                                If statistTab.Rows(0).Item("CB_Len") > Strings.Len(Trim(OleRechDatable.Rows(0).Item("Correspond"))) Then

            '                                Else
            '                                    ExisteLecture = False
            '                                    ErreurJrn.WriteLine("La Longueur (21) de l'info libre de type table : " & Datagridaffiche.Columns(infoListe.Item(0)).Name & " est inférieur au valeur adressée: " & Trim(OleRechDatable.Rows(0).Item("Correspond")) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
            '                                End If
            '                            Else
            '                                If statistTab.Rows(0).Item("CB_Len") > Strings.Len(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value)) Then

            '                                Else
            '                                    ExisteLecture = False
            '                                    ErreurJrn.WriteLine("La Longueur (21) de l'info libre de type table : " & Datagridaffiche.Columns(infoListe.Item(0)).Name & " est inférieur au valeur adressée: " & Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
            '                                End If
            '                            End If
            '                        End If
            '                        'Montant
            '                        If statistTab.Rows(0).Item("CB_Type") = 20 Then
            '                            If Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) <> "" Then
            '                                OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Stock' And Entete=True", OleConnenection)
            '                                OleRecherDataset = New DataSet
            '                                OleRecherAdapter.Fill(OleRecherDataset)
            '                                OleRechDatable = OleRecherDataset.Tables(0)
            '                                If OleRechDatable.Rows.Count <> 0 Then
            '                                    If EstNumeric(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone) = True Then

            '                                    Else
            '                                        ExisteLecture = False
            '                                        ErreurJrn.WriteLine("La valeur entrée : " & Trim(OleRechDatable.Rows(0).Item("Correspond")) & "doit être de type numérique pour l'info libre : " & Datagridaffiche.Columns(infoLigne.Item(0)).Name & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
            '                                    End If
            '                                Else
            '                                    If EstNumeric(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), DecimalNomb, DecimalMone) = True Then

            '                                    Else
            '                                        ExisteLecture = False
            '                                        ErreurJrn.WriteLine("La valeur entrée : " & Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) & "doit être de type numérique pour l'info libre : " & Datagridaffiche.Columns(infoLigne.Item(0)).Name & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
            '                                    End If
            '                                End If
            '                            End If
            '                        End If
            '                        'Valeur
            '                        If statistTab.Rows(0).Item("CB_Type") = 7 Then
            '                            If Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) <> "" Then
            '                                OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Stock' And Entete=True", OleConnenection)
            '                                OleRecherDataset = New DataSet
            '                                OleRecherAdapter.Fill(OleRecherDataset)
            '                                OleRechDatable = OleRecherDataset.Tables(0)
            '                                If OleRechDatable.Rows.Count <> 0 Then
            '                                    If EstNumeric(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone) = True Then

            '                                    Else
            '                                        ExisteLecture = False
            '                                        ErreurJrn.WriteLine("La valeur entrée : " & Trim(OleRechDatable.Rows(0).Item("Correspond")) & "doit être de type numérique pour l'info libre : " & Datagridaffiche.Columns(infoLigne.Item(0)).Name & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
            '                                    End If
            '                                Else
            '                                    If EstNumeric(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), DecimalNomb, DecimalMone) = True Then

            '                                    Else
            '                                        ExisteLecture = False
            '                                        ErreurJrn.WriteLine("La valeur entrée : " & Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) & "doit être de type numérique pour l'info libre : " & Datagridaffiche.Columns(infoLigne.Item(0)).Name & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne))
            '                                    End If
            '                                End If
            '                            End If
            '                        End If

            '                        'Date Court
            '                        If statistTab.Rows(0).Item("CB_Type") = 3 Then
            '                            If Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) <> "" Then
            '                                OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Stock' And Entete=True", OleConnenection)
            '                                OleRecherDataset = New DataSet
            '                                OleRecherAdapter.Fill(OleRecherDataset)
            '                                OleRechDatable = OleRecherDataset.Tables(0)
            '                                If OleRechDatable.Rows.Count <> 0 Then
            '                                    If Verificatdate(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier, Datagridaffiche.Columns(infoListe.Item(0)).Name) = True Then

            '                                    Else
            '                                        ExisteLecture = False
            '                                    End If
            '                                Else
            '                                    If Verificatdate(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), FormatDatefichier, Datagridaffiche.Columns(infoListe.Item(0)).Name) = True Then
            '                                    Else
            '                                        ExisteLecture = False
            '                                    End If
            '                                End If
            '                            End If
            '                        End If
            '                        'Date Longue
            '                        If statistTab.Rows(0).Item("CB_Type") = 14 Then
            '                            If Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) <> "" Then
            '                                OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document Stock' And Entete=True", OleConnenection)
            '                                OleRecherDataset = New DataSet
            '                                OleRecherAdapter.Fill(OleRecherDataset)
            '                                OleRechDatable = OleRecherDataset.Tables(0)
            '                                If OleRechDatable.Rows.Count <> 0 Then
            '                                    If Verificatdate(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier, Datagridaffiche.Columns(infoListe.Item(0)).Name) = True Then

            '                                    Else
            '                                        ExisteLecture = False
            '                                    End If
            '                                Else
            '                                    If Verificatdate(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), FormatDatefichier, Datagridaffiche.Columns(infoListe.Item(0)).Name) = True Then
            '                                    Else
            '                                        ExisteLecture = False
            '                                    End If
            '                                End If
            '                            End If
            '                        End If
            '                    Else
            '                        ExisteLecture = False
            '                        ErreurJrn.WriteLine(" Impossible de traiter l'information libre en Entête de Document :" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "  Il est inexistant dans la table CbsysLibre de Sage")
            '                    End If
            '                End If
            '            Else
            '                'nothing
            '            End If
            '        Else
            '            ExisteLecture = False
            '            ErreurJrn.WriteLine(" Impossible de traiter l'information libre en Entête de Document :" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "  Il est inexistant dans la table de Paramétrage")
            '        End If
            '        'L'info Libre Parametrée par l'utilisateur n'existe pas dans Sage
            '        infoListe.RemoveAt(0)
            '    End While
            'End If
        Catch ex As Exception
            ExisteLecture = False
            exceptionTrouve = True
            ErreurJrn.WriteLine(" Erreur de Création de L'information Libre Entête de Document " & ex.Message & " , vérifiez la longueur de la chaine avec celle paramétrée/ ou la cohérence des données")
        End Try
    End Sub
    Public Sub AperçuElement(ByVal Chemin As String, Optional ByVal EstExecution As String = "")
        Dim OleAdaptaterschema As OleDbDataAdapter
        Dim OleSchemaDataset As DataSet
        Dim OledatableSchema As DataTable
        Dim ArtAdaptater As OleDbDataAdapter
        Dim ArtDataset As DataSet
        Dim Artdatatable As DataTable
        Dim CptaAdaptater As OleDbDataAdapter
        Dim CptaDataset As DataSet
        Dim Cptadatatable As DataTable

        Dim iline As Integer = 0
        Dim iLigne As Integer = 0
        Dim iDetaiLigne As Integer = 0
        Dim NbDetaiLigne As Integer = 0
        Dim NbLigne As Integer = 0
        Dim NbEntete As Integer = 0
        Dim Relation As String = ""
        Dim Statut As Boolean = False
        RegardeStatut = True
        OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='VST' AND ENTETE=true ORDER BY ORDRE", OleConnenection)
        OleSchemaDataset = New DataSet
        OleAdaptaterschema.Fill(OleSchemaDataset)
        OledatableSchema = OleSchemaDataset.Tables(0)

        If EstExecution <> "" Then
            Try
                infosExport.Text = "Début du traitement d'import"
                DataListeIntegrer.Rows.Clear()
                If OledatableSchema.Rows.Count <> 0 Then
                    Dim aRows() As String = Nothing
                    Dim Line_Count As Integer = 0
                    Dim Detail_Count As Integer = 0
                    Dim k As Integer = 1
                    Dim k1 As Integer = 1
                    Dim Cpteur As Integer = 0
                    Dim LigneQuantiteDemandé As Double = 0
                    Dim LigneQuantiteLivre As Double = 0
                    Dim LigneCodeArt As String = ""
                    Dim CountLigne As Integer = 1
                    If GetArrayFile(Chemin, aRows) IsNot Nothing Then
                        aRows = GetArrayFile(Chemin, aRows)
                        For i As Integer = 0 To UBound(aRows)
                            ExisteLecture = True
                            Dim Ligne As String = aRows(i)
                            If GetNombreLigne(Ligne, 1001, 10) <> 0 Or GetNombreLigne(Ligne, 1001, 10) = 0 Then
                                If Not IsNumeric(Strings.Mid(Ligne, 1, 10)) Then
                                    Line_Count = Strings.Mid(Ligne, 1001, 10)
                                    NbEntete += 1
                                    Statut = False
                                    If DataListeIntegrer.RowCount = 0 Then
                                        iline += 1
                                        DataListeIntegrer.RowCount = iline
                                    Else
                                        If DataListeIntegrer.Rows(DataListeIntegrer.RowCount - 1).Cells(0).Value <> "" Then
                                            iline += 1
                                            DataListeIntegrer.RowCount = iline
                                        Else
                                            iline = DataListeIntegrer.RowCount + 1
                                            DataListeIntegrer.RowCount = iline
                                        End If
                                    End If
                                    'ici je dois traite l'entete de la confirmation de commande 
                                    Integrer_Ecriture_Ligne(Ligne)
                                    Dim Pathfichierjournal As String = Pathsfilejournal & "VST_MVTES_" & Format(DateAndTime.Year(Now), "0000") & "" & Format(DateAndTime.Month(Now), "00") & "" & Format(DateAndTime.Day(Now), "00") & "_" & "" & Format(DateAndTime.Hour(Now), "00") & "_" & Format(DateAndTime.Minute(Now), "00") & "_" & Format(DateAndTime.Second(Now), "00") & ".txt"
                                    If Directory.Exists(Pathsfilejournal) = True Then
                                        If File.Exists(Pathfichierjournal) = True Then
                                        Else
                                            ErreurJrn = File.AppendText(Pathfichierjournal)
                                        End If
                                        ArtAdaptater = New OleDbDataAdapter("select * from PARAMETRE  Where  Societe='" & Trim(CodeSociete) & "' and nomtype='COMMERCIAL'", OleConnenection)
                                        ArtDataset = New DataSet
                                        ArtAdaptater.Fill(ArtDataset)
                                        Artdatatable = ArtDataset.Tables(0)
                                        If Artdatatable.Rows.Count <> 0 Then
                                            CptaAdaptater = New OleDbDataAdapter("select * from PARAMETRE  Where  Societe='" & Trim(CodeSociete) & "' and nomtype='COMPTABILITE'", OleConnenection)
                                            CptaDataset = New DataSet
                                            CptaAdaptater.Fill(CptaDataset)
                                            Cptadatatable = CptaDataset.Tables(0)
                                            If Cptadatatable.Rows.Count <> 0 Then
                                                If File.Exists(Trim(Cptadatatable.Rows(0).Item("Chemin1"))) = True Then
                                                    If SocieteConnected(System.IO.Path.GetFileNameWithoutExtension(Trim(Cptadatatable.Rows(0).Item("Chemin1"))), Trim(Cptadatatable.Rows(0).Item("MotPas").ToString), Trim(Cptadatatable.Rows(0).Item("NomUser")), LireChaine(Trim(Cptadatatable.Rows(0).Item("Chemin1")), "CBASE", "ServeurSQL")) = True Then
                                                        FermeBaseCial(BaseCial)
                                                        If OuvreBaseCial(BaseCial, BaseCpta, Trim(Artdatatable.Rows(0).Item("Chemin1")), Trim(Cptadatatable.Rows(0).Item("Chemin1")), Trim(Artdatatable.Rows(0).Item("UserSage")), Trim(Artdatatable.Rows(0).Item("PasseSage").ToString), Trim(Cptadatatable.Rows(0).Item("UserSage")), Trim(Cptadatatable.Rows(0).Item("PasseSage").ToString)) = True Then
                                                            ErreurJrn.WriteLine("Connexion à la Société " & Trim(NomBaseCpta) & " Reussie")
                                                            ErreurJrn.WriteLine("")
                                                            ErreurJrn.WriteLine("Début de traitement du fichier : " & System.IO.Path.GetFileName(Trim(MonFichier)) & " Date de traitement : " & DateTime.Today)
                                                            ErreurJrn.WriteLine("")
                                                        Else
                                                            ErreurJrn.WriteLine("Connexion à la Société - Base Commerciale :" & Trim(NomBaseCpta) & " -Base Comptable :" & Trim(NomBaseCpta) & " Echec de traitement")
                                                            Label2.Text = "Connexion à la Société " & Trim(NomBaseCpta) & " : Echec"
                                                        End If
                                                    Else
                                                        RegardeStatut = False
                                                        ErreurJrn.WriteLine("Echec de Connexion à SQL de base de données :" & System.IO.Path.GetFileNameWithoutExtension(Trim(Cptadatatable.Rows(0).Item("Chemin1"))) & " Serveur : " & LireChaine(Trim(Cptadatatable.Rows(0).Item("Chemin1")), "CBASE", "ServeurSQL"))
                                                        Label2.Text = "Echec de Connexion SQL à la Société " & Trim(NomBaseCpta) & " : Echec de traitement"
                                                    End If
                                                Else
                                                    RegardeStatut = False
                                                    ErreurJrn.WriteLine("Chemin du fichier Comptable : " & Trim(Cptadatatable.Rows(0).Item("Chemin1")) & " inexistant")
                                                    Label2.Text = "Chemin du fichier Comptable : " & Trim(Cptadatatable.Rows(0).Item("Chemin1")) & " inexistant"
                                                End If
                                            Else
                                                RegardeStatut = False
                                                ErreurJrn.WriteLine("Aucune Base Comptable Correspondant à : " & Trim(NomBaseCpta) & " Echec de traitement")
                                            End If
                                        Else
                                            RegardeStatut = False
                                            ErreurJrn.WriteLine("Aucune Base Commerciale Correspondant à : " & Trim(NomBaseCpta) & " Echec de traitement")
                                        End If
                                    Else
                                        Label2.Text = "Le Répertoire Journal :" & Pathsfilejournal & " n'est pas valide "
                                    End If
                                    If RegardeStatut = False Then
                                        Exit Sub
                                    End If
                                    Verification_Parametrage(EnteteIntituleDepot, EntetePieceInterne, EnteteTyPeDocument, Document, infoListe, ComboDate.Text, Nothing, "non", Nothing, Nothing, Nothing)
                                    For LigneCols As Integer = 0 To OledatableSchema.Rows.Count - 2 '- 4
                                        Dim Tableau() As String = OledatableSchema.Rows(LigneCols).Item("Format").ToString.Split("(")(1).Split(".")
                                        If Tableau.Length = 2 Then
                                            Dim Divisuer As Integer = Math.Pow(10, Tableau(1).Split(")")(0))
                                            Dim Valeurs As Object = Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format")))
                                            If Valeurs.ToString.Trim <> "" Then
                                                DataListeIntegrer.Rows(iline - 1).Cells(LigneCols).Value = CDbl(Valeurs / Divisuer).ToString.Trim  ' 
                                            Else
                                                DataListeIntegrer.Rows(iline - 1).Cells(LigneCols).Value = 0
                                            End If
                                        Else
                                            If OledatableSchema.Rows(LigneCols).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Dim Valeurs As Object = Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format")))
                                                If Valeurs.ToString.Trim <> "" Then
                                                    Dim Year As String = Valeurs.ToString.Substring(0, 4)
                                                    Dim Mois As String = Valeurs.ToString.Substring(4, 2)
                                                    Dim Jours As String = Valeurs.ToString.Substring(6, 2)
                                                    Dim Heure As String = Valeurs.ToString.Substring(8, 2)
                                                    Dim Minute As String = Valeurs.ToString.Substring(10, 2)
                                                    Dim Seconde As String = Valeurs.ToString.Substring(12, 2)
                                                    Dim MyDate As String = Year & "-" & Mois & "-" & Jours & " " & Heure & ":" & Minute & ":" & Seconde
                                                    Dim MyNewDate As Date = RenvoieDateValide(MyDate, ComboDate.Text)
                                                    DataListeIntegrer.Rows(iline - 1).Cells(LigneCols).Value = MyDate.Trim 'Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format"))) ' 
                                                Else
                                                    DataListeIntegrer.Rows(iline - 1).Cells(LigneCols).Value = Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format"))).Trim  ' 
                                                End If
                                            Else
                                                DataListeIntegrer.Rows(iline - 1).Cells(LigneCols).Value = Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format"))).Trim  ' 
                                            End If
                                        End If
                                    Next
                                    If ExisteLecture = True Then
                                        If EnteteTyPeDocument = "20" Then
                                            Document = BaseCial.FactoryDocumentStock.CreateType(DocumentType.DocumentTypeStockMouvIn)
                                            Creation_Entete_Document(EnteteTyPeDocument)
                                        ElseIf EnteteTyPeDocument = "21" Then
                                            Document = BaseCial.FactoryDocumentStock.CreateType(DocumentType.DocumentTypeStockMouvOut)
                                            Creation_Entete_Document(EnteteTyPeDocument)
                                        End If
                                        If IsNothing(Document) = False Then
                                            Creation_Ligne_Article(ComboDate.Text, Nothing, Nothing, Nothing, EnteteTyPeDocument)
                                            If exceptionTrouve = False Then
                                                infosExport.Refresh()
                                                infosExport.Text = "Intégration Terminer !"
                                                Datagridaffiche.Rows(i).Cells("C7").Value = My.Resources.accepter
                                                If CheckFille.Checked Then
                                                    File.Move(Chemin, PathsfileSave & "" & Strings.Right(DateAndTime.Year(Now), 2) & "" & Format(DateAndTime.Month(Now), "00") & "" & Format(DateAndTime.Day(Now), "00") & "[" & "" & Format(DateAndTime.Hour(Now), "00") & "-" & Format(DateAndTime.Minute(Now), "00") & "-" & Format(DateAndTime.Second(Now), "00") & "]" & System.IO.Path.GetFileName(Trim(Chemin)))
                                                    infosExport.Refresh()
                                                    infosExport.Text = "Integration Terminée! Suppression des Fichiers exécutée..."
                                                    DataListeIntegrer.Rows.Clear()
                                                End If
                                            End If

                                        End If
                                    End If
                                End If
                            End If
                        Next
                        '''''''''''''''''''''''''''''''''''''''''''
                    End If
                End If
            Catch ex As Exception
                MsgBox("Fonction Aperçu Erreur :" & ex.Message, MsgBoxStyle.Critical)
            End Try
        Else
            Try
                If OledatableSchema.Rows.Count <> 0 Then
                    Dim aRows() As String = Nothing
                    Dim Line_Count As Integer = 0
                    Dim Detail_Count As Integer = 0
                    Dim k As Integer = 1
                    Dim k1 As Integer = 1
                    Dim Cpteur As Integer = 0
                    Dim CountLigne As Integer = 1
                    If GetArrayFile(Chemin, aRows) IsNot Nothing Then
                        aRows = GetArrayFile(Chemin, aRows)
                        For i As Integer = 0 To UBound(aRows)
                            Dim Ligne As String = aRows(i)
                            If GetNombreLigne(Ligne, 1001, 10) <> 0 Or GetNombreLigne(Ligne, 1001, 10) = 0 Then
                                If Not IsNumeric(Strings.Mid(Ligne, 1, 10)) Then
                                    Line_Count = Strings.Mid(Ligne, 1001, 10)
                                    NbEntete += 1
                                    Statut = False
                                    If DataListeIntegrer.RowCount = 0 Then
                                        iline += 1
                                        DataListeIntegrer.RowCount = iline
                                    Else
                                        If DataListeIntegrer.Rows(DataListeIntegrer.RowCount - 1).Cells(0).Value <> "" Then
                                            iline += 1
                                            DataListeIntegrer.RowCount = iline
                                        Else
                                            iline = DataListeIntegrer.RowCount + 1
                                            DataListeIntegrer.RowCount = iline
                                        End If
                                    End If
                                    'ici je dois traite l'entete de la confirmation de commande 
                                    For LigneCols As Integer = 0 To OledatableSchema.Rows.Count - 2 '- 4
                                        Dim Tableau() As String = OledatableSchema.Rows(LigneCols).Item("Format").ToString.Split("(")(1).Split(".")
                                        If Tableau.Length = 2 Then
                                            Dim Divisuer As Integer = Math.Pow(10, Tableau(1).Split(")")(0))
                                            Dim Valeurs As Object = Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format")))
                                            If Valeurs.ToString.Trim <> "" Then
                                                DataListeIntegrer.Rows(iline - 1).Cells(LigneCols).Value = CDbl(Valeurs / Divisuer).ToString.Trim  ' 
                                            Else
                                                DataListeIntegrer.Rows(iline - 1).Cells(LigneCols).Value = 0
                                            End If
                                        Else
                                            If OledatableSchema.Rows(LigneCols).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Dim Valeurs As Object = Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format")))
                                                If Valeurs.ToString.Trim <> "" Then
                                                    Dim Year As String = Valeurs.ToString.Substring(0, 4)
                                                    Dim Mois As String = Valeurs.ToString.Substring(4, 2)
                                                    Dim Jours As String = Valeurs.ToString.Substring(6, 2)
                                                    Dim Heure As String = Valeurs.ToString.Substring(8, 2)
                                                    Dim Minute As String = Valeurs.ToString.Substring(10, 2)
                                                    Dim Seconde As String = Valeurs.ToString.Substring(12, 2)
                                                    Dim MyDate As String = Year & "-" & Mois & "-" & Jours & " " & Heure & ":" & Minute & ":" & Seconde
                                                    Dim MyNewDate As Date = RenvoieDateValide(MyDate, ComboDate.Text)
                                                    DataListeIntegrer.Rows(iline - 1).Cells(LigneCols).Value = MyDate.Trim 'Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format"))) ' 
                                                Else
                                                    DataListeIntegrer.Rows(iline - 1).Cells(LigneCols).Value = Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format"))).Trim  ' 
                                                End If
                                            Else
                                                DataListeIntegrer.Rows(iline - 1).Cells(LigneCols).Value = Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format"))).Trim  ' 
                                            End If
                                        End If
                                    Next
                                    If Line_Count <> 0 Then
                                        'traitement du lot
                                    End If
                                End If
                            End If
                        Next
                        '''''''''''''''''''''''''''''''''''''''''''
                    End If
                End If
            Catch ex As Exception
                MsgBox("Fonction Aperçu Erreur :" & ex.Message, MsgBoxStyle.Critical)
            End Try
        End If
    End Sub
    Public Seconde As Integer = 0
    Public Function RenvoiPachsFile(ByVal codeSociete As String) As String
        Try
            Dim OleAdaptaterschemaCheminIO As OleDbDataAdapter
            Dim OleSchemaDatasetCheminIO As DataSet
            Dim OledatableSchemaCheminIO As DataTable
            OleAdaptaterschemaCheminIO = New OleDbDataAdapter("select distinct CheminFilexport from WIS_SCHEMA WHERE BaseCial='" & codeSociete & "'", OleConnenectionClient)
            OleSchemaDatasetCheminIO = New DataSet
            OleAdaptaterschemaCheminIO.Fill(OleSchemaDatasetCheminIO)
            OledatableSchemaCheminIO = OleSchemaDatasetCheminIO.Tables(0)
            Return OledatableSchemaCheminIO.Rows(0).Item("CheminFilexport")
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Sub Transformateur(ByVal chemin As String)
        Try
            Dim OleAdaptaterschema As OleDbDataAdapter
            Dim OleSchemaDataset As DataSet
            Dim OledatableSchema As DataTable
            Seconde += 1
            Dim iline As Integer = 0
            Dim iLigne As Integer = 0
            Dim iDetaiLigne As Integer = 0
            Dim NbDetaiLigne As Integer = 0
            Dim NbLigne As Integer = 0
            Dim NbEntete As Integer = 0
            Dim Relation As String = ""
            Dim Statut As Boolean = True
            Dim Information As String = ""
            Dim ArtAdaptater As OleDbDataAdapter
            Dim ArtDataset As DataSet
            Dim Artdatatable As DataTable
            Dim CptaAdaptater As OleDbDataAdapter
            Dim CptaDataset As DataSet
            Dim Cptadatatable As DataTable

            RegardeStatut = True
            OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='VST' AND ENTETE=true AND ChampSage<>'' ORDER BY ORDRE", OleConnenection)
            OleSchemaDataset = New DataSet
            OleAdaptaterschema.Fill(OleSchemaDataset)
            OledatableSchema = OleSchemaDataset.Tables(0)

            If OledatableSchema.Rows.Count <> 0 Then
                Dim aRows() As String = Nothing
                Dim Line_Count As Integer = 0
                Dim Detail_Count As Integer = 0
                Dim k As Integer = 1
                Dim k1 As Integer = 1
                Dim Cpteur As Integer = 0
                Dim CountLigne As Integer = 1
                Dim EstPasser As Boolean = False
                If GetArrayFile(chemin, aRows) IsNot Nothing Then
                    aRows = GetArrayFile(chemin, aRows)

                    For i As Integer = 0 To UBound(aRows)
                        Dim Ligne As String = aRows(i)
                        If GetNombreLigne(Ligne, 1001, 10) <> 0 Or GetNombreLigne(Ligne, 1001, 10) = 0 Then
                            If Not IsNumeric(Strings.Mid(Ligne, 1, 10)) Then
                                Line_Count = Strings.Mid(Ligne, 1001, 10)
                                NbEntete += 1
                                CodeSociete = Strings.Mid(Ligne, 175, 12)
                                If EstPasser = False Then
                                    PathsFileVST = RenvoiPachsFile(CodeSociete)
                                    If File.Exists(Trim(chemin)) = True Then
                                        File.Delete(Trim(chemin))
                                        FichierCSO = File.AppendText(PathsFileVST & Seconde & "CW_" & System.IO.Path.GetFileName(chemin))
                                    Else
                                        FichierCSO = File.AppendText(PathsFileVST & Seconde & "CW_" & System.IO.Path.GetFileName(chemin))
                                    End If
                                    EstPasser = True
                                End If

                                ArtAdaptater = New OleDbDataAdapter("select * from PARAMETRE  Where  Societe='" & Trim(CodeSociete) & "' and nomtype='COMMERCIAL'", OleConnenection)
                                ArtDataset = New DataSet
                                ArtAdaptater.Fill(ArtDataset)
                                Artdatatable = ArtDataset.Tables(0)

                                If Artdatatable.Rows.Count <> 0 Then
                                    CptaAdaptater = New OleDbDataAdapter("select * from PARAMETRE  Where  Societe='" & Trim(CodeSociete) & "' and nomtype='COMPTABILITE'", OleConnenection)
                                    CptaDataset = New DataSet
                                    CptaAdaptater.Fill(CptaDataset)
                                    Cptadatatable = CptaDataset.Tables(0)
                                    If Cptadatatable.Rows.Count <> 0 Then
                                        If File.Exists(Trim(Cptadatatable.Rows(0).Item("Chemin1"))) = True Then
                                            If SocieteConnected(System.IO.Path.GetFileNameWithoutExtension(Trim(Cptadatatable.Rows(0).Item("Chemin1"))), Trim(Cptadatatable.Rows(0).Item("MotPas").ToString), Trim(Cptadatatable.Rows(0).Item("NomUser")), LireChaine(Trim(Cptadatatable.Rows(0).Item("Chemin1")), "CBASE", "ServeurSQL")) = False Then
                                                'Echec de connexion a la base xxxxxxxxxxxx
                                                Exit Sub
                                            End If
                                        Else
                                            'le chemin de la societe present dans la table de parametrage n'existe 
                                            Exit Sub
                                        End If
                                    Else
                                        'la societe present dans le fichier n'est pas parametre da la table de correspondance compta
                                        Exit Sub
                                    End If
                                Else
                                    'la societe present dans le fichier n'est pas parametre da la table de correspondance Gescom
                                    Exit Sub
                                End If
                                'ici je dois traite l'entete de la confirmation de commande 
                                For LigneCols As Integer = 0 To OledatableSchema.Rows.Count - 1
                                    Dim Tableau() As String = OledatableSchema.Rows(LigneCols).Item("Format").ToString.Split("(")(1).Split(".")
                                    If LigneCols = OledatableSchema.Rows.Count - 1 Then
                                        If Tableau.Length = 2 Then
                                            Dim Divisuer As Integer = Math.Pow(10, Tableau(1).Split(")")(0))
                                            Dim Valeurs As Object = Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format")))
                                            If Valeurs.ToString.Trim <> "" Then
                                                Information &= CDbl(Valeurs / Divisuer).ToString.Trim & "" ' 
                                            Else
                                                Information &= ""
                                            End If
                                        Else
                                            If OledatableSchema.Rows(LigneCols).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Dim Valeurs As Object = Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format")))
                                                If Valeurs.ToString.Trim <> "" Then
                                                    Dim Year As String = Valeurs.ToString.Substring(0, 4)
                                                    Dim Mois As String = Valeurs.ToString.Substring(4, 2)
                                                    Dim Jours As String = Valeurs.ToString.Substring(6, 2)
                                                    Dim Heure As String = Valeurs.ToString.Substring(8, 2)
                                                    Dim Minute As String = Valeurs.ToString.Substring(10, 2)
                                                    Dim Seconde As String = Valeurs.ToString.Substring(12, 2)
                                                    Dim MyDate As String = Year & "-" & Mois & "-" & Jours '& " " & Heure & ":" & Minute & ":" & Seconde
                                                    Dim MyNewDate As Date = RenvoieDateValide(MyDate, ComboDate.Text)
                                                    Information &= MyDate.Trim & "" ' 
                                                Else
                                                    Information &= Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format"))).Trim & ""  ' 
                                                End If
                                            ElseIf OledatableSchema.Rows(LigneCols).Item("Format").ToString = "N(10)" Then
                                                Dim valeur As String = Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format"))).Trim
                                                If valeur.ToString <> "" Then
                                                    Information &= Convert.ToDecimal(valeur) & ""
                                                Else
                                                    Information &= ""
                                                End If
                                            ElseIf OledatableSchema.Rows(LigneCols).Item("Cols").ToString = "SIGN" Then
                                                If Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format"))).Trim = "+" Then
                                                    Information &= "20"
                                                Else
                                                    Information &= "21"
                                                End If
                                            Else
                                                Information &= Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format"))).Trim
                                            End If
                                        End If
                                        FichierCSO.WriteLine(Information)
                                        Information = ""
                                    Else
                                        If Tableau.Length = 2 Then
                                            Dim Divisuer As Integer = Math.Pow(10, Tableau(1).Split(")")(0))
                                            Dim Valeurs As Object = Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format")))
                                            If Valeurs.ToString.Trim <> "" Then
                                                Information &= CDbl(Valeurs / Divisuer).ToString.Trim & ";" ' 
                                            Else
                                                Information &= ";"
                                            End If
                                        Else
                                            If OledatableSchema.Rows(LigneCols).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Dim Valeurs As Object = Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format")))
                                                If Valeurs.ToString.Trim <> "" Then
                                                    Dim Year As String = Valeurs.ToString.Substring(0, 4)
                                                    Dim Mois As String = Valeurs.ToString.Substring(4, 2)
                                                    Dim Jours As String = Valeurs.ToString.Substring(6, 2)
                                                    Dim Heure As String = Valeurs.ToString.Substring(8, 2)
                                                    Dim Minute As String = Valeurs.ToString.Substring(10, 2)
                                                    Dim Seconde As String = Valeurs.ToString.Substring(12, 2)
                                                    Dim MyDate As String = Year & "-" & Mois & "-" & Jours '& " " & Heure & ":" & Minute & ":" & Seconde
                                                    Dim MyNewDate As Date = RenvoieDateValide(MyDate, ComboDate.Text)
                                                    Information &= MyDate.Trim & ";" ' 
                                                Else
                                                    Information &= Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format"))).Trim & ";"  ' 
                                                End If
                                            ElseIf OledatableSchema.Rows(LigneCols).Item("Format").ToString = "N(10)" Then
                                                Dim valeur As String = Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format"))).Trim
                                                If valeur.ToString <> "" Then
                                                    Information &= Convert.ToDecimal(valeur) & ";"
                                                Else
                                                    Information &= ";"
                                                End If
                                            ElseIf OledatableSchema.Rows(LigneCols).Item("Cols").ToString = "SIGN" Then
                                                If Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format"))).Trim = "+" Then
                                                    Information &= "20;"
                                                Else
                                                    Information &= "21;"
                                                End If
                                            Else
                                                Information &= Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format"))).Trim & ";" ' 
                                            End If
                                        End If
                                    End If
                                Next
                                Statut = False
                                If Line_Count <> 0 Then
                                    'traitement du lot
                                End If
                            End If
                        End If
                    Next
                    If ChEncapsuler.Checked Then
                        FichierCSO.Close()
                        EstPasser = False
                    End If
                    infosExport.Text = "Export Terminer !"
                    'If CheckFille.Checked Then
                    '    File.Move(chemin, PathsfileSave & "" & Strings.Right(DateAndTime.Year(Now), 2) & "" & Format(DateAndTime.Month(Now), "00") & "" & Format(DateAndTime.Day(Now), "00") & "[" & "" & Format(DateAndTime.Hour(Now), "00") & "-" & Format(DateAndTime.Minute(Now), "00") & "-" & Format(DateAndTime.Second(Now), "00") & "]" & System.IO.Path.GetFileName(Trim(chemin)))
                    '    infosExport.Refresh()
                    '    infosExport.Text = "Integration Terminée! Suppression des Fichiers exécutée..."
                    '    DataListeIntegrer.Rows.Clear()
                    '    'OuvreLaListedeFichier(PathsfileExport)
                    'End If
                    '''''''''''''''''''''''''''''''''''''''''''
                End If
            End If
        Catch ex As Exception
            MsgBox("Transformation Variation de Stock Erreur :" & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub
    Public Function GetLongueurChaine(ByVal Format As Object) As Integer
        Dim Position() As Object = Format.ToString.Split("(")(1).ToString.Split(")")(0).Split(".")
        If Position.Length = 2 Then
            Return Position(0)
        Else
            Return Position(0)
        End If
    End Function
    Public Sub EstInfosLibre(ByVal Entete As Boolean, ByVal Ligne As Boolean, ByVal InfosLibre As Boolean)
        Dim OleAdaptaterschema As OleDbDataAdapter
        Dim OleSchemaDataset As DataSet
        Dim OledatableSchema As DataTable
        Dim ListeInfosLibre As String = ""
        NbInfosLibreVue = 0
        'lblSne.Text = "Scenario verification infos libre"
        Try
            OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='VST' AND Entete=" & Entete & " AND InfosLibre=" & InfosLibre & " AND Ligne=" & Ligne, OleConnenection)
            OleSchemaDataset = New DataSet
            OleAdaptaterschema.Fill(OleSchemaDataset)
            OledatableSchema = OleSchemaDataset.Tables(0)
            If OledatableSchema.Rows.Count <> 0 Then
                Dim Pathfichierjournal As String = Pathsfilejournal & "VST_Mouvement_IO_" & Format(DateAndTime.Year(Now), "0000") & "" & Format(DateAndTime.Month(Now), "00") & "" & Format(DateAndTime.Day(Now), "00") & "_" & "" & Format(DateAndTime.Hour(Now), "00") & "_" & Format(DateAndTime.Minute(Now), "00") & "_" & Format(DateAndTime.Second(Now), "00") & ".txt"
                NbInfosLibre = OledatableSchema.Rows.Count
                If Directory.Exists(Pathsfilejournal) = True Then
                    If File.Exists(Pathfichierjournal) = True Then
                    Else
                        ErreurJrn = File.AppendText(Pathfichierjournal)
                    End If
                    ErreurJrn.WriteLine("<-----------------------------------------------------------------DEBUT DU TRAITEMENT DE LA VERIFICATION DES INFOS LIBRES----------------------------------------------------------------->")
                    For i As Integer = 0 To OledatableSchema.Rows.Count - 1
                        If OledatableSchema.Rows(i).Item("InfosLibre") = "true" Then
                            If Entete = True Then
                                If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False Then
                                    If ExisteInfosLibre(OledatableSchema.Rows(i).Item("ChampSage"), Entete) = False Then
                                        ErreurJrn.WriteLine("la zone infos libre {" & OledatableSchema.Rows(i).Item("ChampSage") & "} en Entete du Bon de Commande Client n'existe pas dans Sage")
                                    Else
                                        ErreurJrn.WriteLine("Traitement  de la zone infos libre {" & OledatableSchema.Rows(i).Item("ChampSage") & "} sur le Bon de Commande Client Existe [OK] dans Sage")
                                    End If
                                Else
                                    ErreurJrn.WriteLine("<--Le Champ indiquant l'infos libre est couché mais ne possede pas de mapping Sage-->")
                                End If
                            ElseIf Ligne = True Then
                                If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False Then
                                    If ExisteInfosLibre(OledatableSchema.Rows(i).Item("ChampSage"), Entete) = False Then
                                        ErreurJrn.WriteLine("la zone infos libre {" & OledatableSchema.Rows(i).Item("ChampSage") & "} sur le Bon de Commande Client n'existe pas dans Sage")
                                    End If
                                Else
                                    ErreurJrn.WriteLine("<--Le Champ indiquant l'infos libre est couché mais ne possede pas de mapping Sage-->")
                                End If
                            Else
                                ErreurJrn.WriteLine("<--Aucune information libre n'est parametrée-->")
                            End If
                        End If
                    Next
                    ErreurJrn.WriteLine("<-----------------------------------------------------------------Fin----------------------------------------------------------------->")
                    ErreurJrn.Close()
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub
    Public Function ExisteInfosLibre(ByVal InfosLibre As String, ByVal Entete As Boolean) As Boolean
        Try
            Dim OledatableSchemaSage As DataTable
            If ExcelConnect(NomServersql, NomBaseCpta, Nom_Utilsql, Mot_Passql) Then
                If Entete Then
                    OleAdaptaterschemaSage = New OleDbDataAdapter("select * from cbSysLibre WHERE CB_File='F_DOCENTETE' And CB_Name='" & Join(Split(Trim(InfosLibre), "'"), "''") & "'", OleExcelConnect)
                Else
                    OleAdaptaterschemaSage = New OleDbDataAdapter("select * from cbSysLibre WHERE CB_File='F_DOCLIGNE' And CB_Name='" & Join(Split(Trim(InfosLibre), "'"), "''") & "'", OleExcelConnect)
                End If
                OleSchemaDatasetSage = New DataSet
                OleAdaptaterschemaSage.Fill(OleSchemaDatasetSage)
                OledatableSchemaSage = OleSchemaDatasetSage.Tables(0)
                If OledatableSchemaSage.Rows.Count <> 0 Then
                    NbInfosLibreVue += 1
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
        End Try
    End Function
    Public Sub OledbInitialiseur()
        Try
            OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='VST' AND ENTETE=true ORDER BY ORDRE", OleConnenection)
            OleSchemaDataset = New DataSet
            OleAdaptaterschema.Fill(OleSchemaDataset)
            OledatableSchema = OleSchemaDataset.Tables(0)
        Catch ex As Exception
        End Try
    End Sub
    Public Function GetNombreLigne(ByVal Item As String, ByVal position As Integer, ByVal Longueur As Integer) As Integer
        Try
            Return CInt(Strings.Mid(Item, position, Longueur))
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Public Sub PicLigne_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PicLigne.Click
        Dim IligneCE As Integer = 0
        Dim IligneCV As Integer = 0
        With Frm_OptionGrid
            .choix = "Entete"
            .ShowsForm = "3"
            .DGVCV.Rows.Clear()
            .DGVCV.Rows.Clear()
            For i As Integer = 0 To DataListeIntegrer.ColumnCount - 1
                If DataListeIntegrer.Columns.Item(i).Visible = True Then
                    IligneCV += 1
                    .DGVCV.RowCount = IligneCV
                    .DGVCV.Rows(IligneCV - 1).Cells(0).Value = DataListeIntegrer.Columns.Item(i).HeaderText.Trim
                    .DGVCV.Rows(IligneCV - 1).Cells(1).Value = i
                Else
                    IligneCE += 1
                    .DGVCE.RowCount = IligneCE
                    .DGVCE.Rows(IligneCE - 1).Cells(0).Value = DataListeIntegrer.Columns.Item(i).HeaderText.Trim
                    .DGVCE.Rows(IligneCE - 1).Cells(1).Value = i
                End If
            Next i
            If Vsate = True Then
                .ShowDialog()
            End If
        End With
    End Sub
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
    Public Function SocieteConnected(ByRef BaseConsolide As String, ByRef Mot_Psql As String, ByRef Nom_Utsql As String, ByRef Serveur As String) As Boolean
        Try
            OleSocieteConnect = New OleDbConnection("provider=SQLOLEDB;UID=" & Trim(Nom_Utsql) & ";Pwd=" & Trim(Mot_Psql) & ";Initial Catalog=" & Trim(BaseConsolide) & ";Data Source=" & Trim(Serveur) & "")
            OleSocieteConnect.Open()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public MonFichier As String = ""
    Public IfrowErreur As Integer
    Public ExisteLectures As Boolean
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_integrer.Click
        ExisteLecture = True
        OledbInitialiseur()
        'vidage()
        Label2.Text = ""
        IfrowErreur = 0
        Dim i As Integer
        Try
            CountChecked = IsChecked()
            If IsChecked() Then
                EstInfosLibre(False, False, True)
                If NbInfosLibre = NbInfosLibreVue Then 'NbInfosLibreVue
                    For i = 0 To Datagridaffiche.RowCount - 1
                        If Datagridaffiche.Rows(i).Cells("C6").Value = True Then
                            IfrowErreur = i
                            MonFichier = Datagridaffiche.Rows(i).Cells("C8").Value
                            AperçuElement(Datagridaffiche.Rows(i).Cells("C8").Value, "Execution")
                            If RegardeStatut = True And ExisteLecture = True And exceptionTrouve = False Then
                                Datagridaffiche.Rows(i).Cells("C7").Value = My.Resources.accepter
                                DataListeIntegrer.Rows.Clear()
                                ErreurJrn.Close()
                            Else
                                Datagridaffiche.Rows(i).Cells("C7").Value = My.Resources._error
                                ErreurJrn.Close()
                            End If
                        End If
                    Next i
                Else

                End If
                If CheckFille.Checked Then
                    OuvreLaListedeFichier(PathsfileExport)
                End If
            Else
                RegardeStatut = True
                MessageBox.Show("Un choix de traitement doit être fait " & vbCrLf & vbCrLf & " Merci de faire votre choix", "Infos Choix Traitement", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
            If RegardeStatut = False Then
                ErreurJrn.Close()
            End If
        Catch ex As Exception
            Datagridaffiche.Rows(IfrowErreur).Cells("C7").Value = My.Resources.criticalind_status
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Action Execution")
        End Try
    End Sub 'Execution
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnview.Click
        Dim i As Integer = 0
        CountChecked = IsChecked()
        DataListeIntegrer.Rows.Clear()
        For i = 0 To Datagridaffiche.RowCount - 1
            If Datagridaffiche.Rows(i).Cells("C6").Value = True Then
                AperçuElement(Datagridaffiche.Rows(i).Cells("C8").Value, "")
                Datagridaffiche.Rows(i).Cells("C7").Value = My.Resources.exlamation
            End If
        Next i
    End Sub 'aperçu
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()
    End Sub
    Public Sub Frm_FluxEntrantCritére_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.WindowState = FormWindowState.Maximized
            ComboDate.SelectedIndex = 4
            BtnListe_Click_1(sender, e)
        Catch ex As Exception
        End Try
    End Sub
    Public Function ExisteMappingSage(ByVal NameColonne As String, Optional ByVal TableLie As String = "") As String
        Dim OleAdaptaterschema As OleDbDataAdapter
        Dim OleSchemaDataset As DataSet
        Dim OledatableSchema As DataTable
        Dim ColonneMappe As String = ""
        Try
            OleAdaptaterschema = New OleDbDataAdapter("select Champ,Libelle from COLIMPMOUV WHERE  Champ='" & NameColonne & "'", OleConnenection) 'Fichier='F_DOCENTETE' AND
            OleSchemaDataset = New DataSet
            OleAdaptaterschema.Fill(OleSchemaDataset)
            OledatableSchema = OleSchemaDataset.Tables(0)
            If OledatableSchema.Rows.Count <> 0 Then
                ColonneMappe = OledatableSchema.Rows(0).Item("Libelle")
            End If
        Catch ex As Exception
            ColonneMappe = ""
        End Try
        ExisteMappingSage = ColonneMappe
    End Function
    Public DepotPrincipal As String = ""
    Private Sub Integrer_Ecriture_Ligne(ByVal Document_Ligne As String, Optional ByVal TableLie As String = "")
        Dim fournisseurAdap As OleDbDataAdapter
        Dim fournisseurDs As DataSet
        Dim fournisseurTab As DataTable
        Dim PieceAutoma As String = ""
        infosExport.Refresh()
        infosExport.Text = "Integration des En Cours..."
        fournisseurAdap = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='VST' AND Entete=" & True & " AND InfosLibre=" & False & " AND Ligne=" & False & " AND ChampSage<>'' ORDER BY ORDRE", OleConnenection)
        fournisseurDs = New DataSet
        fournisseurAdap.Fill(fournisseurDs)
        fournisseurTab = fournisseurDs.Tables(0)
        Try
            If fournisseurTab.Rows.Count <> 0 Then
                For numColonne As Integer = 0 To fournisseurTab.Rows.Count - 1
                    'Entête Document
                    If fournisseurTab.Rows(numColonne).Item("ChampSage").ToString.Trim <> "" Then
                        '-----------------------------------------------------------------------------------------------
                        DefaultValeur = fournisseurTab.Rows(numColonne).Item("DefaultValue").ToString.Trim
                        PositionG = fournisseurTab.Rows(numColonne).Item("PositionG").ToString.Trim
                        Longueur = GetLongueurChaine(fournisseurTab.Rows(numColonne).Item("Format").ToString.Trim)
                        '------------------------------------------------------------------------------------------------
                        'Entête Document
                        CodeSociete = Strings.Mid(Document_Ligne, 175, 12).Trim
                        If Strings.Mid(Document_Ligne, 237, 1).Trim = "+" Then
                            EnteteTyPeDocument = 20
                            lblType.Text = "Mouvement d'entrée"
                        ElseIf Strings.Mid(Document_Ligne, 237, 1).Trim = "-" Then
                            EnteteTyPeDocument = 21
                            lblType.Text = "Mouvement de sortie"
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "EntetePieceInterne" Then
                            EntetePieceInterne = Strings.Mid(Document_Ligne, PositionG, Longueur)
                            If EntetePieceInterne.ToString.Trim <> "" Then
                                If Strings.Len(Trim(EntetePieceInterne.ToString.Trim)) <= 8 Then
                                    EntetePieceInterne = Formatage_Chaine(Trim(EntetePieceInterne.ToString.Trim))
                                Else
                                    EntetePieceInterne = Formatage_Chaine(Strings.Left(Trim(EntetePieceInterne.ToString.Trim), 8))
                                End If
                            ElseIf DefaultValeur.Trim <> "" Then
                                EntetePieceInterne = DefaultValeur.Trim
                                If Strings.Len(Trim(DefaultValeur.Trim)) <= 8 Then
                                    EntetePieceInterne = Formatage_Chaine(Trim(DefaultValeur.Trim))
                                Else
                                    EntetePieceInterne = Formatage_Chaine(Strings.Left(Trim(DefaultValeur.ToString.Trim), 8))
                                End If
                            End If
                            Continue For
                        End If
                        'If Datagridaffiche.Columns.Contains(IdentifiantArticle) = True ThenStrings.Left(
                        '    PieceArticle = Trim(Datagridaffiche.Rows(numLigne).Cells(IdentifiantArticle).Value)
                        'End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "EnteteTyPeDocument" Then
                            EnteteTyPeDocument = Trim(Strings.Mid(Document_Ligne, PositionG, Longueur))
                            If Strings.Mid(Document_Ligne, 237, 1).Trim = "+" Then
                                EnteteTyPeDocument = 20
                            ElseIf Strings.Mid(Document_Ligne, 237, 1).Trim = "-" Then
                                EnteteTyPeDocument = 21
                            End If
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "EnteteCodeAffaire" Then
                            EnteteCodeAffaire = Trim(Strings.Mid(Document_Ligne, PositionG, Longueur))
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "EntetePlanAnalytique" Then
                            EntetePlanAnalytique = Trim(Strings.Mid(Document_Ligne, PositionG, Longueur))
                        End If

                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "EnteteDateDocument" Then
                            EnteteDateDocument = Trim(Strings.Mid(Document_Ligne, PositionG, Longueur))
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "EnteteIntituleDepot" Then
                            EnteteIntituleDepot = Trim(Strings.Mid(Document_Ligne, PositionG, Longueur))
                            If EnteteIntituleDepot = "" Then
                                EnteteIntituleDepot = Trim(DefaultValeur)
                            End If
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "EnteteReference" Then
                            If Strings.Len(Trim(Strings.Mid(Document_Ligne, PositionG, Longueur))) <= 17 Then
                                EnteteReference = Trim(Strings.Mid(Document_Ligne, PositionG, Longueur))
                            Else
                                EnteteReference = Strings.Left(Trim(Strings.Mid(Document_Ligne, PositionG, Longueur)), 17)
                            End If
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "EnteteSoucheDocument" Then
                            EnteteSoucheDocument = Trim(Strings.Mid(Document_Ligne, PositionG, Longueur))
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LigneDatedeFabrication" Then
                            LigneDatedeFabrication = Trim(Strings.Mid(Document_Ligne, PositionG, Longueur))
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LigneDatedeLivraison" Then
                            LigneDatedeLivraison = Trim(Strings.Mid(Document_Ligne, PositionG, Longueur))
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LigneDatedePeremption" Then
                            LigneDatedePeremption = Trim(Strings.Mid(Document_Ligne, PositionG, Longueur))
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LigneDesignationArticle" Then
                            If Strings.Len(Trim(Strings.Mid(Document_Ligne, PositionG, Longueur))) <= 69 Then
                                LigneDesignationArticle = Trim(Strings.Mid(Document_Ligne, PositionG, Longueur))
                            Else
                                LigneDesignationArticle = Strings.Left(Trim(Strings.Mid(Document_Ligne, PositionG, Longueur)), 69)
                            End If
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "IDDepotEntete" Then
                            IDDepotEntete = Trim(Strings.Mid(Document_Ligne, PositionG, Longueur))
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LigneNSerieLot" Then
                            If Strings.Len(Trim(Strings.Mid(Document_Ligne, PositionG, Longueur))) <= 30 Then
                                LigneNSerieLot = Trim(Strings.Mid(Document_Ligne, PositionG, Longueur))
                            Else
                                LigneNSerieLot = Strings.Left(Trim(Strings.Mid(Document_Ligne, PositionG, Longueur)), 30)
                            End If
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LignePoidsBrut" Then
                            LignePoidsBrut = Trim(Strings.Mid(Document_Ligne, PositionG, Longueur))
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LignePoidsNet" Then
                            LignePoidsNet = Trim(Strings.Mid(Document_Ligne, PositionG, Longueur))
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LignePrixUnitaire" Then
                            LignePrixUnitaire = Trim(Strings.Mid(Document_Ligne, PositionG, Longueur))
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LigneQuantite" Then
                            LigneQuantite = Trim(Strings.Mid(Document_Ligne, PositionG, Longueur)) / Math.Pow(10, 5)
                        End If

                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LigneReference" Then
                            If Strings.Len(Trim(Strings.Mid(Document_Ligne, PositionG, Longueur))) <= 17 Then
                                LigneReference = Trim(Strings.Mid(Document_Ligne, PositionG, Longueur))
                            Else
                                LigneReference = Strings.Left(Trim(Strings.Mid(Document_Ligne, PositionG, Longueur)), 17)
                            End If
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LigneCodeArticle" Then
                            LigneCodeArticle = Formatage_Article(Trim(Strings.Mid(Document_Ligne, PositionG, Longueur)))
                        End If

                        'RECHERCHE DE L'INTITULE DE L'INFO LIBRE
                        Dim InfoTableau() As String = {"oui2", "F_DOCENTETE"} 'Split(Trim(Strings.Left(Strings.Right(Datagridaffiche.Columns(numColonne).HeaderText, Strings.Len(Datagridaffiche.Columns(numColonne).HeaderText) - InStr(Datagridaffiche.Columns(numColonne).HeaderText, "[")), Strings.Len(Strings.Right(Datagridaffiche.Columns(numColonne).HeaderText, Strings.Len(Datagridaffiche.Columns(numColonne).HeaderText) - InStr(Datagridaffiche.Columns(numColonne).HeaderText, "["))) - 1)), "-")
                        If Trim(InfoTableau(0)) = "oui" Then
                            If Trim(InfoTableau(1)) = "F_DOCENTETE" Then
                                infoListe.Add("----")
                            End If
                            If Trim(InfoTableau(1)) = "F_DOCLIGNE" Then
                                infoLigne.Add("----")
                            End If
                        End If
                    End If
                Next
                Dim r = 1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Fonction d'integration")
        End Try
    End Sub

    'Private Sub vidage()
    '    PieceArticle = Nothing
    '    'EnteteStatutdocument = Nothing
    '    LigneCodeArticle = Nothing

    '    EnteteCodeAffaire = Nothing

    '    EnteteDateDocument = Nothing

    '    IDDepotEntete = Nothing

    '    EnteteCatégorietarifaire = Nothing
    '    EnteteConditiondeLivraison = Nothing
    '    EnteteIntituleDepot = Nothing
    '    EnteteIntituleDepotClient = Nothing
    '    EnteteIntituleDevise = Nothing
    '    EnteteIntituleExpédition = Nothing
    '    EntetePieceInterne = Nothing
    '    EnteteNatureTransaction = Nothing
    '    EnteteNomReprésentant = Nothing
    '    EnteteNombredeFacture = Nothing
    '    EntetePlanAnalytique = Nothing
    '    EntetePrenomReprésentant = Nothing
    '    EnteteReference = Nothing
    '    EnteteRegimeDocument = Nothing
    '    EnteteSoucheDocument = Nothing
    '    EnteteTauxescompte = Nothing
    '    EnteteTyPeDocument = Nothing
    '    LigneCodeAffaire = Nothing
    '    LigneDatedeFabrication = Nothing
    '    LigneDatedeLivraison = Nothing
    '    LigneDatedePeremption = Nothing
    '    LigneDesignationArticle = Nothing
    '    LigneLibelleComplementaire = Nothing
    '    LigneEnumereConditionnement = Nothing
    '    LigneFraisApproche = Nothing
    '    LigneIntituleDepot = Nothing
    '    LigneNSerieLot = Nothing
    '    LigneNomRepresentant = Nothing
    '    LignePlanAnalytique = Nothing
    '    LignePoidsBrut = Nothing
    '    LignePoidsNet = Nothing
    '    LignePrenomRepresentant = Nothing
    '    LignePrixdeRevientUnitaire = Nothing
    '    LignePrixUnitaire = Nothing
    '    LigneQuantite = Nothing
    '    LigneQuantiteConditionne = Nothing
    '    LigneReference = Nothing
    '    LigneArticleCompose = Nothing
    '    LigneReferenceArticleTiers = Nothing
    '    LigneTauxRemise1 = Nothing
    '    LigneTauxRemise2 = Nothing
    '    LigneTauxRemise3 = Nothing
    '    LigneTypeRemise1 = Nothing
    '    LigneTypeRemise2 = Nothing
    '    LigneTypeRemise3 = Nothing
    '    EnteteContact = Nothing
    '    EnteteLangue = Nothing
    '    EnteteCours = Nothing
    '    LignePrixUnitaireDevise = Nothing
    '    'LigneValorisé = Nothing
    'End Sub

    Private Sub PictureBox4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox4.Click
        For i As Integer = 0 To Datagridaffiche.RowCount - 1
            Datagridaffiche.Rows(i).Cells("C6").Value = True
        Next i
    End Sub

    Private Sub PictureBox3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox3.Click
        For i As Integer = 0 To Datagridaffiche.RowCount - 1
            Datagridaffiche.Rows(i).Cells("C6").Value = False
        Next i
    End Sub

    Public Sub BtnListe_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnListe.Click
        Try
            LirefichierConfig()
            OuvreLaListedeFichier(PathsfileExport)
            Connected()
        Catch ex As Exception
        End Try
    End Sub

    Public Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Button1_Click(sender, e)
            If Directory.Exists(PathsFileVST) = True Then
                For i As Integer = 0 To Datagridaffiche.RowCount - 1
                    If Datagridaffiche.Rows(i).Cells("C6").Value = True Then
                        Transformateur(Datagridaffiche.Rows(i).Cells("C8").Value)
                        If RegardeStatut = True Then
                            Datagridaffiche.Rows(i).Cells("C7").Value = My.Resources.accepter
                            If ChEncapsuler.Checked = False Then
                                FichierCSO.Close()
                            End If
                        Else
                            Datagridaffiche.Rows(i).Cells("C7").Value = My.Resources._error
                        End If
                    End If
                Next i
                Seconde = 0
            Else
                MsgBox("Repertoire inexistant", MsgBoxStyle.Information, "Test existance du dossier source extraction ")
            End If
        Catch ex As Exception
        End Try
    End Sub
End Class