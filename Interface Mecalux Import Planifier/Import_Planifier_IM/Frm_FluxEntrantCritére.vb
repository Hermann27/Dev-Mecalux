Imports System.IO
Imports Objets100Lib
Imports System.Data.OleDb
Public Class Frm_FluxEntrantCritére
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
    Public RegardeStatut As Boolean = True
    Public MonFichier As String = ""
    Public IfrowErreur As Integer
    Public Vsate As Boolean = True
    Public StatutCreationEnteteDoc As Boolean = False
    Public StatutCreationLigneDoc As Boolean = False
    Public Achat As String = ""
#End Region
#Region "Variable import BL"
    Public ProgresMax, IndexPrec, numLigne, numColonne, NbreLigne, NbreTotal, iRow As Integer
    Public Result, sColumnsSepar, DecFormat, LigneTypePrixUnitaire, EnteteStatutdocument, LigneValorisé As Object
    Public Filebool As Boolean
    Public ListeReliquat As List(Of String)
    Public PieceCommande, NLignePieceCommande, PieceArticle, IDDepotEntete, IDDepotLigne, LignePrixUnitaireDevise As Object
    Public EcheanceConditionPaiement, EcheanceModeleReglement, EcheanceModeReglement, EcheanceDatePied As Object
    Public EnteteBLFacture, EnteteCodeAffaire, EnteteCodeTiers, EnteteCodeTiersPayeur As Object
    Public EnteteColisagedeLivraison, EnteteCompteGeneral, EnteteDateDocument As Object
    Public EnteteDateLivraison, EnteteEcartValorisation, Entete1, Entete2, Entete3, PieceReliquat As Object
    Public Entete4, EnteteCatégorieComptable, EnteteCatégorietarifaire, EnteteConditiondeLivraison As Object
    Public EnteteIntituleDepot, EnteteIntituleDepotClient, EnteteCodeFournisseur, EnteteCodeTransfertDepot, EnteteTyPeDocumentTry, EnteteIntituleDevise, EnteteIntituleExpédition, EntetePieceInterne, EntetePiecePrecedent As Object
    Public EnteteNatureTransaction, EnteteNomReprésentant, EnteteNombredeFacture As Object
    Public EntetePlanAnalytique, EntetePrenomReprésentant, CodeSociete, EnteteReference, EnteteRegimeDocument As Object
    Public EnteteSoucheDocument, EnteteTauxescompte, EnteteTyPeDocument, EnteteCours As Object
    Public LigneCodeAffaire, LigneDatedeFabrication, EnteteDoExpedition, LigneDatedeLivraison, LigneDatedePeremption As Object
    Public LigneDesignationArticle, LigneEnumereConditionnement, LigneFraisApproche, LigneLibelleComplementaire As Object
    Public LigneIntituleDepot, LigneNSerieLot, LigneNomRepresentant, LignePlanAnalytique As Object
    Public LignePoidsBrut, LignePoidsNet, LignePrenomRepresentant, LigneFraisExpedition, EnteteFraisExpedition, LignePrixdeRevientUnitaire As Object
    Public LignePrixUnitaire, LigneQuantite, LigneQuantiteConditionne, LigneReference, ProvenanceFacture As Object
    Public LigneReferenceArticleTiers, LigneTauxRemise1, LigneTauxRemise2, LigneTauxRemise3, LigneArticleCompose As Object
    Public LigneTypeRemise1, LigneTypeRemise2, LigneTypeRemise3, EnteteContact, EnteteLangue, LigneCodeArticle, EnteteUniteColis As Object
    Public OleSocieteConnect As OleDbConnection
    Public OM_ArticleComposant As IBOArticle3 = Nothing
    Public OM_QteCompose As Double = Nothing
    Public OM_MontantCompose As Double = Nothing
    'Variable d'exception du deplacement de fichier
    Public DeviseTiers As Object = Nothing
    Public exceptionTrouve As Boolean = False
    Public Er_cre_entete_doc As Boolean = False
    Public ExisteLecture As Boolean = True
    Public infoListe As List(Of Integer)
    Public infoLigne As List(Of Integer)
    Public ListePiece As List(Of String)
    Public Documents As IBODocumentVente3 = Nothing
    Public LigneDocument As IBODocumentVenteLigne3 = Nothing
    Public DocumentInfolibre As IBODocumentVente3 = Nothing
    Public DocumentReliquat As IBODocumentVente3 = Nothing
    Public LigneReliquat As IBODocumentVenteLigne3 = Nothing
    Public PlanAna As IBPAnalytique3
    Public NumeroLot As String
    Public TraitementID As Integer
    Public ListeStock As List(Of String)
    Public DateLancer As Date
    Public FormatQte As Integer
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
                For i = 0 To UBound(aLines)
                    NomFichier = Trim(aLines(i))
                    Do While InStr(Trim(NomFichier), "\") <> 0
                        NomFichier = Strings.Right(NomFichier, Strings.Len(Trim(NomFichier)) - InStr(Trim(NomFichier), "\"))
                    Loop
                    If Critere = NomFichier.Substring(0, 3) Then
                        Select Case NomFichier.Substring(0, 3)
                            Case "CSO"
                                Datagridaffiche.RowCount = jRow + 1
                                Datagridaffiche.Rows(jRow).Cells("C1").Value = "Confirmation de Commande"
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
                                'Exit For
                        End Select
                    End If
                Next i
                aLines = Nothing
            Else
                'MsgBox("Ce Repertoire n'est pas valide! " & Directpath, MsgBoxStyle.Information, "Repertoire des Fichiers à Traiter")
            End If
        Catch ex As Exception
        End Try
    End Sub
    Public Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        Dim IligneCE As Integer = 0
        Dim IligneCV As Integer = 0
        With Frm_OptionGrid
            .choix = "Ligne"
            .ShowsForm = "1"
            .DGVCV.Rows.Clear()
            .DGVCV.Rows.Clear()
            For i As Integer = 0 To DataListeIntegrerLigne.ColumnCount - 1
                If DataListeIntegrerLigne.Columns.Item(i).Visible = True Then
                    IligneCV += 1
                    .DGVCV.RowCount = IligneCV
                    .DGVCV.Rows(IligneCV - 1).Cells(0).Value = DataListeIntegrerLigne.Columns.Item(i).HeaderText.Trim
                    .DGVCV.Rows(IligneCV - 1).Cells(1).Value = i
                Else
                    IligneCE += 1
                    .DGVCE.RowCount = IligneCE
                    .DGVCE.Rows(IligneCE - 1).Cells(0).Value = DataListeIntegrerLigne.Columns.Item(i).HeaderText.Trim
                    .DGVCE.Rows(IligneCE - 1).Cells(1).Value = i
                End If
            Next i
            If Vsate = True Then
                .ShowDialog()
            End If
        End With
    End Sub
    Public Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click
        Dim IligneCE As Integer = 0
        Dim IligneCV As Integer = 0
        With Frm_OptionGrid
            .choix = "SousLigne"
            .ShowsForm = "1"
            .DGVCV.Rows.Clear()
            .DGVCV.Rows.Clear()
            For i As Integer = 0 To DataListeIntegrerDétailLigne.ColumnCount - 1
                If DataListeIntegrerDétailLigne.Columns.Item(i).Visible = True Then
                    IligneCV += 1
                    .DGVCV.RowCount = IligneCV
                    .DGVCV.Rows(IligneCV - 1).Cells(0).Value = DataListeIntegrerDétailLigne.Columns.Item(i).HeaderText.Trim
                    .DGVCV.Rows(IligneCV - 1).Cells(1).Value = i
                Else
                    IligneCE += 1
                    .DGVCE.RowCount = IligneCE
                    .DGVCE.Rows(IligneCE - 1).Cells(0).Value = DataListeIntegrerDétailLigne.Columns.Item(i).HeaderText.Trim
                    .DGVCE.Rows(IligneCE - 1).Cells(1).Value = i
                End If
            Next i
            If Vsate = True Then
                .ShowDialog()
            End If
        End With
    End Sub
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
    Private Function RenvoieStockNegatif() As Boolean
        Dim DossierAdap As OleDbDataAdapter
        Dim DossierDs As DataSet
        Dim DossierTab As DataTable
        DossierAdap = New OleDbDataAdapter("select * from P_PREFERENCES WHERE PR_StockNeg=1", OleSocieteConnect)
        DossierDs = New DataSet
        DossierAdap.Fill(DossierDs)
        DossierTab = DossierDs.Tables(0)
        If DossierTab.Rows.Count <> 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Sub Creation_Entete_Document(ByRef typedoc)
        Dim OleRecherAdapter As OleDbDataAdapter = Nothing
        Dim OleRecherDataset As DataSet = Nothing
        Dim OleRechDatable As DataTable = Nothing
        Dim fournisseurAdap As OleDbDataAdapter = Nothing
        Dim fournisseurDs As DataSet = Nothing
        Dim fournisseurTab As DataTable = Nothing
        Dim statistAdap As OleDbDataAdapter
        Dim statistDs As DataSet
        Dim statistTab As DataTable
        Dim OleAdaptaterDelete As OleDbDataAdapter = Nothing
        Dim OleDeleteDataset As DataSet = Nothing
        Dim OledatableDelete As DataTable = Nothing
        Dim DossierAdap As OleDbDataAdapter
        Dim DossierDs As DataSet
        Dim DossierTab As DataTable
        DeviseTiers = Nothing
        FormatQte = 0
        Dim FormatMnt As Integer = 0

        Dim FormatDatefichier As Object = ComboDate.Text
        Dim PieceAutomatique As String = ""
        Dim IdentifiantCommande As String = ""

        ListeReliquat = New List(Of String)
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
        Select Case typedoc
            Case "3"
                Documents = BaseCial.FactoryDocumentVente.CreateType(DocumentType.DocumentTypeVenteLivraison)
            Case "14"
                Documents = BaseCial.FactoryDocumentAchat.CreateType(DocumentType.DocumentTypeAchatReprise)
            Case "23"
                Documents = BaseCial.FactoryDocumentVente.CreateType(DocumentType.DocumentTypeStockVirement)
        End Select

        With Documents
            If Trim(EnteteCatégorieComptable) <> Nothing Then
                If BaseCial.FactoryCategorieComptaVente.ExistIntitule(Trim(EnteteCatégorieComptable)) = True Then
                    .CategorieCompta = BaseCial.FactoryCategorieComptaVente.ReadIntitule(Trim(EnteteCatégorieComptable))
                End If
            End If
            If Trim(EnteteCatégorietarifaire) <> Nothing Then
                If BaseCial.FactoryCategorieTarif.ExistIntitule(Trim(EnteteCatégorietarifaire)) = True Then
                    .CategorieTarif = BaseCial.FactoryCategorieTarif.ReadIntitule(Trim(EnteteCatégorietarifaire))
                End If
            End If
            If Trim(EnteteFraisExpedition) <> Nothing Then
                If EstNumeric(Trim(EnteteFraisExpedition), DecimalNomb, DecimalMone) = True Then
                    .FraisExpedition = CDbl(RenvoiMontant(Trim(EnteteFraisExpedition), FormatMnt, DecimalNomb, DecimalMone))
                End If
            End If
            If Trim(EntetePlanAnalytique) <> Nothing Then
                If Trim(LigneCodeAffaire) <> "" Then
                    If BaseCpta.FactoryAnalytique.ExistIntitule(Trim(EntetePlanAnalytique)) = True Then
                        PlanAna = BaseCpta.FactoryAnalytique.ReadIntitule(Trim(EntetePlanAnalytique))
                        If BaseCpta.FactoryCompteA.ExistNumero(PlanAna, Trim(LigneCodeAffaire)) = True Then
                            .CompteA = BaseCpta.FactoryCompteA.ReadNumero(PlanAna, Trim(LigneCodeAffaire))
                        End If
                    End If
                End If
            End If
            If Trim(EnteteCompteGeneral) <> Nothing Then
                If BaseCpta.FactoryCompteG.ExistNumero(Trim(EnteteCompteGeneral)) = True Then
                    .CompteG = BaseCpta.FactoryCompteG.ReadNumero(Trim(EnteteCompteGeneral))
                End If
            End If
            If Trim(EcheanceConditionPaiement) <> Nothing Then
                If Trim(EcheanceConditionPaiement) = "0" Then
                    .ConditionPaiement = DocumentConditionPaiementType.DocumentConditionPaiementTiers
                Else
                    If Trim(EcheanceConditionPaiement) = "1" Then
                        .ConditionPaiement = DocumentConditionPaiementType.DocumentConditionPaiementSaisie
                    Else
                        If Trim(EcheanceConditionPaiement) = "2" Then
                            .ConditionPaiement = DocumentConditionPaiementType.DocumentConditionPaiementModele
                        Else
                            .ConditionPaiement = DocumentConditionPaiementType.DocumentConditionPaiementTiers
                        End If
                    End If
                End If
            End If

            If Trim(EcheanceModeleReglement) <> Nothing Then
                If BaseCpta.FactoryModeleReglement.ExistIntitule(Trim(EcheanceModeleReglement)) = True Then
                    .ModeleReglement = BaseCpta.FactoryModeleReglement.ReadIntitule(Trim(EcheanceModeleReglement))
                End If
            End If
            If Trim(EnteteConditiondeLivraison) <> Nothing Then
                If BaseCial.FactoryConditionLivraison.ExistIntitule(Trim(EnteteConditiondeLivraison)) = True Then
                    .ConditionLivraison = BaseCial.FactoryConditionLivraison.ReadIntitule(Trim(EnteteConditiondeLivraison))
                End If
            End If
            'traitement depot
            If Trim(EnteteIntituleDepot) <> Nothing Then
                If BaseCial.FactoryDepot.ExistIntitule(Trim(EnteteIntituleDepot)) = True Then
                    .DepotStockage = BaseCial.FactoryDepot.ReadIntitule(Trim(EnteteIntituleDepot))
                End If
            End If
            If Trim(IDDepotEntete) <> Nothing Then
                If IsNumeric(Trim(IDDepotEntete)) = True Then
                    statistAdap = New OleDbDataAdapter("select * from F_DEPOT where DE_No ='" & CInt(Trim(IDDepotEntete)) & "'", OleSocieteConnect)
                    statistDs = New DataSet
                    statistAdap.Fill(statistDs)
                    statistTab = statistDs.Tables(0)
                    If statistTab.Rows.Count <> 0 Then
                        .DepotStockage = BaseCial.FactoryDepot.ReadIntitule(Trim(statistTab.Rows(0).Item("DE_Intitule")))
                    End If
                Else
                    If BaseCial.FactoryDepot.ExistIntitule(Trim(IDDepotEntete)) = True Then
                        .DepotStockage = BaseCial.FactoryDepot.ReadIntitule(Trim(IDDepotEntete))
                    End If
                End If
            End If
            'If Datagridaffiche.Columns.Contains("EnteteIntituleDepot") = False And Datagridaffiche.Columns.Contains("IDDepotEntete") = False Then
            '    If BaseCial.FactoryDepot.ExistIntitule(RenvoieDepotPrincipal) = True Then
            '        .DepotStockage = BaseCial.FactoryDepot.ReadIntitule(RenvoieDepotPrincipal)
            '    End If
            'Else
            '    If Trim(EnteteIntituleDepot) = "" Then
            '        If Trim(IDDepotEntete) = "" Then
            '            If BaseCial.FactoryDepot.ExistIntitule(RenvoieDepotPrincipal) = True Then
            '                .DepotStockage = BaseCial.FactoryDepot.ReadIntitule(RenvoieDepotPrincipal)
            '            End If
            '        End If
            '    End If
            'End If
            If Trim(EnteteBLFacture) <> Nothing Then
                If Trim(EnteteBLFacture) = "1" Then
                    .DO_BLFact = True
                Else
                    .DO_BLFact = False
                End If
            End If
            If Trim(EnteteColisagedeLivraison) <> Nothing Then
                If EstNumeric(Trim(EnteteColisagedeLivraison), DecimalNomb, DecimalMone) = True Then
                    .DO_Colisage = CInt(RenvoiTaux(Trim(EnteteColisagedeLivraison), DecimalNomb, DecimalMone))
                End If
            End If
            If Trim(EnteteContact) <> Nothing Then
                .DO_Contact = EnteteContact
            End If
            If Trim(Entete1) <> Nothing Then
                .DO_Coord(1) = Entete1
            End If
            If Trim(Entete2) <> Nothing Then
                .DO_Coord(2) = Entete2
            End If
            If Trim(Entete3) <> Nothing Then
                .DO_Coord(3) = Entete3
            End If
            If Trim(Entete4) <> Nothing Then
                .DO_Coord(4) = Entete4
            End If
            If Trim(EnteteCours) <> Nothing Then
                If EstNumeric(Trim(EnteteCours), DecimalNomb, DecimalMone) = True Then
                    .DO_Cours = CDbl(RenvoiTaux(Trim(EnteteCours), DecimalNomb, DecimalMone))
                End If
            End If
            If Trim(EnteteDateDocument) <> Nothing Then
                If Trim(EnteteDateDocument) <> Nothing Then
                    .DO_Date = RenvoieDateValide(Trim(EnteteDateDocument), FormatDatefichier)
                End If
            End If
            If Trim(EnteteDateLivraison) <> Nothing Then
                If Trim(EnteteDateLivraison) <> Nothing Then
                    .DO_DateLivr = RenvoieDateValide(Trim(EnteteDateLivraison), FormatDatefichier)
                End If
            End If
            If Trim(EnteteEcartValorisation) <> Nothing Then
                If EstNumeric(Trim(EnteteEcartValorisation), DecimalNomb, DecimalMone) = True Then
                    .DO_Ecart = CDbl(RenvoiTaux(Trim(EnteteEcartValorisation), DecimalNomb, DecimalMone))
                End If
            End If
            If Trim(EnteteLangue) <> Nothing Then
                If Trim(EnteteLangue) = "0" Then
                    .DO_Langue = LangueType.LangueTypeAucune
                Else
                    If Trim(EnteteLangue) = "1" Then
                        .DO_Langue = LangueType.LangueTypeLangue1
                    Else
                        If Trim(EnteteLangue) = "2" Then
                            .DO_Langue = LangueType.LangueTypeLangue2
                        End If
                    End If
                End If
            End If
            If Trim(EnteteNombredeFacture) <> Nothing Then
                If EstNumeric(Trim(EnteteNombredeFacture), DecimalNomb, DecimalMone) = True Then
                    .DO_NbFacture = CInt(RenvoiTaux(Trim(EnteteNombredeFacture), DecimalNomb, DecimalMone))
                End If
            End If
            If Trim(EnteteRegimeDocument) <> Nothing Then
                If EstNumeric(Trim(EnteteRegimeDocument), DecimalNomb, DecimalMone) = True Then
                    .DO_Regime = CInt(RenvoiTaux(Trim(EnteteRegimeDocument), DecimalNomb, DecimalMone))
                End If
            End If
            If ChkPieceAuto.Checked = False Then
                If Trim(EnteteSoucheDocument) <> "" Then
                    If Trim(EnteteSoucheDocument) <> "" Then
                        If EstNumeric(Trim(EnteteSoucheDocument), DecimalNomb, DecimalMone) = True Then
                            statistAdap = New OleDbDataAdapter("select * from P_SOUCHEVENTE where cbIndice ='" & CInt(CInt(RenvoiTaux(Trim(EnteteSoucheDocument), DecimalNomb, DecimalMone)) + 1) & "'", OleSocieteConnect)
                            statistDs = New DataSet
                            statistAdap.Fill(statistDs)
                            statistTab = statistDs.Tables(0)
                            If statistTab.Rows.Count <> 0 Then
                                If BaseCial.FactorySoucheVente.ReadIntitule(statistTab.Rows(0).Item("S_Intitule")).IsValide = True Then
                                    If typedoc = "0" Then
                                        .DO_Souche = CInt(RenvoiTaux(Trim(EnteteSoucheDocument), DecimalNomb, DecimalMone))
                                    Else
                                        If typedoc = "1" Then
                                            .DO_Souche = CInt(RenvoiTaux(Trim(EnteteSoucheDocument), DecimalNomb, DecimalMone))
                                        Else
                                            If typedoc = "3" Then
                                                .DO_Souche = CInt(RenvoiTaux(Trim(EnteteSoucheDocument), DecimalNomb, DecimalMone))
                                            Else
                                                If typedoc = "4" Then
                                                    .DO_Souche = CInt(RenvoiTaux(Trim(EnteteSoucheDocument), DecimalNomb, DecimalMone))
                                                Else
                                                    If typedoc = "5" Then
                                                        .DO_Souche = CInt(RenvoiTaux(Trim(EnteteSoucheDocument), DecimalNomb, DecimalMone))
                                                    Else
                                                        If typedoc = "6" Then
                                                            .DO_Souche = CInt(RenvoiTaux(Trim(EnteteSoucheDocument), DecimalNomb, DecimalMone))
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Else
                            If BaseCial.FactorySoucheVente.ExistIntitule(Trim(EnteteSoucheDocument)) = True Then
                                statistAdap = New OleDbDataAdapter("select * from P_SOUCHEVENTE where S_Intitule ='" & Join(Split(Trim(EnteteSoucheDocument), "'"), "''") & "'", OleSocieteConnect)
                                statistDs = New DataSet
                                statistAdap.Fill(statistDs)
                                statistTab = statistDs.Tables(0)
                                If statistTab.Rows.Count <> 0 Then
                                    If CInt(statistTab.Rows(0).Item("cbIndice")) - 1 >= 0 Then
                                        If BaseCial.FactorySoucheVente.ReadIntitule(Trim(EnteteSoucheDocument)).IsValide = True Then
                                            If typedoc = "0" Then
                                                .DO_Souche = CInt(statistTab.Rows(0).Item("cbIndice")) - 1
                                            Else
                                                If typedoc = "1" Then
                                                    .DO_Souche = CInt(statistTab.Rows(0).Item("cbIndice")) - 1
                                                Else
                                                    If typedoc = "3" Then
                                                        .DO_Souche = CInt(statistTab.Rows(0).Item("cbIndice")) - 1
                                                    Else
                                                        If typedoc = "4" Then
                                                            .DO_Souche = CInt(statistTab.Rows(0).Item("cbIndice")) - 1
                                                        Else
                                                            'Hermann importation des avoir financie Type (Doc=5 or 6)
                                                            If typedoc = "6" Or typedoc = "5" Then
                                                                .DO_Souche = CInt(statistTab.Rows(0).Item("cbIndice")) - 1
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
                End If
                .DO_Piece = Strings.UCase(EntetePieceInterne)
            Else
                If Trim(EnteteSoucheDocument) <> "" Then
                    If Trim(EnteteSoucheDocument) <> "" Then
                        If EstNumeric(Trim(EnteteSoucheDocument), DecimalNomb, DecimalMone) = True Then
                            statistAdap = New OleDbDataAdapter("select * from P_SOUCHEVENTE where cbIndice ='" & CInt(CInt(RenvoiTaux(Trim(EnteteSoucheDocument), DecimalNomb, DecimalMone)) + 1) & "'", OleSocieteConnect)
                            statistDs = New DataSet
                            statistAdap.Fill(statistDs)
                            statistTab = statistDs.Tables(0)
                            If statistTab.Rows.Count <> 0 Then
                                If BaseCial.FactorySoucheVente.ReadIntitule(statistTab.Rows(0).Item("S_Intitule")).IsValide = True Then
                                    If typedoc = "0" Then
                                        .DO_Souche = CInt(RenvoiTaux(Trim(EnteteSoucheDocument), DecimalNomb, DecimalMone))
                                        .DO_Piece = BaseCial.FactorySoucheVente.ReadIntitule(Trim(statistTab.Rows(0).Item("S_Intitule"))).ReadCurrentPiece(DocumentType.DocumentTypeVenteDevis, DocumentProvenanceType.DocProvenanceNormale)
                                    Else
                                        If typedoc = "1" Then
                                            .DO_Souche = CInt(RenvoiTaux(Trim(EnteteSoucheDocument), DecimalNomb, DecimalMone))
                                            .DO_Piece = BaseCial.FactorySoucheVente.ReadIntitule(Trim(statistTab.Rows(0).Item("S_Intitule"))).ReadCurrentPiece(DocumentType.DocumentTypeVenteCommande, DocumentProvenanceType.DocProvenanceNormale)
                                        Else
                                            If typedoc = "3" Then
                                                .DO_Souche = CInt(RenvoiTaux(Trim(EnteteSoucheDocument), DecimalNomb, DecimalMone))
                                                .DO_Piece = BaseCial.FactorySoucheVente.ReadIntitule(Trim(statistTab.Rows(0).Item("S_Intitule"))).ReadCurrentPiece(DocumentType.DocumentTypeVenteLivraison, DocumentProvenanceType.DocProvenanceNormale)
                                            Else
                                                If typedoc = "4" Then
                                                    .DO_Souche = CInt(RenvoiTaux(Trim(EnteteSoucheDocument), DecimalNomb, DecimalMone))
                                                    .DO_Piece = BaseCial.FactorySoucheVente.ReadIntitule(Trim(statistTab.Rows(0).Item("S_Intitule"))).ReadCurrentPiece(DocumentType.DocumentTypeVenteReprise, DocumentProvenanceType.DocProvenanceNormale)
                                                Else
                                                    If typedoc = "5" Then
                                                        .DO_Souche = CInt(RenvoiTaux(Trim(EnteteSoucheDocument), DecimalNomb, DecimalMone))
                                                        .DO_Piece = BaseCial.FactorySoucheVente.ReadIntitule(Trim(statistTab.Rows(0).Item("S_Intitule"))).ReadCurrentPiece(DocumentType.DocumentTypeVenteAvoir, DocumentProvenanceType.DocProvenanceNormale)
                                                    Else
                                                        If typedoc = "6" Then
                                                            .DO_Souche = CInt(RenvoiTaux(Trim(EnteteSoucheDocument), DecimalNomb, DecimalMone))
                                                            .DO_Piece = BaseCial.FactorySoucheVente.ReadIntitule(Trim(statistTab.Rows(0).Item("S_Intitule"))).ReadCurrentPiece(DocumentType.DocumentTypeVenteFacture, DocumentProvenanceType.DocProvenanceNormale)
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Else
                            If BaseCial.FactorySoucheVente.ExistIntitule(Trim(EnteteSoucheDocument)) = True Then
                                statistAdap = New OleDbDataAdapter("select * from P_SOUCHEVENTE where S_Intitule ='" & Join(Split(Trim(EnteteSoucheDocument), "'"), "''") & "'", OleSocieteConnect)
                                statistDs = New DataSet
                                statistAdap.Fill(statistDs)
                                statistTab = statistDs.Tables(0)
                                If statistTab.Rows.Count <> 0 Then
                                    If CInt(statistTab.Rows(0).Item("cbIndice")) - 1 >= 0 Then
                                        If BaseCial.FactorySoucheVente.ReadIntitule(Trim(EnteteSoucheDocument)).IsValide = True Then
                                            If typedoc = "0" Then
                                                .DO_Souche = CInt(statistTab.Rows(0).Item("cbIndice")) - 1
                                                .DO_Piece = BaseCial.FactorySoucheVente.ReadIntitule(Trim(EnteteSoucheDocument)).ReadCurrentPiece(DocumentType.DocumentTypeVenteDevis, DocumentProvenanceType.DocProvenanceNormale)
                                            Else
                                                If typedoc = "1" Then
                                                    .DO_Souche = CInt(statistTab.Rows(0).Item("cbIndice")) - 1
                                                    .DO_Piece = BaseCial.FactorySoucheVente.ReadIntitule(Trim(EnteteSoucheDocument)).ReadCurrentPiece(DocumentType.DocumentTypeVenteCommande, DocumentProvenanceType.DocProvenanceNormale)
                                                Else
                                                    If typedoc = "3" Then
                                                        .DO_Souche = CInt(statistTab.Rows(0).Item("cbIndice")) - 1
                                                        .DO_Piece = BaseCial.FactorySoucheVente.ReadIntitule(Trim(EnteteSoucheDocument)).ReadCurrentPiece(DocumentType.DocumentTypeVenteLivraison, DocumentProvenanceType.DocProvenanceNormale)
                                                    Else
                                                        If typedoc = "4" Then
                                                            .DO_Souche = CInt(statistTab.Rows(0).Item("cbIndice")) - 1
                                                            .DO_Piece = BaseCial.FactorySoucheVente.ReadIntitule(Trim(EnteteSoucheDocument)).ReadCurrentPiece(DocumentType.DocumentTypeVenteReprise, DocumentProvenanceType.DocProvenanceNormale)
                                                        Else
                                                            If typedoc = "5" Then
                                                                .DO_Souche = CInt(statistTab.Rows(0).Item("cbIndice")) - 1
                                                                .DO_Piece = BaseCial.FactorySoucheVente.ReadIntitule(Trim(EnteteSoucheDocument)).ReadCurrentPiece(DocumentType.DocumentTypeVenteAvoir, DocumentProvenanceType.DocProvenanceNormale)
                                                            Else
                                                                If typedoc = "6" Then
                                                                    .DO_Souche = CInt(statistTab.Rows(0).Item("cbIndice")) - 1
                                                                    If Datagridaffiche.Columns.Contains("ProvenanceFacture") = True Then
                                                                        If Trim(ProvenanceFacture) = "1" Then
                                                                            .DO_Piece = BaseCial.FactorySoucheVente.ReadIntitule(Trim(EnteteSoucheDocument)).ReadCurrentPiece(DocumentType.DocumentTypeVenteFacture, DocumentProvenanceType.DocProvenanceRetour)
                                                                        Else
                                                                            If Trim(ProvenanceFacture) = "2" Then
                                                                                .DO_Piece = BaseCial.FactorySoucheVente.ReadIntitule(Trim(EnteteSoucheDocument)).ReadCurrentPiece(DocumentType.DocumentTypeVenteFacture, DocumentProvenanceType.DocProvenanceAvoir)
                                                                            Else
                                                                                If Trim(ProvenanceFacture) = "3" Then
                                                                                    .DO_Piece = BaseCial.FactorySoucheVente.ReadIntitule(Trim(EnteteSoucheDocument)).ReadCurrentPiece(DocumentType.DocumentTypeVenteFacture, DocumentProvenanceType.DocProvenanceTicket)
                                                                                Else
                                                                                    If Trim(ProvenanceFacture) = "4" Then
                                                                                        .DO_Piece = BaseCial.FactorySoucheVente.ReadIntitule(Trim(EnteteSoucheDocument)).ReadCurrentPiece(DocumentType.DocumentTypeVenteFacture, DocumentProvenanceType.DocProvenanceRectif)
                                                                                    Else
                                                                                        If Trim(ProvenanceFacture) = "0" Then
                                                                                            .DO_Piece = BaseCial.FactorySoucheVente.ReadIntitule(Trim(EnteteSoucheDocument)).ReadCurrentPiece(DocumentType.DocumentTypeVenteFacture, DocumentProvenanceType.DocProvenanceNormale)
                                                                                        Else
                                                                                            .DO_Piece = BaseCial.FactorySoucheVente.ReadIntitule(Trim(EnteteSoucheDocument)).ReadCurrentPiece(DocumentType.DocumentTypeVenteFacture, DocumentProvenanceType.DocProvenanceNormale)
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    Else
                                                                        .DO_Piece = BaseCial.FactorySoucheVente.ReadIntitule(Trim(EnteteSoucheDocument)).ReadCurrentPiece(DocumentType.DocumentTypeVenteFacture, DocumentProvenanceType.DocProvenanceNormale)
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
                        End If
                    End If
                End If
            End If
            EnteteStatutdocument = CmbStatut.Text
            If Trim(EnteteStatutdocument) <> "" Then
                If Trim(EnteteStatutdocument) = "Saisi" Then
                    .DO_Statut = DocumentStatutType.DocumentStatutTypeSaisie
                Else
                    If Trim(EnteteStatutdocument) = "Confirmé" Then
                        .DO_Statut = DocumentStatutType.DocumentStatutTypeConfirme
                    Else
                        If Trim(EnteteStatutdocument) = "A préparer" Then
                            .DO_Statut = DocumentStatutType.DocumentStatutTypeAPrepare
                        Else
                            If Trim(EnteteStatutdocument) <> "Réceptionné" Then
                                .DO_Statut = DocumentStatutType.DocumentStatutTypeAPrepare
                            End If
                        End If
                    End If
                End If
            Else
                .DO_Statut = DocumentStatutType.DocumentStatutTypeAPrepare
            End If
            If Trim(EnteteNatureTransaction) <> "" Then
                If EstNumeric(Trim(EnteteNatureTransaction), DecimalNomb, DecimalMone) = True Then
                    .DO_Transaction = CInt(RenvoiTaux(Trim(EnteteNatureTransaction), DecimalNomb, DecimalMone))
                End If
            End If
            If Trim(EnteteTauxescompte) <> "" Then
                If EstNumeric(Trim(EnteteTauxescompte), DecimalNomb, DecimalMone) = True Then
                    .DO_TxEscompte = CDbl(RenvoiMontant(Trim(EnteteTauxescompte), FormatMnt, DecimalNomb, DecimalMone))
                End If
            End If
            If Trim(EnteteIntituleExpédition) <> "" Then
                If BaseCial.FactoryExpedition.ExistIntitule(Trim(EnteteIntituleExpédition)) = True Then
                    .Expedition = BaseCial.FactoryExpedition.ReadIntitule(Trim(EnteteIntituleExpédition))
                End If
            End If
            If Trim(EnteteNomReprésentant) <> "" Then
                If Trim(EntetePrenomReprésentant) <> "" Then
                    If BaseCpta.FactoryCollaborateur.ExistNomPrenom(Trim(EnteteNomReprésentant), Trim(EntetePrenomReprésentant)) = True Then
                        .Collaborateur = BaseCpta.FactoryCollaborateur.ReadNomPrenom(Trim(EnteteNomReprésentant), Trim(EntetePrenomReprésentant))
                    End If
                End If
            End If
            If Trim(EnteteCodeTiers) <> "" Then
                If BaseCpta.FactoryClient.ExistNumero(Trim(EnteteCodeTiers)) = True Then
                    .SetDefaultClient(BaseCpta.FactoryClient.ReadNumero(Trim(EnteteCodeTiers)))
                    If IsNothing(BaseCpta.FactoryClient.ReadNumero(Trim(EnteteCodeTiers)).Devise) = False Then
                        DeviseTiers = BaseCpta.FactoryClient.ReadNumero(Trim(EnteteCodeTiers)).Devise.D_Intitule
                    End If
                End If
            End If
            If Trim(EnteteCodeTiersPayeur) <> "" Then
                If BaseCpta.FactoryClient.ExistNumero(Trim(EnteteCodeTiersPayeur)) = True Then
                    .TiersPayeur = BaseCpta.FactoryClient.ReadNumero(Trim(EnteteCodeTiersPayeur))
                End If
            End If

            If Trim(EnteteUniteColis) <> "" Then
                If Trim(EnteteUniteColis) <> "" Then
                    If BaseCial.FactoryUnite.ExistIntitule(Trim(EnteteUniteColis)) = True Then
                        .Unite = BaseCial.FactoryUnite.ReadIntitule(Trim(EnteteUniteColis))
                    End If
                End If
            End If
            If Trim(EnteteReference) <> "" Then
                .DO_Ref = EnteteReference
            End If
            .Write()
            .CouldModified()
            Try
                If Trim(EcheanceModeleReglement) <> "" Then
                    If BaseCpta.FactoryModeleReglement.ExistIntitule(Trim(EcheanceModeleReglement)) = True Then
                        .ModeleReglement = BaseCpta.FactoryModeleReglement.ReadIntitule(Trim(EcheanceModeleReglement))
                    End If
                    .Write()
                End If
                If Trim(EcheanceConditionPaiement) <> "" Then
                    If Trim(EcheanceConditionPaiement) = "0" Then
                        .ConditionPaiement = DocumentConditionPaiementType.DocumentConditionPaiementTiers
                    Else
                        If Trim(EcheanceConditionPaiement) = "1" Then
                            .ConditionPaiement = DocumentConditionPaiementType.DocumentConditionPaiementSaisie
                        Else
                            If Trim(EcheanceConditionPaiement) = "2" Then
                                .ConditionPaiement = DocumentConditionPaiementType.DocumentConditionPaiementModele
                            Else
                                .ConditionPaiement = DocumentConditionPaiementType.DocumentConditionPaiementTiers
                            End If
                        End If
                    End If
                    .Write()
                End If

                If Trim(EnteteCompteGeneral) <> "" Then
                    If BaseCpta.FactoryCompteG.ExistNumero(Trim(EnteteCompteGeneral)) = True Then
                        .CompteG = BaseCpta.FactoryCompteG.ReadNumero(Trim(EnteteCompteGeneral))
                        .Write()
                    End If
                End If
                If Trim(EnteteFraisExpedition) <> "" Then
                    If EstNumeric(Trim(EnteteFraisExpedition), DecimalNomb, DecimalMone) = True Then
                        .FraisExpedition = CDbl(RenvoiMontant(Trim(EnteteFraisExpedition), FormatMnt, DecimalNomb, DecimalMone))
                        .Write()
                    End If
                End If
                If Trim(EnteteNomReprésentant) <> "" Then
                    If Trim(EntetePrenomReprésentant) <> "" Then
                        If BaseCpta.FactoryCollaborateur.ExistNomPrenom(Trim(EnteteNomReprésentant), Trim(EntetePrenomReprésentant)) = True Then
                            .Collaborateur = BaseCpta.FactoryCollaborateur.ReadNomPrenom(Trim(EnteteNomReprésentant), Trim(EntetePrenomReprésentant))
                            .Write()
                        End If
                    End If
                End If
                If Trim(EnteteCatégorieComptable) <> "" Then
                    If BaseCial.FactoryCategorieComptaVente.ExistIntitule(Trim(EnteteCatégorieComptable)) = True Then
                        .CategorieCompta = BaseCial.FactoryCategorieComptaVente.ReadIntitule(Trim(EnteteCatégorieComptable))
                        .Write()
                    End If
                End If
                If Trim(EnteteCatégorietarifaire) <> "" Then
                    If BaseCial.FactoryCategorieTarif.ExistIntitule(Trim(EnteteCatégorietarifaire)) = True Then
                        .CategorieTarif = BaseCial.FactoryCategorieTarif.ReadIntitule(Trim(EnteteCatégorietarifaire))
                        .Write()
                    End If
                End If
                If Trim(EnteteIntituleDevise) <> "" Then
                    If BaseCpta.FactoryDevise.ExistIntitule(Trim(EnteteIntituleDevise)) = True Then
                        .Devise = BaseCpta.FactoryDevise.ReadIntitule(Trim(EnteteIntituleDevise))
                        .Write()
                    End If
                Else
                    If Trim(LignePrixUnitaire) <> "" Then
                        If Trim(LignePrixUnitaireDevise) <> "" Then
                            If BaseCpta.FactoryDevise.ExistIntitule(Trim(DeviseTiers)) = True Then
                                .Devise = BaseCpta.FactoryDevise.ReadIntitule(Trim(DeviseTiers))
                                .Write()
                            End If
                        End If
                    End If
                End If
                If Trim(EnteteIntituleExpédition) <> "" Then
                    If BaseCial.FactoryExpedition.ExistIntitule(Trim(EnteteIntituleExpédition)) = True Then
                        .Expedition = BaseCial.FactoryExpedition.ReadIntitule(Trim(EnteteIntituleExpédition))
                        .Write()
                    End If
                End If
                If InStr(IdentifiantCommande, ",") <> 0 Then
                    If Datagridaffiche.Columns.Contains(Strings.Left(IdentifiantCommande, InStr(IdentifiantCommande, ",") - 1)) = True Then
                        Dim OleAdaptaterCa As OleDbDataAdapter
                        Dim OleCaDataset As DataSet
                        Dim OledatableCa As DataTable
                        OleAdaptaterCa = New OleDbDataAdapter("select * from COLIMPMOUV WHERE Libelle='" & Trim(Strings.Left(IdentifiantCommande, InStr(IdentifiantCommande, ",") - 1)) & "' And Fichier='" & Trim(Strings.Right(IdentifiantCommande, Len(IdentifiantCommande) - InStr(IdentifiantCommande, ","))) & "'", OleConnenection)
                        OleCaDataset = New DataSet
                        OleAdaptaterCa.Fill(OleCaDataset)
                        OledatableCa = OleCaDataset.Tables(0)
                        If OledatableCa.Rows.Count <> 0 Then
                            Try
                                Dim OleDocAdapter As OleDbDataAdapter
                                Dim OleDocDataset As DataSet
                                Dim OleDocDatable As DataTable

                                OleDocAdapter = New OleDbDataAdapter("Select  * From F_DOCENTETE WHERE " & OledatableCa.Rows(0).Item("Champ") & " ='" & Join(Split(Trim(PieceCommande), "'"), "''") & "' And DO_Type=1 And DO_Domaine=0", OleSocieteConnect)
                                OleDocDataset = New DataSet
                                OleDocAdapter.Fill(OleDocDataset)
                                OleDocDatable = OleDocDataset.Tables(0)
                                If OleDocDatable.Rows.Count = 1 Then
                                    Dim Reliquat As IBODocumentVente3 = Nothing
                                    If BaseCial.FactoryDocumentVente.ExistPiece(DocumentType.DocumentTypeVenteCommande, Join(Split(Trim(PieceCommande), "'"), "''")) = True Then
                                        Reliquat = BaseCial.FactoryDocumentVente.ReadPiece(DocumentType.DocumentTypeVenteCommande, Join(Split(Trim(PieceCommande), "'"), "''"))
                                    Else
                                        If BaseCial.FactoryDocumentVente.ExistPiece(DocumentType.DocumentTypeVenteCommande, Join(Split(Trim(PieceCommande), "'"), "''")) = True Then
                                            Reliquat = BaseCial.FactoryDocumentVente.ReadPiece(DocumentType.DocumentTypeVenteCommande, Join(Split(Trim(PieceCommande), "'"), "''"))
                                        End If
                                    End If
                                    'verification du champs reliquat hermann
                                    Dim QueryUpdateChampDo_Reliquat As String = "UPDATE F_DOCENTETE SET Do_Reliquat=1 WHERE DO_PIECE='" & PieceCommande & "'"
                                    Dim Commande As OleDbCommand
                                    If BaseCial.FactoryDocumentVente.ExistPiece(DocumentType.DocumentTypeVenteCommande, Join(Split(Trim(PieceCommande), "'"), "''")) = True Then
                                        Try
                                            Commande = New OleDbCommand(QueryUpdateChampDo_Reliquat, OleSocieteConnect)
                                            Commande.ExecuteNonQuery()
                                        Catch ex As Exception
                                        End Try
                                    End If
                                    If IsNothing(Reliquat) = False Then
                                        If Trim(EnteteIntituleDevise) <> "" Then
                                            .Devise = Reliquat.Devise
                                        End If
                                        If Datagridaffiche.Columns.Contains("Entete1") = False Then
                                            .DO_Coord(1) = Reliquat.DO_Coord(1)
                                        End If
                                        If Datagridaffiche.Columns.Contains("Entete2") = False Then
                                            .DO_Coord(2) = Reliquat.DO_Coord(2)
                                        End If
                                        If Datagridaffiche.Columns.Contains("Entete3") = False Then
                                            .DO_Coord(3) = Reliquat.DO_Coord(3)
                                        End If
                                        If Datagridaffiche.Columns.Contains("Entete4") = False Then
                                            .DO_Coord(4) = Reliquat.DO_Coord(4)
                                        End If
                                        If Trim(EnteteNomReprésentant) <> "" Then
                                            If Trim(EntetePrenomReprésentant) <> "" Then
                                                .Collaborateur = Reliquat.Collaborateur
                                            End If
                                        End If
                                        'If Trim(LingeCodeAffaire) <> "" Then
                                        '    .CompteA = Reliquat.CompteA
                                        'End If
                                        If Trim(EnteteCatégorieComptable) <> "" Then
                                            .CategorieCompta = Reliquat.CategorieCompta
                                        End If
                                        If Trim(EnteteReference) <> "" Then
                                            .DO_Ref = Reliquat.DO_Ref
                                        End If
                                        If Trim(EnteteIntituleDepotClient) <> "" Then
                                            .LieuLivraison = Reliquat.LieuLivraison
                                        End If
                                        If Trim(EnteteIntituleExpédition) <> "" Then
                                            .Expedition = Reliquat.Expedition
                                        End If
                                        If Trim(EnteteConditiondeLivraison) <> "" Then
                                            .ConditionLivraison = Reliquat.ConditionLivraison
                                        End If
                                        If Trim(EnteteCodeTiersPayeur) <> "" Then
                                            .TiersPayeur = Reliquat.TiersPayeur
                                        End If
                                        For i As Integer = 1 To Reliquat.InfoLibre.Count
                                            .InfoLibre.Item(i) = Reliquat.InfoLibre(i)
                                        Next
                                        .Write()
                                    End If
                                End If
                            Catch ex As Exception
                            End Try
                        End If
                    End If
                End If
                If Trim(EnteteIntituleDepotClient) <> "" Then
                    Dim OleDocAdapter As OleDbDataAdapter
                    Dim OleDocDataset As DataSet
                    Dim OleDocDatable As DataTable
                    OleDocAdapter = New OleDbDataAdapter("Select  * From F_LIVRAISON WHERE LI_Intitule ='" & Join(Split(Trim(EnteteIntituleDepotClient), "'"), "''") & "' And CT_Num ='" & Join(Split(Trim(Documents.Client.CT_Num), "'"), "''") & "'", OleSocieteConnect)
                    OleDocDataset = New DataSet
                    OleDocAdapter.Fill(OleDocDataset)
                    OleDocDatable = OleDocDataset.Tables(0)
                    If OleDocDatable.Rows.Count <> 0 Then
                        Try
                            Dim VClient As IBOClient3
                            VClient = BaseCpta.FactoryClient.ReadNumero(Join(Split(OleDocDatable.Rows(0).Item("CT_Num"), ","), "."))
                            For Each VDepotClient As IBOClientLivraison3 In VClient.FactoryClientLivraison.List
                                If VDepotClient.LI_Intitule = OleDocDatable.Rows(0).Item("LI_Intitule") Then
                                    .LieuLivraison = VDepotClient
                                    .Write()
                                    Exit For
                                End If
                            Next
                        Catch ex As Exception
                        End Try
                    End If
                Else
                    'ErreurJrn.WriteLine("Pièce N° : " & Document.DO_Piece & ", Le dépôt Client provenant du fichier à traiter est vide")
                End If
                ErreurJrn.WriteLine("-----------------------------------------------------------------------------------------------------")
                ErreurJrn.WriteLine("")
                If typedoc = "3" Then
                    ErreurJrn.WriteLine("Bon de Livraison N° : " & Trim(Documents.DO_Piece) & " Créé Pour la pièce N° :" & Trim(EntetePieceInterne))
                Else
                    If typedoc = "14" Then
                        ErreurJrn.WriteLine("Bon de Retour N° : " & Trim(Documents.DO_Piece) & " Créé Pour la pièce N° :" & Trim(EntetePieceInterne))
                    Else
                        If typedoc = "23" Then
                            ErreurJrn.WriteLine("Transfert de Dépot à Dépot du  N° : " & Trim(Documents.DO_Piece) & " Créé Pour la pièce N° :" & Trim(EntetePieceInterne))
                        End If
                    End If
                End If
                'Application des mise a jour du cours (parite) du document a importe  hermann
                If Trim(EnteteCours) <> "" Then
                    If EstNumeric(Trim(EnteteCours), DecimalNomb, DecimalMone) = True Then
                        .DO_Cours = CDbl(RenvoiTaux(Trim(EnteteCours), DecimalNomb, DecimalMone))
                        .Write()
                    End If
                End If
            Catch ex As Exception
            End Try
            'Traitement des Infos Libres
            Try
                If 1 > 0 Then 'infoListe.Count > 0
                    'While infoListe.Count <> 0
                    '    OleAdaptaterDelete = New OleDbDataAdapter("select * From COLIMPMOUV where Libelle='" & Trim(Datagridaffiche.Columns(infoListe.Item(0)).Name) & "' And Libre=True", OleConnenection)
                    '    OleDeleteDataset = New DataSet
                    '    OleAdaptaterDelete.Fill(OleDeleteDataset)
                    '    OledatableDelete = OleDeleteDataset.Tables(0)
                    '    If OledatableDelete.Rows.Count <> 0 Then
                    '        'L'info Libre Parametrée par l'utilisateur existe dans Sage
                    '        If Documents.InfoLibre.Count <> 0 Then
                    '            If IsNothing(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) = False Then
                    '                If Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) <> "" Then
                    '                    statistAdap = New OleDbDataAdapter("select * from cbSysLibre where CB_File='F_DOCENTETE' and CB_Name ='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "'", OleSocieteConnect)
                    '                    statistDs = New DataSet
                    '                    statistAdap.Fill(statistDs)
                    '                    statistTab = statistDs.Tables(0)
                    '                    If statistTab.Rows.Count <> 0 Then
                    '                        'Texte
                    '                        If statistTab.Rows(0).Item("CB_Type") = 9 Then
                    '                            OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document' And Entete=True", OleConnenection)
                    '                            OleRecherDataset = New DataSet
                    '                            OleRecherAdapter.Fill(OleRecherDataset)
                    '                            OleRechDatable = OleRecherDataset.Tables(0)
                    '                            If OleRechDatable.Rows.Count <> 0 Then
                    '                                If statistTab.Rows(0).Item("CB_Len") >= Strings.Len(Trim(OleRechDatable.Rows(0).Item("Correspond"))) Then
                    '                                    DOCUMENT.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = Trim(OleRechDatable.Rows(0).Item("Correspond"))
                    '                                    DOCUMENT.Write()
                    '                                End If
                    '                            Else
                    '                                If statistTab.Rows(0).Item("CB_Len") >= Strings.Len(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value)) Then
                    '                                    DOCUMENT.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value)
                    '                                    DOCUMENT.Write()
                    '                                End If
                    '                            End If
                    '                        End If
                    '                        'Table
                    '                        If statistTab.Rows(0).Item("CB_Type") = 22 Then
                    '                            OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document' And Entete=True", OleConnenection)
                    '                            OleRecherDataset = New DataSet
                    '                            OleRecherAdapter.Fill(OleRecherDataset)
                    '                            OleRechDatable = OleRecherDataset.Tables(0)
                    '                            If OleRechDatable.Rows.Count <> 0 Then
                    '                                If statistTab.Rows(0).Item("CB_Len") > Strings.Len(Trim(OleRechDatable.Rows(0).Item("Correspond"))) Then
                    '                                    DOCUMENT.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = Trim(OleRechDatable.Rows(0).Item("Correspond"))
                    '                                    DOCUMENT.Write()
                    '                                End If
                    '                            Else
                    '                                If statistTab.Rows(0).Item("CB_Len") > Strings.Len(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value)) Then
                    '                                    DOCUMENT.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value)
                    '                                    DOCUMENT.Write()
                    '                                End If
                    '                            End If
                    '                        End If
                    '                        'Montant
                    '                        If statistTab.Rows(0).Item("CB_Type") = 20 Then
                    '                            If Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) <> "" Then
                    '                                OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document' And Entete=True", OleConnenection)
                    '                                OleRecherDataset = New DataSet
                    '                                OleRecherAdapter.Fill(OleRecherDataset)
                    '                                OleRechDatable = OleRecherDataset.Tables(0)
                    '                                If OleRechDatable.Rows.Count <> 0 Then
                    '                                    If EstNumeric(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone) = True Then
                    '                                        DOCUMENT.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone))
                    '                                        DOCUMENT.Write()
                    '                                    End If
                    '                                Else
                    '                                    If EstNumeric(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), DecimalNomb, DecimalMone) = True Then
                    '                                        DOCUMENT.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), DecimalNomb, DecimalMone))
                    '                                        DOCUMENT.Write()
                    '                                    End If
                    '                                End If
                    '                            End If
                    '                        End If
                    '                        'Valeur
                    '                        If statistTab.Rows(0).Item("CB_Type") = 7 Then
                    '                            If Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) <> "" Then
                    '                                OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document' And Entete=True", OleConnenection)
                    '                                OleRecherDataset = New DataSet
                    '                                OleRecherAdapter.Fill(OleRecherDataset)
                    '                                OleRechDatable = OleRecherDataset.Tables(0)
                    '                                If OleRechDatable.Rows.Count <> 0 Then
                    '                                    If EstNumeric(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone) = True Then
                    '                                        DOCUMENT.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(OleRechDatable.Rows(0).Item("Correspond")), DecimalNomb, DecimalMone))
                    '                                        DOCUMENT.Write()
                    '                                    End If
                    '                                Else
                    '                                    If EstNumeric(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), DecimalNomb, DecimalMone) = True Then
                    '                                        DOCUMENT.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), DecimalNomb, DecimalMone))
                    '                                        DOCUMENT.Write()
                    '                                    End If
                    '                                End If
                    '                            End If
                    '                        End If

                    '                        'Date Court
                    '                        If statistTab.Rows(0).Item("CB_Type") = 3 Then
                    '                            If Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) <> "" Then
                    '                                OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document' And Entete=True", OleConnenection)
                    '                                OleRecherDataset = New DataSet
                    '                                OleRecherAdapter.Fill(OleRecherDataset)
                    '                                OleRechDatable = OleRecherDataset.Tables(0)
                    '                                If OleRechDatable.Rows.Count <> 0 Then
                    '                                    If Verificatdate(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier, Datagridaffiche.Columns(infoListe.Item(0)).Name, Trim(formatdetype)) = True Then
                    '                                        DOCUMENT.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = RenvoieDateValide(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier)
                    '                                        DOCUMENT.Write()
                    '                                    End If
                    '                                Else
                    '                                    If Verificatdate(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), FormatDatefichier, Datagridaffiche.Columns(infoListe.Item(0)).Name, Trim(formatdetype)) = True Then
                    '                                        DOCUMENT.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = RenvoieDateValide(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), FormatDatefichier)
                    '                                        DOCUMENT.Write()
                    '                                    End If
                    '                                End If
                    '                            End If
                    '                        End If
                    '                        'Date Longue
                    '                        If statistTab.Rows(0).Item("CB_Type") = 14 Then
                    '                            If Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value) <> "" Then
                    '                                OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document' And Entete=True", OleConnenection)
                    '                                OleRecherDataset = New DataSet
                    '                                OleRecherAdapter.Fill(OleRecherDataset)
                    '                                OleRechDatable = OleRecherDataset.Tables(0)
                    '                                If OleRechDatable.Rows.Count <> 0 Then
                    '                                    If Verificatdate(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier, Datagridaffiche.Columns(infoListe.Item(0)).Name, Trim(formatdetype)) = True Then
                    '                                        DOCUMENT.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = RenvoieDateValide(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier)
                    '                                        DOCUMENT.Write()
                    '                                    End If
                    '                                Else
                    '                                    If Verificatdate(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), FormatDatefichier, Datagridaffiche.Columns(infoListe.Item(0)).Name, Trim(formatdetype)) = True Then
                    '                                        DOCUMENT.InfoLibre.Item("" & Datagridaffiche.Columns(infoListe.Item(0)).Name & "") = RenvoieDateValide(Trim(Datagridaffiche.Item(infoListe.Item(0), numLigne).Value), FormatDatefichier)
                    '                                        DOCUMENT.Write()
                    '                                    End If
                    '                                End If
                    '                            End If
                    '                        End If
                    '                    End If
                    '                End If
                    '            Else
                    '                'nothing
                    '            End If
                    '        End If
                    '    End If
                    '    'L'info Libre Parametrée par l'utilisateur n'existe pas dans Sage
                    '    infoListe.RemoveAt(0)
                    'End While
                End If
                .WriteDefault()
            Catch ex As Exception
                exceptionTrouve = True
                If typedoc = "3" Then
                    ErreurJrn.WriteLine("Bon de Livraison N° : " & Trim(Documents.DO_Piece) & " Erreur de Modification de L'information Libre , vérifiez la longueur de la chaine avec celle paramétrée/ ou la cohérence des données " & ex.Message)
                Else
                    If typedoc = "14" Then
                        ErreurJrn.WriteLine("Bon de Retour N° : " & Trim(Documents.DO_Piece) & " Erreur de Modification de L'information Libre , vérifiez la longueur de la chaine avec celle paramétrée/ ou la cohérence des données " & ex.Message)
                    Else
                        If typedoc = "23" Then
                            ErreurJrn.WriteLine("Transfert de Dépot à Dépot du  N° : " & Trim(Documents.DO_Piece) & " Erreur de Modification de L'information Libre , vérifiez la longueur de la chaine avec celle paramétrée/ ou la cohérence des données " & ex.Message)
                        End If
                    End If
                End If
            End Try
        End With
    End Sub
    Private Sub MiseAJourEcheance(ByRef Idocuments As IBODocumentVente3, ByRef FormatDatefichier As Object)
        If IsNothing(Idocuments) = False Then
            For Each Echeancedocument As IBODocumentEcheance3 In Idocuments.FactoryDocumentEcheance.List
                With Echeancedocument
                    If Trim(EcheanceModeReglement) <> "" Then
                        If BaseCpta.FactoryReglement.ExistIntitule(Trim(EcheanceModeReglement)) = True Then
                            .Reglement = BaseCpta.FactoryReglement.ReadIntitule(Trim(EcheanceModeReglement))
                        End If
                    End If
                    If Trim(EcheanceDatePied) <> "" Then
                        If Trim(EcheanceDatePied) <> "" Then
                            .DR_Date = RenvoieDateValide(Trim(EcheanceDatePied), FormatDatefichier)
                        End If
                    End If
                    .Write()
                End With
            Next
        End If
    End Sub
    Private Function RenvoieDepotPrincipal() As String
        Dim DossierAdap As OleDbDataAdapter
        Dim DossierDs As DataSet
        Dim DossierTab As DataTable
        DossierAdap = New OleDbDataAdapter("select * from F_DEPOT WHERE DE_Principal=1", OleSocieteConnect)
        DossierDs = New DataSet
        DossierAdap.Fill(DossierDs)
        DossierTab = DossierDs.Tables(0)
        If DossierTab.Rows.Count <> 0 Then
            Return DossierTab.Rows(0).Item("DE_Intitule")
        Else
            Return Nothing
        End If
    End Function
    Private Sub CreationArticleDepotPrincipal(ByRef C_Lignedocument As IBODocumentVenteLigne3, ByRef C_Article As String, ByRef C_DepotPrincipal As String, ByRef C_Serie As String, ByRef C_TypedeFormat As String)
        With C_Lignedocument
            If BaseCial.FactoryDepot.ExistIntitule(C_DepotPrincipal) = True Then
                If BaseCial.FactoryArticle.ReadReference(Trim(C_Article)).FactoryArticleDepot.ExistDepot(BaseCial.FactoryDepot.ReadIntitule(C_DepotPrincipal)) = True Then
                    If BaseCial.FactoryArticle.ReadReference(Trim(C_Article)).FactoryArticleDepot.ReadDepot(BaseCial.FactoryDepot.ReadIntitule(C_DepotPrincipal)).FactoryArticleDepotLot.ExistNoSerie(Trim(C_Serie)) = True Then
                        .SetDefaultLot(BaseCial.FactoryArticle.ReadReference(Trim(C_Article)).FactoryArticleDepot.ReadDepot(BaseCial.FactoryDepot.ReadIntitule(C_DepotPrincipal)).FactoryArticleDepotLot.ReadNoSerie(Trim(C_Serie)), CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)))
                    End If
                End If
            End If
        End With
    End Sub
    Private Sub CreationArticleDepotIntituleLot(ByRef C_Lignedocument As IBODocumentVenteLigne3, ByRef C_Article As String, ByRef C_DepotEntete As String, ByRef C_DepotLigne As String, ByRef C_Serie As String, ByRef C_TypedeFormat As String)
        With C_Lignedocument
            If Trim(LigneIntituleDepot) <> "" Then
                If Trim(C_DepotLigne) <> "" Then
                    If BaseCial.FactoryDepot.ExistIntitule(Trim(C_DepotLigne)) = True Then
                        If BaseCial.FactoryArticle.ReadReference(Trim(C_Article)).FactoryArticleDepot.ExistDepot(BaseCial.FactoryDepot.ReadIntitule(Trim(C_DepotLigne))) = True Then
                            If BaseCial.FactoryArticle.ReadReference(Trim(C_Article)).FactoryArticleDepot.ReadDepot(BaseCial.FactoryDepot.ReadIntitule(Trim(C_DepotLigne))).FactoryArticleDepotLot.ExistNoSerie(Trim(C_Serie)) = True Then
                                .SetDefaultLot(BaseCial.FactoryArticle.ReadReference(Trim(C_Article)).FactoryArticleDepot.ReadDepot(BaseCial.FactoryDepot.ReadIntitule(Trim(C_DepotLigne))).FactoryArticleDepotLot.ReadNoSerie(Trim(C_Serie)), CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)))
                            End If
                        End If
                    End If
                End If
            Else
                If Trim(EnteteIntituleDepot) <> "" Then
                    If Trim(C_DepotEntete) <> "" Then
                        If BaseCial.FactoryDepot.ExistIntitule(Trim(C_DepotEntete)) = True Then
                            If BaseCial.FactoryArticle.ReadReference(Trim(C_Article)).FactoryArticleDepot.ExistDepot(BaseCial.FactoryDepot.ReadIntitule(Trim(C_DepotEntete))) = True Then
                                If BaseCial.FactoryArticle.ReadReference(Trim(C_Article)).FactoryArticleDepot.ReadDepot(BaseCial.FactoryDepot.ReadIntitule(Trim(C_DepotEntete))).FactoryArticleDepotLot.ExistNoSerie(Trim(C_Serie)) = True Then
                                    .SetDefaultLot(BaseCial.FactoryArticle.ReadReference(Trim(C_Article)).FactoryArticleDepot.ReadDepot(BaseCial.FactoryDepot.ReadIntitule(Trim(C_DepotEntete))).FactoryArticleDepotLot.ReadNoSerie(Trim(C_Serie)), CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)))
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End With
    End Sub
    Private Sub CreationArticleDepotIDLot(ByRef C_Lignedocument As IBODocumentVenteLigne3, ByRef I_Article As String, ByRef I_DepotEntete As String, ByRef I_DepotLigne As String, ByRef I_Serie As String, ByRef I_TypedeFormat As String)
        Dim statistAdap As OleDbDataAdapter
        Dim statistDs As DataSet
        Dim statistTab As DataTable
        With C_Lignedocument
            If Trim(IDDepotLigne) <> "" Then
                If IsNumeric(Trim(I_DepotLigne)) = True Then
                    statistAdap = New OleDbDataAdapter("select * from F_DEPOT where DE_No ='" & CInt(Trim(I_DepotLigne)) & "'", OleSocieteConnect)
                    statistDs = New DataSet
                    statistAdap.Fill(statistDs)
                    statistTab = statistDs.Tables(0)
                    If statistTab.Rows.Count <> 0 Then
                        If BaseCial.FactoryArticle.ReadReference(Trim(I_Article)).FactoryArticleDepot.ExistDepot(BaseCial.FactoryDepot.ReadIntitule(Trim(statistTab.Rows(0).Item("DE_Intitule")))) = True Then
                            If BaseCial.FactoryArticle.ReadReference(Trim(I_Article)).FactoryArticleDepot.ReadDepot(BaseCial.FactoryDepot.ReadIntitule(Trim(statistTab.Rows(0).Item("DE_Intitule")))).FactoryArticleDepotLot.ExistNoSerie(Trim(I_Serie)) = True Then
                                .SetDefaultLot(BaseCial.FactoryArticle.ReadReference(Trim(I_Article)).FactoryArticleDepot.ReadDepot(BaseCial.FactoryDepot.ReadIntitule(Trim(statistTab.Rows(0).Item("DE_Intitule")))).FactoryArticleDepotLot.ReadNoSerie(Trim(I_Serie)), CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)))
                            End If
                        End If
                    End If
                End If
            Else
                If Trim(IDDepotEntete) <> "" Then
                    If Trim(I_DepotEntete) <> "" Then
                        statistAdap = New OleDbDataAdapter("select * from F_DEPOT where DE_Intitule ='" & Trim(I_DepotEntete) & "'", OleSocieteConnect)
                        statistDs = New DataSet
                        statistAdap.Fill(statistDs)
                        statistTab = statistDs.Tables(0)
                        If statistTab.Rows.Count <> 0 Then
                            If BaseCial.FactoryArticle.ReadReference(Trim(I_Article)).FactoryArticleDepot.ExistDepot(BaseCial.FactoryDepot.ReadIntitule(Trim(statistTab.Rows(0).Item("DE_Intitule")))) = True Then
                                If BaseCial.FactoryArticle.ReadReference(Trim(I_Article)).FactoryArticleDepot.ReadDepot(BaseCial.FactoryDepot.ReadIntitule(Trim(statistTab.Rows(0).Item("DE_Intitule")))).FactoryArticleDepotLot.ExistNoSerie(Trim(I_Serie)) = True Then
                                    .SetDefaultLot(BaseCial.FactoryArticle.ReadReference(Trim(I_Article)).FactoryArticleDepot.ReadDepot(BaseCial.FactoryDepot.ReadIntitule(Trim(statistTab.Rows(0).Item("DE_Intitule")))).FactoryArticleDepotLot.ReadNoSerie(Trim(I_Serie)), CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)))
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End With
    End Sub
    Public Function RenvoieQteCompose(ByRef OM_LigneDocumentVente As IBODocumentVenteLigne3) As Object
        If IsNothing(OM_LigneDocumentVente.Article) = False Then
            If OM_LigneDocumentVente.Article.AR_Nomencl = NomenclatureType.NomenclatureTypeComposant Or OM_LigneDocumentVente.Article.AR_Nomencl = NomenclatureType.NomenclatureTypeCompose Or OM_LigneDocumentVente.Article.AR_Nomencl = NomenclatureType.NomenclatureTypeFabrication Or OM_LigneDocumentVente.Article.AR_Nomencl = NomenclatureType.NomenclatureTypeLies Then
                Return OM_LigneDocumentVente.DL_Qte
            Else
                Return ""
            End If
        Else
            Return ""
        End If
    End Function
    Public Function RenvoieQteComposant(ByRef ArticleCompos As IBOArticle3, ByRef ArticleComposant As IBOArticle3, ByRef QteCompose As Double) As Double
        Dim QteComposant As Double = Nothing
        If ArticleCompos.AR_Nomencl = NomenclatureType.NomenclatureTypeCompose Or ArticleCompos.AR_Nomencl = NomenclatureType.NomenclatureTypeComposant Or ArticleCompos.AR_Nomencl = NomenclatureType.NomenclatureTypeFabrication Or ArticleCompos.AR_Nomencl = NomenclatureType.NomenclatureTypeLies Then
            For Each OM_Nomenclature As IBOArticleNomenclature3 In ArticleCompos.FactoryArticleNomenclature.List
                If OM_Nomenclature.ArticleComposant.AR_Ref = ArticleComposant.AR_Ref Then
                    If OM_Nomenclature.NO_Type = ComposantType.ComposantTypeVariable Then
                        QteComposant = OM_Nomenclature.NO_Qte * QteCompose
                        Exit For
                    Else
                        QteComposant = OM_Nomenclature.NO_Qte
                        Exit For
                    End If
                End If
            Next
            RenvoieQteComposant = QteComposant
        End If
    End Function
    Public Function RenvoiMontantConditionnement(ByVal Valeur As Object, ByVal Decimale As Integer, ByRef Sepnombre As String, ByVal SepMonetaire As String) As Double
        If Sepnombre = 1 Then 'SepMonetaire
            Valeur = CDbl(Join(Split(Join(Split(Valeur, "."), Trim(SepMonetaire)), ","), Trim(SepMonetaire)))
        Else
            Valeur = CDbl(Join(Split(Valeur, "."), ","))
        End If
        If Decimale = 0 Then
            Valeur = Math.Round(Valeur, MidpointRounding.AwayFromZero)
        Else
            Valeur = Math.Round(Valeur, Decimale)
        End If
        Valeur = Math.Ceiling(Valeur)
        RenvoiMontantConditionnement = Valeur
    End Function
    Private Sub Creation_Ligne_Article(ByRef FormatDatefichier As String, ByRef PieceCommande As String, ByRef PieceArticle As String, ByRef IdentifiantCommande As String, ByRef IdentifiantArticle As String, ByRef formatdetype As String, ByRef ErreurCreationEntete As Boolean)
        Dim fournisseurAdap As OleDbDataAdapter
        Dim fournisseurDs As DataSet
        Dim fournisseurTab As DataTable
        Dim statistAdap As OleDbDataAdapter
        Dim statistDs As DataSet
        Dim statistTab As DataTable
        Dim OleAdaptaterDelete As OleDbDataAdapter = Nothing
        Dim OleDeleteDataset As DataSet = Nothing
        Dim OledatableDelete As DataTable = Nothing
        Dim OleDocAdapter As OleDbDataAdapter
        Dim OleDocDataset As DataSet
        Dim OleDocDatable As DataTable
        Dim OleRecherAdapter As OleDbDataAdapter
        Dim OleRecherDataset As DataSet
        Dim OleRechDatable As DataTable
        Dim DossierAdap As OleDbDataAdapter
        Dim DossierDs As DataSet
        Dim DossierTab As DataTable
        Dim OM_ArticleCompose As IBOArticle3 = Nothing
        Dim EstBLReliquat As Boolean = False
        FormatQte = 0
        Dim FormatMnt As Integer = 0
        Dim PrixUnitDevise As Double = 0
        Dim PrixUnit As Double = 0
        Dim PrixUnitTC As Double = 0
        Dim ValRemise1 As Double = 0
        Dim ValRemise2 As Double = 0
        Dim ValRemise3 As Double = 0
        Dim TypRemise1 As Object = 0
        Dim TypRemise2 As Object = 0
        Dim TypRemise3 As Object = 0
        Dim CbMarq As Integer = 0
        Dim EstLivraisonTotal As Boolean = False
        Dim LigneReliquatInfo As IBODocumentVenteLigne3 = Nothing
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
            If Trim(LignePrixUnitaireDevise) <> "" Then
                If EstNumeric(Trim(LignePrixUnitaireDevise), DecimalNomb, DecimalMone) = True Then
                    If CDbl(RenvoiMontant(Trim(LignePrixUnitaireDevise), FormatMnt, DecimalNomb, DecimalMone)) <> 0 Then
                        BaseCial.FactoryDocumentLigne.AutoSet_PrixLigne = True
                    End If
                End If
            End If
            If Trim(LignePrixUnitaire) <> "" Then
                BaseCial.FactoryDocumentLigne.AutoSet_PrixLigne = False
            End If
            Try
                MiseAJourEcheance(Documents, FormatDatefichier)
            Catch ex As Exception
                ErreurJrn.WriteLine("Erreur de mise à jour de l'échéance - n°Pièce du Fichier : " & Trim(EntetePieceInterne))
            End Try
            LigneDocument = Documents.FactoryDocumentLigne.Create
            With LigneDocument

                If Trim(LigneReferenceArticleTiers) <> "" Then
                    .AC_RefClient = LigneReferenceArticleTiers
                End If
                If Trim(LigneNomRepresentant) <> "" Then
                    If Trim(LignePrenomRepresentant) <> "" Then
                        If BaseCpta.FactoryCollaborateur.ExistNomPrenom(Trim(LigneNomRepresentant), Trim(LignePrenomRepresentant)) = True Then
                            .Collaborateur = BaseCpta.FactoryCollaborateur.ReadNomPrenom(Trim(LigneNomRepresentant), Trim(LignePrenomRepresentant))
                        End If
                    End If
                End If
                If Trim(LignePlanAnalytique) <> "" Then
                    If Trim(LigneCodeAffaire) <> "" Then
                        If BaseCpta.FactoryAnalytique.ExistIntitule(Trim(LignePlanAnalytique)) = True Then
                            PlanAna = BaseCpta.FactoryAnalytique.ReadIntitule(Trim(LignePlanAnalytique))
                            If BaseCpta.FactoryCompteA.ExistNumero(PlanAna, Trim(LigneCodeAffaire)) = True Then
                                .CompteA = BaseCpta.FactoryCompteA.ReadNumero(PlanAna, Trim(LigneCodeAffaire))
                            Else
                                ErreurJrn.WriteLine("< Le code Affaire : " & Trim(LigneCodeAffaire) & " n'existe pas dans les tables paramètres>")
                            End If
                        End If
                    End If
                End If
                If Trim(LigneIntituleDepot) <> "" Then
                    If BaseCial.FactoryDepot.ExistIntitule(Trim(LigneIntituleDepot)) = True Then
                        .Depot = BaseCial.FactoryDepot.ReadIntitule(Trim(LigneIntituleDepot))
                    End If
                End If

                If Trim(IDDepotLigne) <> "" Then
                    If IsNumeric(Trim(IDDepotLigne)) = True Then
                        statistAdap = New OleDbDataAdapter("select * from F_DEPOT where DE_No ='" & CInt(Trim(IDDepotLigne)) & "'", OleSocieteConnect)
                        statistDs = New DataSet
                        statistAdap.Fill(statistDs)
                        statistTab = statistDs.Tables(0)
                        If statistTab.Rows.Count <> 0 Then
                            .Depot = BaseCial.FactoryDepot.ReadIntitule(Trim(statistTab.Rows(0).Item("DE_Intitule")))
                        Else
                            If BaseCial.FactoryDepot.ExistIntitule(Trim(IDDepotLigne)) = True Then
                                .Depot = BaseCial.FactoryDepot.ReadIntitule(Trim(IDDepotLigne))
                            End If
                        End If
                    End If
                End If
                If RbtG3.Checked Then
                    .DL_Design = LigneCodeArticle & " xxxx " & LigneNSerieLot & " DLUO xx/xx/xx " & LigneQuantite & " xxx"
                Else
                    If Trim(LigneDesignationArticle) <> "" Then
                        .DL_Design = LigneDesignationArticle
                    End If
                End If
                If Trim(LigneLibelleComplementaire) <> "" Then
                    .TxtComplementaire = LigneLibelleComplementaire
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
                If Trim(LignePrixdeRevientUnitaire) <> "" Then
                    If EstNumeric(Trim(LignePrixdeRevientUnitaire), DecimalNomb, DecimalMone) = True Then
                        .DL_PrixRU = CDbl(RenvoiMontant(Trim(LignePrixdeRevientUnitaire), FormatMnt, DecimalNomb, DecimalMone))
                    End If
                End If
                If Trim(LigneReference) <> "" Then
                    .DO_Ref = Trim(LigneReference)
                End If
                '.DO_Ref = LigneCodeArticle
                If Trim(LigneValorisé) <> "" Then
                    If Trim(LigneValorisé) = "0" Then
                        .Valorisee = False
                    Else
                        .Valorisee = True
                    End If
                End If
                If Trim(LigneDatedeLivraison) <> "" Then
                    If Trim(LigneDatedeLivraison) <> "" Then
                        .DO_DateLivr = RenvoieDateValide(Trim(LigneDatedeLivraison), FormatDatefichier)
                    End If
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
                If Trim(EnteteCodeTiers) <> "" Then
                    If Trim(LigneArticleCompose) <> "" Then
                        If BaseCpta.FactoryClient.ExistNumero(Trim(EnteteCodeTiers)) = True Then
                            If BaseCial.FactoryArticle.ExistReference(Trim(LigneArticleCompose)) = True Then
                                .ArticleCompose = BaseCial.FactoryArticle.ReadReference(Trim(LigneArticleCompose))
                                OM_ArticleCompose = BaseCial.FactoryArticle.ReadReference(Trim(LigneArticleCompose))
                            Else
                                fournisseurAdap = New OleDbDataAdapter("select * from ARTICLE where Fournisseur ='" & Join(Split(Trim(EnteteCodeTiers), "'"), "''") & "' and Code_Article_Fo ='" & Join(Split(Trim(LigneArticleCompose), "'"), "''") & "'", OleConnenection)
                                fournisseurDs = New DataSet
                                fournisseurAdap.Fill(fournisseurDs)
                                fournisseurTab = fournisseurDs.Tables(0)
                                If fournisseurTab.Rows.Count <> 0 Then
                                    If BaseCial.FactoryArticle.ExistReference(Trim(fournisseurTab.Rows(0).Item("Code_Article_Dis"))) = True Then
                                        .ArticleCompose = BaseCial.FactoryArticle.ReadReference(Trim(fournisseurTab.Rows(0).Item("Code_Article_Dis")))
                                        OM_ArticleCompose = BaseCial.FactoryArticle.ReadReference(Trim(fournisseurTab.Rows(0).Item("Code_Article_Dis")))
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                If Trim(EnteteCodeTiers) <> "" Then
                    If Trim(LigneCodeArticle) <> "" Then
                        If BaseCpta.FactoryClient.ExistNumero(Trim(EnteteCodeTiers)) = True Then
                            If BaseCial.FactoryArticle.ExistReference(Trim(LigneCodeArticle)) = True Then
                                If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then 'je suis 
                                    If BaseCial.FactoryArticle.ReadReference(Trim(LigneCodeArticle)).AR_SuiviStock = SuiviStockType.SuiviStockTypeLot Or BaseCial.FactoryArticle.ReadReference(Trim(LigneCodeArticle)).AR_SuiviStock = SuiviStockType.SuiviStockTypeSerie Then
                                        If Trim(IDDepotLigne) = "" Then
                                            If Trim(IDDepotEntete) = "" Then
                                                If Trim(LigneIntituleDepot) = "" Then
                                                    If Trim(EnteteIntituleDepot) = "" Then
                                                        If Trim(RenvoieDepotPrincipal()) <> "" Then
                                                            If Trim(LigneNSerieLot) <> "" Then
                                                                CreationArticleDepotPrincipal(LigneDocument, Trim(LigneCodeArticle), RenvoieDepotPrincipal(), Trim(LigneNSerieLot), Trim(formatdetype))
                                                            Else
                                                                If RbtG3.Checked = False Then
                                                                    .SetDefaultArticle(BaseCial.FactoryArticle.ReadReference(Trim(LigneCodeArticle)), CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)))
                                                                End If
                                                            End If
                                                        End If
                                                    Else
                                                        If Trim(LigneNSerieLot) <> "" Then
                                                            CreationArticleDepotIntituleLot(LigneDocument, Trim(LigneCodeArticle), Trim(EnteteIntituleDepot), Trim(LigneIntituleDepot), Trim(LigneNSerieLot), Trim(formatdetype))
                                                            CreationArticleDepotIDLot(LigneDocument, Trim(LigneCodeArticle), Trim(IDDepotEntete), Trim(IDDepotLigne), Trim(LigneNSerieLot), Trim(formatdetype))
                                                        Else
                                                            If RbtG3.Checked = False Then
                                                                .SetDefaultArticle(BaseCial.FactoryArticle.ReadReference(Trim(LigneCodeArticle)), CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)))
                                                            End If
                                                        End If
                                                    End If
                                                Else
                                                    If Trim(LigneNSerieLot) <> "" Then
                                                        CreationArticleDepotIntituleLot(LigneDocument, Trim(LigneCodeArticle), Trim(EnteteIntituleDepot), Trim(LigneIntituleDepot), Trim(LigneNSerieLot), Trim(formatdetype))
                                                        CreationArticleDepotIDLot(LigneDocument, Trim(LigneCodeArticle), Trim(IDDepotEntete), Trim(IDDepotLigne), Trim(LigneNSerieLot), Trim(formatdetype))
                                                    Else
                                                        If RbtG3.Checked = False Then
                                                            .SetDefaultArticle(BaseCial.FactoryArticle.ReadReference(Trim(LigneCodeArticle)), CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)))
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                If Trim(LigneNSerieLot) <> "" Then
                                                    CreationArticleDepotIntituleLot(LigneDocument, Trim(LigneCodeArticle), Trim(EnteteIntituleDepot), Trim(LigneIntituleDepot), Trim(LigneNSerieLot), Trim(formatdetype))
                                                    CreationArticleDepotIDLot(LigneDocument, Trim(LigneCodeArticle), Trim(IDDepotEntete), Trim(IDDepotLigne), Trim(LigneNSerieLot), Trim(formatdetype))
                                                Else
                                                    If RbtG3.Checked = False Then
                                                        .SetDefaultArticle(BaseCial.FactoryArticle.ReadReference(Trim(LigneCodeArticle)), CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)))
                                                    End If
                                                End If
                                            End If
                                        Else
                                            If Trim(LigneNSerieLot) <> "" Then
                                                CreationArticleDepotIntituleLot(LigneDocument, Trim(LigneCodeArticle), Trim(EnteteIntituleDepot), Trim(LigneIntituleDepot), Trim(LigneNSerieLot), Trim(formatdetype))
                                                CreationArticleDepotIDLot(LigneDocument, Trim(LigneCodeArticle), Trim(IDDepotEntete), Trim(IDDepotLigne), Trim(LigneNSerieLot), Trim(formatdetype))
                                            Else
                                                .SetDefaultArticle(BaseCial.FactoryArticle.ReadReference(Trim(LigneCodeArticle)), CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)))
                                            End If
                                        End If
                                    Else
                                        If RbtG3.Checked = False Then
                                            .SetDefaultArticle(BaseCial.FactoryArticle.ReadReference(Trim(LigneCodeArticle)), CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)))
                                        End If
                                    End If
                                Else
                                    If IsNothing(OM_ArticleCompose) = False Then
                                        If IsNumeric(OM_QteCompose) = True Then
                                            If IsNumeric(RenvoieQteComposant(OM_ArticleCompose, BaseCial.FactoryArticle.ReadReference(Trim(LigneCodeArticle)), OM_QteCompose)) = True Then
                                                If RbtG3.Checked = False Then
                                                    .SetDefaultArticle(BaseCial.FactoryArticle.ReadReference(Trim(LigneCodeArticle)), RenvoieQteComposant(OM_ArticleCompose, BaseCial.FactoryArticle.ReadReference(Trim(LigneCodeArticle)), OM_QteCompose))
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                'fournisseurAdap = New OleDbDataAdapter("select * from ARTICLE where Fournisseur ='" & Join(Split(Trim(EnteteCodeTiers), "'"), "''") & "' and Code_Article_Fo ='" & Join(Split(Trim(LigneCodeArticle), "'"), "''") & "'", OleConnenection)
                                'fournisseurDs = New DataSet
                                'fournisseurAdap.Fill(fournisseurDs)
                                'fournisseurTab = fournisseurDs.Tables(0)
                                'If fournisseurTab.Rows.Count <> 0 Then
                                '    If BaseCial.FactoryArticle.ExistReference(Trim(fournisseurTab.Rows(0).Item("Code_Article_Dis"))) = True Then
                                '        If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                                '            If BaseCial.FactoryArticle.ReadReference(Trim(fournisseurTab.Rows(0).Item("Code_Article_Dis"))).AR_SuiviStock = SuiviStockType.SuiviStockTypeLot Or BaseCial.FactoryArticle.ReadReference(Trim(fournisseurTab.Rows(0).Item("Code_Article_Dis"))).AR_SuiviStock = SuiviStockType.SuiviStockTypeSerie Then
                                '                If Trim(IDDepotLigne) = "" Then
                                '                    If Trim(IDDepotEntete) = "" Then
                                '                        If Trim(LigneIntituleDepot) = "" Then
                                '                            If Trim(EnteteIntituleDepot) = "" Then
                                '                                If Trim(RenvoieDepotPrincipal()) <> "" Then
                                '                                    If Trim(LigneNSerieLot) <> "" Then
                                '                                        CreationArticleDepotPrincipal(LigneDocument, Trim(fournisseurTab.Rows(0).Item("Code_Article_Dis")), RenvoieDepotPrincipal(), Trim(LigneNSerieLot), Trim(formatdetype))
                                '                                    Else
                                '                                        .SetDefaultArticle(BaseCial.FactoryArticle.ReadReference(Trim(fournisseurTab.Rows(0).Item("Code_Article_Dis"))), CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)))
                                '                                    End If
                                '                                End If
                                '                            Else
                                '                                If Trim(LigneNSerieLot) <> "" Then
                                '                                    CreationArticleDepotIntituleLot(LigneDocument, Trim(fournisseurTab.Rows(0).Item("Code_Article_Dis")), Trim(EnteteIntituleDepot), Trim(LigneIntituleDepot), Trim(LigneNSerieLot), Trim(formatdetype))
                                '                                    CreationArticleDepotIDLot(LigneDocument, Trim(fournisseurTab.Rows(0).Item("Code_Article_Dis")), Trim(IDDepotEntete), Trim(IDDepotLigne), Trim(LigneNSerieLot), Trim(formatdetype))
                                '                                Else
                                '                                    .SetDefaultArticle(BaseCial.FactoryArticle.ReadReference(Trim(fournisseurTab.Rows(0).Item("Code_Article_Dis"))), CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)))
                                '                                End If
                                '                            End If
                                '                        Else
                                '                            If Trim(LigneNSerieLot) <> "" Then
                                '                                CreationArticleDepotIntituleLot(LigneDocument, Trim(fournisseurTab.Rows(0).Item("Code_Article_Dis")), Trim(EnteteIntituleDepot), Trim(LigneIntituleDepot), Trim(LigneNSerieLot), Trim(formatdetype))
                                '                                CreationArticleDepotIDLot(LigneDocument, Trim(fournisseurTab.Rows(0).Item("Code_Article_Dis")), Trim(IDDepotEntete), Trim(IDDepotLigne), Trim(LigneNSerieLot), Trim(formatdetype))
                                '                            Else
                                '                                .SetDefaultArticle(BaseCial.FactoryArticle.ReadReference(Trim(fournisseurTab.Rows(0).Item("Code_Article_Dis"))), CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)))
                                '                            End If
                                '                        End If
                                '                    Else
                                '                        If Trim(LigneNSerieLot) <> "" Then
                                '                            CreationArticleDepotIntituleLot(LigneDocument, Trim(fournisseurTab.Rows(0).Item("Code_Article_Dis")), Trim(EnteteIntituleDepot), Trim(LigneIntituleDepot), Trim(LigneNSerieLot), Trim(formatdetype))
                                '                            CreationArticleDepotIDLot(LigneDocument, Trim(fournisseurTab.Rows(0).Item("Code_Article_Dis")), Trim(IDDepotEntete), Trim(IDDepotLigne), Trim(LigneNSerieLot), Trim(formatdetype))
                                '                        Else
                                '                            .SetDefaultArticle(BaseCial.FactoryArticle.ReadReference(Trim(fournisseurTab.Rows(0).Item("Code_Article_Dis"))), CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)))
                                '                        End If
                                '                    End If
                                '                Else
                                '                    If Trim(LigneNSerieLot) <> "" Then
                                '                        CreationArticleDepotIntituleLot(LigneDocument, Trim(fournisseurTab.Rows(0).Item("Code_Article_Dis")), Trim(EnteteIntituleDepot), Trim(LigneIntituleDepot), Trim(LigneNSerieLot), Trim(formatdetype))
                                '                        CreationArticleDepotIDLot(LigneDocument, Trim(fournisseurTab.Rows(0).Item("Code_Article_Dis")), Trim(IDDepotEntete), Trim(IDDepotLigne), Trim(LigneNSerieLot), Trim(formatdetype))
                                '                    Else
                                '                        .SetDefaultArticle(BaseCial.FactoryArticle.ReadReference(Trim(fournisseurTab.Rows(0).Item("Code_Article_Dis"))), CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)))
                                '                    End If
                                '                End If
                                '            Else
                                '                .SetDefaultArticle(BaseCial.FactoryArticle.ReadReference(Trim(fournisseurTab.Rows(0).Item("Code_Article_Dis"))), CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone)))
                                '            End If
                                '        Else
                                '            If IsNothing(OM_ArticleCompose) = False Then
                                '                If IsNumeric(OM_QteCompose) = True Then
                                '                    If IsNumeric(RenvoieQteComposant(OM_ArticleCompose, BaseCial.FactoryArticle.ReadReference(Trim(fournisseurTab.Rows(0).Item("Code_Article_Dis"))), OM_QteCompose)) = True Then
                                '                        .SetDefaultArticle(BaseCial.FactoryArticle.ReadReference(Trim(fournisseurTab.Rows(0).Item("Code_Article_Dis"))), RenvoieQteComposant(OM_ArticleCompose, BaseCial.FactoryArticle.ReadReference(Trim(fournisseurTab.Rows(0).Item("Code_Article_Dis"))), OM_QteCompose))
                                '                    End If
                                '                End If
                                '            End If
                                '        End If
                                '    End If
                                'End If
                            End If
                        End If
                    End If
                End If
                If RbtG3.Checked = False Then
                    If IsNumeric(RenvoieQteCompose(LigneDocument)) = True Then
                        OM_QteCompose = RenvoieQteCompose(LigneDocument)
                    End If
                    If Trim(LigneEnumereConditionnement) <> "" Then
                        .EU_Enumere = Trim(LigneEnumereConditionnement)
                    End If
                    If Trim(LigneQuantiteConditionne) <> "" Then
                        If EstNumeric(Trim(LigneQuantiteConditionne), DecimalNomb, DecimalMone) = True Then
                            .EU_Qte = CDbl(RenvoiMontant(Trim(LigneQuantiteConditionne), FormatQte, DecimalNomb, DecimalMone))
                        End If
                    End If
                End If
                Dim DateCmde As Date
                Dim ExisteCommande As Boolean = False
                Dim EstPieceCommande As Object = Nothing
                Dim EstDL_NoRef As Object = Nothing
                If ErreurCreationEntete = False Then
                    If Trim(EnteteCodeTiers) <> "" Then
                        If BaseCpta.FactoryClient.ExistNumero(Trim(EnteteCodeTiers)) = True Then
                            If Trim(IdentifiantArticle) <> "" Then
                                If InStr(IdentifiantCommande, ",") <> 0 Then
                                    If Trim(IdentifiantCommande) <> "" Then
                                        If Trim(LigneQuantite) <> "" Then
                                            Dim OleAdaptaterLe As OleDbDataAdapter
                                            Dim OleLeDataset As DataSet
                                            Dim OledatableLe As DataTable
                                            OleAdaptaterLe = New OleDbDataAdapter("select * from COLIMPMOUV WHERE Libelle='" & Trim(IdentifiantArticle) & "' And Fichier='F_DOCLIGNE'", OleConnenection)
                                            OleLeDataset = New DataSet
                                            OleAdaptaterLe.Fill(OleLeDataset)
                                            OledatableLe = OleLeDataset.Tables(0)
                                            If OledatableLe.Rows.Count <> 0 Then
                                                Dim OleAdaptaterCa As OleDbDataAdapter
                                                Dim OleCaDataset As DataSet
                                                Dim OledatableCa As DataTable
                                                OleAdaptaterCa = New OleDbDataAdapter("select * from COLIMPMOUV WHERE Libelle='" & Trim(Strings.Left(IdentifiantCommande, InStr(IdentifiantCommande, ",") - 1)) & "' And Fichier='" & Trim(Strings.Right(IdentifiantCommande, Len(IdentifiantCommande) - InStr(IdentifiantCommande, ","))) & "'", OleConnenection)
                                                OleCaDataset = New DataSet
                                                OleAdaptaterCa.Fill(OleCaDataset)
                                                OledatableCa = OleCaDataset.Tables(0)
                                                If OledatableCa.Rows.Count <> 0 Then
                                                    Try
                                                        If EnteteTyPeDocument = "3" Then
                                                            OleDocAdapter = New OleDbDataAdapter("Select  * From F_DOCENTETE WHERE " & OledatableCa.Rows(0).Item("Champ") & " ='" & Join(Split(Trim(PieceCommande), "'"), "''") & "' And DO_Type=3 And DO_Domaine=0", OleSocieteConnect)
                                                        ElseIf EnteteTyPeDocument = "14" Then
                                                            OleDocAdapter = New OleDbDataAdapter("Select  * From F_DOCENTETE WHERE " & OledatableCa.Rows(0).Item("Champ") & " ='" & Join(Split(Trim(PieceCommande), "'"), "''") & "' And DO_Type=14 And DO_Domaine=1", OleSocieteConnect)
                                                        End If
                                                        OleDocDataset = New DataSet
                                                        OleDocAdapter.Fill(OleDocDataset)
                                                        OleDocDatable = OleDocDataset.Tables(0)
                                                        If OleDocDatable.Rows.Count = 1 Then
                                                            ExisteCommande = True
                                                            DateCmde = OleDocDatable.Rows(0).Item("DO_Date")
                                                            EstPieceCommande = Join(Split(Trim(OleDocDatable.Rows(0).Item("DO_Piece")), "'"), "''")
                                                            If EnteteTyPeDocument = "3" Then
                                                                OleRecherAdapter = New OleDbDataAdapter("Select  * From F_DOCLIGNE WHERE DO_Piece='" & Join(Split(Trim(OleDocDatable.Rows(0).Item("DO_Piece")), "'"), "''") & "' And  " & OledatableLe.Rows(0).Item("Champ") & " ='" & Join(Split(Trim(PieceArticle), "'"), "''") & "' And CT_Num='" & Join(Split(Trim(EnteteCodeTiers), "'"), "''") & "'And DO_Type=3 And DO_Domaine=0", OleSocieteConnect)
                                                            ElseIf EnteteTyPeDocument = "14" Then
                                                                OleRecherAdapter = New OleDbDataAdapter("Select  * From F_DOCLIGNE WHERE DO_Piece='" & Join(Split(Trim(OleDocDatable.Rows(0).Item("DO_Piece")), "'"), "''") & "' And  " & OledatableLe.Rows(0).Item("Champ") & " ='" & Join(Split(Trim(PieceArticle), "'"), "''") & "' And CT_Num='" & Join(Split(Trim(EnteteCodeTiers), "'"), "''") & "'And DO_Type=14 And DO_Domaine=1", OleSocieteConnect)
                                                            End If
                                                            OleRecherDataset = New DataSet
                                                            OleRecherAdapter.Fill(OleRecherDataset)
                                                            OleRechDatable = OleRecherDataset.Tables(0)
                                                            If OleRechDatable.Rows.Count = 1 Then
                                                                Dim OleAdaptaterCa1 As OleDbDataAdapter
                                                                Dim OleCaDataset1 As DataSet
                                                                Dim OledatableCa1 As DataTable
                                                                OleAdaptaterCa1 = New OleDbDataAdapter("Select  * From cbSysLibre WHERE CB_File='F_DOCLIGNE' And CB_Name ='" & Trim(IdentifiantArticle) & "'", OleSocieteConnect)
                                                                OleCaDataset1 = New DataSet
                                                                OleAdaptaterCa1.Fill(OleCaDataset1)
                                                                OledatableCa1 = OleCaDataset1.Tables(0)
                                                                If OledatableCa1.Rows.Count <> 0 Then
                                                                    EstDL_NoRef = OleRechDatable.Rows(0).Item("DL_NoRef")
                                                                    If OleRechDatable.Rows(0).Item("DO_Type") = 0 Then
                                                                        If PieceReliquat <> OleRechDatable.Rows(0).Item("DO_Piece") Then
                                                                            DocumentReliquat = BaseCial.FactoryDocumentVente.ReadPiece(DocumentType.DocumentTypeVenteDevis, OleRechDatable.Rows(0).Item("DO_Piece"))
                                                                            DocumentReliquat.CouldModified()
                                                                        End If
                                                                    Else
                                                                        If OleRechDatable.Rows(0).Item("DO_Type") = 1 Then
                                                                            If PieceReliquat <> OleRechDatable.Rows(0).Item("DO_Piece") Then
                                                                                DocumentReliquat = BaseCial.FactoryDocumentVente.ReadPiece(DocumentType.DocumentTypeVenteCommande, OleRechDatable.Rows(0).Item("DO_Piece"))
                                                                                DocumentReliquat.CouldModified()
                                                                            End If
                                                                        Else
                                                                            If OleRechDatable.Rows(0).Item("DO_Type") = 2 Then
                                                                                If PieceReliquat <> OleRechDatable.Rows(0).Item("DO_Piece") Then
                                                                                    DocumentReliquat = BaseCial.FactoryDocumentVente.ReadPiece(DocumentType.DocumentTypeVentePrepaLivraison, OleRechDatable.Rows(0).Item("DO_Piece"))
                                                                                    DocumentReliquat.CouldModified()
                                                                                End If
                                                                            Else
                                                                                If OleRechDatable.Rows(0).Item("DO_Type") = 3 Then
                                                                                    If PieceReliquat <> OleRechDatable.Rows(0).Item("DO_Piece") Then
                                                                                        DocumentReliquat = BaseCial.FactoryDocumentVente.ReadPiece(DocumentType.DocumentTypeVenteLivraison, OleRechDatable.Rows(0).Item("DO_Piece"))
                                                                                        DocumentReliquat.CouldModified()
                                                                                    End If
                                                                                Else
                                                                                    If OleRechDatable.Rows(0).Item("DO_Type") = 4 Then
                                                                                        If PieceReliquat <> OleRechDatable.Rows(0).Item("DO_Piece") Then
                                                                                            DocumentReliquat = BaseCial.FactoryDocumentVente.ReadPiece(DocumentType.DocumentTypeVenteReprise, OleRechDatable.Rows(0).Item("DO_Piece"))
                                                                                            DocumentReliquat.CouldModified()
                                                                                        End If
                                                                                    Else
                                                                                        If OleRechDatable.Rows(0).Item("DO_Type") = 5 Then
                                                                                            If PieceReliquat <> OleRechDatable.Rows(0).Item("DO_Piece") Then
                                                                                                DocumentReliquat = BaseCial.FactoryDocumentVente.ReadPiece(DocumentType.DocumentTypeVenteAvoir, OleRechDatable.Rows(0).Item("DO_Piece"))
                                                                                                DocumentReliquat.CouldModified()
                                                                                            End If
                                                                                        Else
                                                                                            If OleRechDatable.Rows(0).Item("DO_Type") = 6 Then
                                                                                                If PieceReliquat <> OleRechDatable.Rows(0).Item("DO_Piece") Then
                                                                                                    DocumentReliquat = BaseCial.FactoryDocumentVente.ReadPiece(DocumentType.DocumentTypeVenteFacture, OleRechDatable.Rows(0).Item("DO_Piece"))
                                                                                                    DocumentReliquat.CouldModified()
                                                                                                End If
                                                                                            Else
                                                                                                If OleRechDatable.Rows(0).Item("DO_Type") = 7 Then
                                                                                                    If PieceReliquat <> OleRechDatable.Rows(0).Item("DO_Piece") Then
                                                                                                        DocumentReliquat = BaseCial.FactoryDocumentVente.ReadPiece(DocumentType.DocumentTypeVenteFactureCpta, OleRechDatable.Rows(0).Item("DO_Piece"))
                                                                                                        DocumentReliquat.CouldModified()
                                                                                                    End If
                                                                                                Else
                                                                                                    If OleRechDatable.Rows(0).Item("DO_Type") = 8 Then
                                                                                                        If PieceReliquat <> OleRechDatable.Rows(0).Item("DO_Piece") Then
                                                                                                            DocumentReliquat = BaseCial.FactoryDocumentVente.ReadPiece(DocumentType.DocumentTypeVenteArchive, OleRechDatable.Rows(0).Item("DO_Piece"))
                                                                                                            DocumentReliquat.CouldModified()
                                                                                                        End If
                                                                                                    Else
                                                                                                        If OleRechDatable.Rows(0).Item("DO_Type") = 14 Then
                                                                                                            If PieceReliquat <> OleRechDatable.Rows(0).Item("DO_Piece") Then
                                                                                                                DocumentReliquat = BaseCial.FactoryDocumentAchat.ReadPiece(DocumentType.DocumentTypeAchatReprise, OleRechDatable.Rows(0).Item("DO_Piece"))
                                                                                                                DocumentReliquat.CouldModified()
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
                                                                    PieceReliquat = OleRechDatable.Rows(0).Item("DO_Piece")
                                                                    If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                                                                        If (OleRechDatable.Rows(0).Item("DL_Qte") - CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone))) <= 0 Then
                                                                            Try
                                                                                Try
                                                                                    If IsNothing(DocumentReliquat) = False Then
                                                                                        For Each LigneReliquat In DocumentReliquat.FactoryDocumentLigne.List
                                                                                            With LigneReliquat
                                                                                                If IsNothing(.Article) = False Then
                                                                                                    If Trim(OleRechDatable.Rows(0).Item("AR_Ref")) = Trim(.Article.AR_Ref) And LigneReliquat.InfoLibre.Item("" & IdentifiantArticle & "") = Trim(OleRechDatable.Rows(0).Item("" & IdentifiantArticle & "")) Then
                                                                                                        LigneDocument.DL_Qte = LigneReliquat.DL_Qte
                                                                                                        If Trim(LigneCodeAffaire) <> "" Then
                                                                                                            LigneDocument.CompteA = LigneReliquat.CompteA
                                                                                                        End If
                                                                                                        If Trim(LigneEnumereConditionnement) <> "" Then
                                                                                                            LigneDocument.EU_Enumere = LigneReliquat.EU_Enumere
                                                                                                        End If
                                                                                                        If Trim(LigneNomRepresentant) <> "" Then
                                                                                                            If Trim(LignePrenomRepresentant) <> "" Then
                                                                                                                LigneDocument.Collaborateur = LigneReliquat.Collaborateur
                                                                                                            End If
                                                                                                        End If
                                                                                                        LigneReliquatInfo = LigneReliquat
                                                                                                        PrixUnitTC = .DL_PUTTC
                                                                                                        PrixUnitDevise = .DL_PUDevise
                                                                                                        PrixUnit = .DL_PrixUnitaire
                                                                                                        ValRemise1 = .Remise.Remise(1).REM_Valeur
                                                                                                        If .Remise.Remise(1).REM_Type = RemiseType.RemiseTypeMontant Then
                                                                                                            TypRemise1 = 0
                                                                                                        Else
                                                                                                            If .Remise.Remise(1).REM_Type = RemiseType.RemiseTypePourcent Then
                                                                                                                TypRemise1 = 1
                                                                                                            Else
                                                                                                                If .Remise.Remise(1).REM_Type = RemiseType.RemiseTypeUnite Then
                                                                                                                    TypRemise1 = 2
                                                                                                                End If
                                                                                                            End If
                                                                                                        End If
                                                                                                        ValRemise2 = .Remise.Remise(2).REM_Valeur
                                                                                                        If .Remise.Remise(2).REM_Type = RemiseType.RemiseTypeMontant Then
                                                                                                            TypRemise2 = 0
                                                                                                        Else
                                                                                                            If .Remise.Remise(2).REM_Type = RemiseType.RemiseTypePourcent Then
                                                                                                                TypRemise2 = 1
                                                                                                            Else
                                                                                                                If .Remise.Remise(2).REM_Type = RemiseType.RemiseTypeUnite Then
                                                                                                                    TypRemise2 = 2
                                                                                                                End If
                                                                                                            End If
                                                                                                        End If
                                                                                                        ValRemise3 = .Remise.Remise(3).REM_Valeur
                                                                                                        If .Remise.Remise(3).REM_Type = RemiseType.RemiseTypeMontant Then
                                                                                                            TypRemise3 = 0
                                                                                                        Else
                                                                                                            If .Remise.Remise(3).REM_Type = RemiseType.RemiseTypePourcent Then
                                                                                                                TypRemise3 = 1
                                                                                                            Else
                                                                                                                If .Remise.Remise(3).REM_Type = RemiseType.RemiseTypeUnite Then
                                                                                                                    TypRemise3 = 2
                                                                                                                End If
                                                                                                            End If
                                                                                                        End If
                                                                                                        CbMarq = OleRechDatable.Rows(0).Item("cbMarq")
                                                                                                        EstLivraisonTotal = True
                                                                                                    End If
                                                                                                End If
                                                                                            End With
                                                                                        Next
                                                                                    End If
                                                                                Catch ex As Exception

                                                                                End Try
                                                                            Catch ex As Exception
                                                                                exceptionTrouve = True
                                                                            End Try
                                                                        Else
                                                                            Try
                                                                                If IsNothing(DocumentReliquat) = False Then
                                                                                    For Each LigneReliquat In DocumentReliquat.FactoryDocumentLigne.List
                                                                                        With LigneReliquat
                                                                                            If IsNothing(.Article) = False Then
                                                                                                If Trim(OleRechDatable.Rows(0).Item("AR_Ref")) = Trim(.Article.AR_Ref) And LigneReliquat.InfoLibre.Item("" & IdentifiantArticle & "") = Trim(OleRechDatable.Rows(0).Item("" & IdentifiantArticle & "")) Then
                                                                                                    If Trim(LigneCodeAffaire) <> "" Then
                                                                                                        LigneDocument.CompteA = LigneReliquat.CompteA
                                                                                                    End If
                                                                                                    If Trim(LigneEnumereConditionnement) <> "" Then
                                                                                                        LigneDocument.EU_Enumere = LigneReliquat.EU_Enumere
                                                                                                    End If
                                                                                                    If Trim(LigneNomRepresentant) <> "" Then
                                                                                                        If Trim(LignePrenomRepresentant) <> "" Then
                                                                                                            LigneDocument.Collaborateur = LigneReliquat.Collaborateur
                                                                                                        End If
                                                                                                    End If
                                                                                                    LigneReliquatInfo = LigneReliquat
                                                                                                    PrixUnitTC = .DL_PUTTC
                                                                                                    PrixUnitDevise = .DL_PUDevise
                                                                                                    PrixUnit = .DL_PrixUnitaire
                                                                                                    ValRemise1 = .Remise.Remise(1).REM_Valeur
                                                                                                    If .Remise.Remise(1).REM_Type = RemiseType.RemiseTypeMontant Then
                                                                                                        TypRemise1 = 0
                                                                                                    Else
                                                                                                        If .Remise.Remise(1).REM_Type = RemiseType.RemiseTypePourcent Then
                                                                                                            TypRemise1 = 1
                                                                                                        Else
                                                                                                            If .Remise.Remise(1).REM_Type = RemiseType.RemiseTypeUnite Then
                                                                                                                TypRemise1 = 2
                                                                                                            End If
                                                                                                        End If
                                                                                                    End If
                                                                                                    ValRemise2 = .Remise.Remise(2).REM_Valeur
                                                                                                    If .Remise.Remise(2).REM_Type = RemiseType.RemiseTypeMontant Then
                                                                                                        TypRemise2 = 0
                                                                                                    Else
                                                                                                        If .Remise.Remise(2).REM_Type = RemiseType.RemiseTypePourcent Then
                                                                                                            TypRemise2 = 1
                                                                                                        Else
                                                                                                            If .Remise.Remise(2).REM_Type = RemiseType.RemiseTypeUnite Then
                                                                                                                TypRemise2 = 2
                                                                                                            End If
                                                                                                        End If
                                                                                                    End If
                                                                                                    ValRemise3 = .Remise.Remise(3).REM_Valeur
                                                                                                    If .Remise.Remise(3).REM_Type = RemiseType.RemiseTypeMontant Then
                                                                                                        TypRemise3 = 0
                                                                                                    Else
                                                                                                        If .Remise.Remise(3).REM_Type = RemiseType.RemiseTypePourcent Then
                                                                                                            TypRemise3 = 1
                                                                                                        Else
                                                                                                            If .Remise.Remise(3).REM_Type = RemiseType.RemiseTypeUnite Then
                                                                                                                TypRemise3 = 2
                                                                                                            End If
                                                                                                        End If
                                                                                                    End If
                                                                                                End If
                                                                                            End If
                                                                                        End With
                                                                                    Next
                                                                                End If
                                                                            Catch ex As Exception
                                                                                exceptionTrouve = True
                                                                            End Try
                                                                        End If
                                                                    Else
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    Catch ex As Exception
                                                        exceptionTrouve = False
                                                    End Try
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                If Trim(LignePrixUnitaire) <> "" Then
                    If Trim(LigneTypePrixUnitaire) <> "" Then
                        If Trim(LigneTypePrixUnitaire) = "1" Then
                            If EstNumeric(Trim(LignePrixUnitaire), DecimalNomb, DecimalMone) = True Then
                                .TTC = True
                                .DL_PUTTC = CDbl(RenvoiMontant(Trim(LignePrixUnitaire), FormatMnt, DecimalNomb, DecimalMone))
                            End If
                        Else
                            If EstNumeric(Trim(LignePrixUnitaire), DecimalNomb, DecimalMone) = True Then
                                .TTC = False
                                .DL_PrixUnitaire = CDbl(RenvoiMontant(Trim(LignePrixUnitaire), FormatMnt, DecimalNomb, DecimalMone))
                            End If
                        End If
                    Else
                        If IsNothing(.Article) = False Then
                            If .Article.AR_PrixTTC = False Then
                                If EstNumeric(Trim(LignePrixUnitaire), DecimalNomb, DecimalMone) = True Then
                                    .TTC = False
                                    .DL_PrixUnitaire = CDbl(RenvoiMontant(Trim(LignePrixUnitaire), FormatMnt, DecimalNomb, DecimalMone))
                                End If
                            Else
                                If EstNumeric(Trim(LignePrixUnitaire), DecimalNomb, DecimalMone) = True Then
                                    .TTC = True
                                    .DL_PUTTC = CDbl(RenvoiMontant(Trim(LignePrixUnitaire), FormatMnt, DecimalNomb, DecimalMone))
                                End If
                            End If
                        End If
                    End If
                Else
                    If ExisteCommande = True Then
                        If Trim(LignePrixUnitaireDevise) <> "" Then
                            If Trim(EnteteIntituleDevise) <> "" Then
                                If Trim(EnteteIntituleDevise) = "" Then
                                    If .DO_Type = DocumentType.DocumentTypeVenteLivraison Or .DO_Type = DocumentType.DocumentTypeVenteFacture Then
                                        .DL_PUTTC = PrixUnitTC
                                        .DL_PrixUnitaire = PrixUnit
                                    End If
                                End If
                            Else
                                If .DO_Type = DocumentType.DocumentTypeVenteLivraison Or .DO_Type = DocumentType.DocumentTypeVenteFacture Then
                                    .DL_PUTTC = PrixUnitTC
                                    .DL_PrixUnitaire = PrixUnit
                                End If
                            End If
                        End If
                    Else
                        If Trim(LignePrixUnitaireDevise) <> "" Then
                            If Trim(EnteteIntituleDevise) <> "" Then
                                If Trim(EnteteIntituleDevise) = "" Then
                                    OleDocAdapter = New OleDbDataAdapter("Select  * From F_COMPTET WHERE CT_Num ='" & Join(Split(LigneDocument.DocumentVente.Client.CT_Num, "'"), "''") & "'", OleSocieteConnect)
                                    OleDocDataset = New DataSet
                                    OleDocAdapter.Fill(OleDocDataset)
                                    OleDocDatable = OleDocDataset.Tables(0)
                                    If OleDocDatable.Rows.Count <> 0 Then
                                        If IsNothing(.Article) = False Then
                                            OleRecherAdapter = New OleDbDataAdapter("Select  * From F_ARTCLIENT WHERE AR_Ref='" & Join(Split(.Article.AR_Ref, "'"), "''") & "' And CT_Num='" & Join(Split(LigneDocument.DocumentVente.Client.CT_Num, "'"), "''") & "'", OleSocieteConnect)
                                            OleRecherDataset = New DataSet
                                            OleRecherAdapter.Fill(OleRecherDataset)
                                            OleRechDatable = OleRecherDataset.Tables(0)
                                            If OleRechDatable.Rows.Count <> 0 Then
                                                If Convert.IsDBNull(OleRechDatable.Rows(0).Item("AC_PrixTTC")) = False Then
                                                    If OleRechDatable.Rows(0).Item("AC_PrixTTC") = 1 Then
                                                        .TTC = True
                                                        .DL_PUTTC = CDbl(OleRechDatable.Rows(0).Item("AC_PrixVen"))
                                                    Else
                                                        .TTC = False
                                                        .DL_PrixUnitaire = CDbl(OleRechDatable.Rows(0).Item("AC_PrixVen"))
                                                    End If
                                                Else
                                                    If Convert.IsDBNull(OleRechDatable.Rows(0).Item("AC_PrixVen")) = False Then
                                                        .TTC = False
                                                        .DL_PrixUnitaire = CDbl(OleRechDatable.Rows(0).Item("AC_PrixVen"))
                                                    End If
                                                End If
                                            Else
                                                If IsNothing(Documents.CategorieTarif) = False Then
                                                    OleDocAdapter = New OleDbDataAdapter("Select  * From P_CATTARIF WHERE CT_Intitule ='" & Join(Split(Documents.CategorieTarif.CT_Intitule, "'"), "''") & "'", OleSocieteConnect)
                                                    OleDocDataset = New DataSet
                                                    OleDocAdapter.Fill(OleDocDataset)
                                                    OleDocDatable = OleDocDataset.Tables(0)
                                                    If OleDocDatable.Rows.Count <> 0 Then
                                                        If IsNothing(.Article) = False Then
                                                            If Trim(LignePrixUnitaireDevise) <> "" Then
                                                                If Trim(EnteteIntituleDevise) <> "" Then
                                                                    If Trim(EnteteIntituleDevise) = "" Then
                                                                        OleRecherAdapter = New OleDbDataAdapter("Select  * From F_ARTCLIENT WHERE (AR_Ref='" & Join(Split(.Article.AR_Ref, "'"), "''") & "' And  AC_Categorie =" & OleDocDatable.Rows(0).Item("cbIndice") & ") And (CT_Num IS NULL OR CT_Num='')", OleSocieteConnect)
                                                                        OleRecherDataset = New DataSet
                                                                        OleRecherAdapter.Fill(OleRecherDataset)
                                                                        OleRechDatable = OleRecherDataset.Tables(0)
                                                                        If OleRechDatable.Rows.Count <> 0 Then
                                                                            If Convert.IsDBNull(OleRechDatable.Rows(0).Item("AC_PrixTTC")) = False Then
                                                                                If OleRechDatable.Rows(0).Item("AC_PrixTTC") = 1 Then
                                                                                    .TTC = True
                                                                                    .DL_PUTTC = CDbl(OleRechDatable.Rows(0).Item("AC_PrixVen"))
                                                                                Else
                                                                                    .TTC = False
                                                                                    .DL_PrixUnitaire = CDbl(OleRechDatable.Rows(0).Item("AC_PrixVen"))
                                                                                End If
                                                                            Else
                                                                                If Convert.IsDBNull(OleRechDatable.Rows(0).Item("AC_PrixVen")) = False Then
                                                                                    .TTC = False
                                                                                    .DL_PrixUnitaire = CDbl(OleRechDatable.Rows(0).Item("AC_PrixVen"))
                                                                                End If
                                                                            End If
                                                                        Else
                                                                            .TTC = False
                                                                            .DL_PrixUnitaire = .Article.AR_PrixVen
                                                                        End If
                                                                    End If
                                                                Else
                                                                    OleRecherAdapter = New OleDbDataAdapter("Select  * From F_ARTCLIENT WHERE (AR_Ref='" & Join(Split(.Article.AR_Ref, "'"), "''") & "' And  AC_Categorie =" & OleDocDatable.Rows(0).Item("cbIndice") & ") And (CT_Num IS NULL OR CT_Num='')", OleSocieteConnect)
                                                                    OleRecherDataset = New DataSet
                                                                    OleRecherAdapter.Fill(OleRecherDataset)
                                                                    OleRechDatable = OleRecherDataset.Tables(0)
                                                                    If OleRechDatable.Rows.Count <> 0 Then
                                                                        If Convert.IsDBNull(OleRechDatable.Rows(0).Item("AC_PrixTTC")) = False Then
                                                                            If OleRechDatable.Rows(0).Item("AC_PrixTTC") = 1 Then
                                                                                .TTC = True
                                                                                .DL_PUTTC = CDbl(OleRechDatable.Rows(0).Item("AC_PrixVen"))
                                                                            Else
                                                                                .TTC = False
                                                                                .DL_PrixUnitaire = CDbl(OleRechDatable.Rows(0).Item("AC_PrixVen"))
                                                                            End If
                                                                        Else
                                                                            If Convert.IsDBNull(OleRechDatable.Rows(0).Item("AC_PrixVen")) = False Then
                                                                                .TTC = False
                                                                                .DL_PrixUnitaire = CDbl(OleRechDatable.Rows(0).Item("AC_PrixVen"))
                                                                            End If
                                                                        End If
                                                                    Else
                                                                        If .Article.AR_PrixTTC = False Then
                                                                            .TTC = False
                                                                            .DL_PrixUnitaire = .Article.AR_PrixVen
                                                                        Else
                                                                            .TTC = True
                                                                            .DL_PUTTC = .Article.AR_PrixVen
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
                            Else
                                If BaseCpta.FactoryDevise.ExistIntitule(Trim(DeviseTiers)) = False Then
                                    OleDocAdapter = New OleDbDataAdapter("Select  * From F_COMPTET WHERE CT_Num ='" & Join(Split(LigneDocument.DocumentVente.Client.CT_Num, "'"), "''") & "'", OleSocieteConnect)
                                    OleDocDataset = New DataSet
                                    OleDocAdapter.Fill(OleDocDataset)
                                    OleDocDatable = OleDocDataset.Tables(0)
                                    If OleDocDatable.Rows.Count <> 0 Then
                                        If IsNothing(.Article) = False Then
                                            OleRecherAdapter = New OleDbDataAdapter("Select  * From F_ARTCLIENT WHERE AR_Ref='" & Join(Split(.Article.AR_Ref, "'"), "''") & "' And CT_Num='" & Join(Split(LigneDocument.DocumentVente.Client.CT_Num, "'"), "''") & "'", OleSocieteConnect)
                                            OleRecherDataset = New DataSet
                                            OleRecherAdapter.Fill(OleRecherDataset)
                                            OleRechDatable = OleRecherDataset.Tables(0)
                                            If OleRechDatable.Rows.Count <> 0 Then
                                                If Convert.IsDBNull(OleRechDatable.Rows(0).Item("AC_PrixTTC")) = False Then
                                                    If OleRechDatable.Rows(0).Item("AC_PrixTTC") = 1 Then
                                                        .TTC = True
                                                        .DL_PUTTC = CDbl(OleRechDatable.Rows(0).Item("AC_PrixVen"))

                                                    Else
                                                        .TTC = False
                                                        .DL_PrixUnitaire = CDbl(OleRechDatable.Rows(0).Item("AC_PrixVen"))
                                                    End If
                                                Else
                                                    If Convert.IsDBNull(OleRechDatable.Rows(0).Item("AC_PrixVen")) = False Then
                                                        .TTC = False
                                                        .DL_PrixUnitaire = CDbl(OleRechDatable.Rows(0).Item("AC_PrixVen"))
                                                    End If
                                                End If
                                            Else
                                                If IsNothing(Documents.CategorieTarif) = False Then
                                                    OleDocAdapter = New OleDbDataAdapter("Select  * From P_CATTARIF WHERE CT_Intitule ='" & Join(Split(Documents.CategorieTarif.CT_Intitule, "'"), "''") & "'", OleSocieteConnect)
                                                    OleDocDataset = New DataSet
                                                    OleDocAdapter.Fill(OleDocDataset)
                                                    OleDocDatable = OleDocDataset.Tables(0)
                                                    If OleDocDatable.Rows.Count <> 0 Then
                                                        If IsNothing(.Article) = False Then
                                                            If Trim(LignePrixUnitaireDevise) <> "" Then
                                                                If Trim(EnteteIntituleDevise) <> "" Then
                                                                    If Trim(EnteteIntituleDevise) = "" Then
                                                                        OleRecherAdapter = New OleDbDataAdapter("Select  * From F_ARTCLIENT WHERE (AR_Ref='" & Join(Split(.Article.AR_Ref, "'"), "''") & "' And  AC_Categorie =" & OleDocDatable.Rows(0).Item("cbIndice") & ") And (CT_Num IS NULL OR CT_Num='')", OleSocieteConnect)
                                                                        OleRecherDataset = New DataSet
                                                                        OleRecherAdapter.Fill(OleRecherDataset)
                                                                        OleRechDatable = OleRecherDataset.Tables(0)
                                                                        If OleRechDatable.Rows.Count <> 0 Then
                                                                            If Convert.IsDBNull(OleRechDatable.Rows(0).Item("AC_PrixTTC")) = False Then
                                                                                If CDbl(OleRechDatable.Rows(0).Item("AC_Prixven")) = 0 Then
                                                                                    If .Article.AR_PrixTTC = False Then
                                                                                        .TTC = False
                                                                                        .DL_PrixUnitaire = .Article.AR_PrixVen
                                                                                    Else
                                                                                        .TTC = True
                                                                                        .DL_PUTTC = .Article.AR_PrixVen
                                                                                    End If
                                                                                Else
                                                                                    If OleRechDatable.Rows(0).Item("AC_PrixTTC") = 1 Then
                                                                                        .TTC = True
                                                                                        .DL_PUTTC = CDbl(OleRechDatable.Rows(0).Item("AC_PrixVen"))
                                                                                    Else
                                                                                        .TTC = False
                                                                                        .DL_PrixUnitaire = CDbl(OleRechDatable.Rows(0).Item("AC_PrixVen"))
                                                                                    End If
                                                                                End If
                                                                            Else
                                                                                If Convert.IsDBNull(OleRechDatable.Rows(0).Item("AC_PrixVen")) = False Then
                                                                                    If CDbl(OleRechDatable.Rows(0).Item("AC_Prixven")) = 0 Then
                                                                                        If .Article.AR_PrixTTC = False Then
                                                                                            .TTC = False
                                                                                            .DL_PrixUnitaire = .Article.AR_PrixVen
                                                                                        Else
                                                                                            .TTC = True
                                                                                            .DL_PUTTC = .Article.AR_PrixVen
                                                                                        End If
                                                                                    Else
                                                                                        .TTC = False
                                                                                        .DL_PrixUnitaire = CDbl(OleRechDatable.Rows(0).Item("AC_PrixVen"))
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        Else
                                                                            .TTC = False
                                                                            .DL_PrixUnitaire = .Article.AR_PrixVen
                                                                        End If
                                                                    End If
                                                                Else
                                                                    OleRecherAdapter = New OleDbDataAdapter("Select  * From F_ARTCLIENT WHERE (AR_Ref='" & Join(Split(.Article.AR_Ref, "'"), "''") & "' And  AC_Categorie =" & OleDocDatable.Rows(0).Item("cbIndice") & ") And (CT_Num IS NULL OR CT_Num='')", OleSocieteConnect)
                                                                    OleRecherDataset = New DataSet
                                                                    OleRecherAdapter.Fill(OleRecherDataset)
                                                                    OleRechDatable = OleRecherDataset.Tables(0)
                                                                    If OleRechDatable.Rows.Count <> 0 Then
                                                                        If Convert.IsDBNull(OleRechDatable.Rows(0).Item("AC_PrixTTC")) = False Then
                                                                            If CDbl(OleRechDatable.Rows(0).Item("AC_Prixven")) = 0 Then
                                                                                If .Article.AR_PrixTTC = False Then
                                                                                    .TTC = False
                                                                                    .DL_PrixUnitaire = .Article.AR_PrixVen
                                                                                Else
                                                                                    .TTC = True
                                                                                    .DL_PUTTC = .Article.AR_PrixVen
                                                                                End If
                                                                            Else
                                                                                If OleRechDatable.Rows(0).Item("AC_PrixTTC") = 1 Then
                                                                                    .TTC = True
                                                                                    .DL_PUTTC = CDbl(OleRechDatable.Rows(0).Item("AC_PrixVen"))
                                                                                Else
                                                                                    .TTC = False
                                                                                    .DL_PrixUnitaire = CDbl(OleRechDatable.Rows(0).Item("AC_PrixVen"))
                                                                                End If
                                                                            End If
                                                                        Else
                                                                            If Convert.IsDBNull(OleRechDatable.Rows(0).Item("AC_PrixVen")) = False Then
                                                                                If CDbl(OleRechDatable.Rows(0).Item("AC_Prixven")) = 0 Then
                                                                                    If .Article.AR_PrixTTC = False Then
                                                                                        .TTC = False
                                                                                        .DL_PrixUnitaire = .Article.AR_PrixVen
                                                                                    Else
                                                                                        .TTC = True
                                                                                        .DL_PUTTC = .Article.AR_PrixVen
                                                                                    End If
                                                                                Else
                                                                                    .TTC = False
                                                                                    .DL_PrixUnitaire = CDbl(OleRechDatable.Rows(0).Item("AC_PrixVen"))
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    Else
                                                                        If .Article.AR_PrixTTC = False Then
                                                                            .TTC = False
                                                                            '------------------HermannMVT--------------------------
                                                                            If .DL_PrixUnitaire = .Article.AR_PrixVen Then
                                                                                .DL_PrixUnitaire = .Article.AR_PrixVen
                                                                            End If
                                                                        Else
                                                                            .TTC = True
                                                                            .DL_PUTTC = .Article.AR_PrixVen
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
                            End If
                        End If
                    End If
                End If
                If Trim(LignePrixUnitaireDevise) <> "" Then
                    If Trim(EnteteIntituleDevise) <> "" Then
                        If Trim(EnteteIntituleDevise) <> "" Then
                            If EstNumeric(Trim(LignePrixUnitaireDevise), DecimalNomb, DecimalMone) = True Then
                                .DL_PUDevise = CDbl(RenvoiMontant(Trim(LignePrixUnitaireDevise), FormatMnt, DecimalNomb, DecimalMone))
                            End If
                        Else
                            If EstNumeric(Trim(LignePrixUnitaireDevise), DecimalNomb, DecimalMone) = True Then
                                .DL_PrixUnitaire = CDbl(RenvoiMontant(Trim(LignePrixUnitaireDevise), FormatMnt, DecimalNomb, DecimalMone))
                            End If
                        End If
                    Else
                        If EstNumeric(Trim(LignePrixUnitaireDevise), DecimalNomb, DecimalMone) = True Then
                            .DL_PUDevise = CDbl(RenvoiMontant(Trim(LignePrixUnitaireDevise), FormatMnt, DecimalNomb, DecimalMone))
                        End If
                    End If
                Else
                    If ExisteCommande = True Then
                        If Trim(LignePrixUnitaire) <> "" Then
                            If Trim(EnteteIntituleDevise) <> "" Then
                                If Trim(EnteteIntituleDevise) <> "" Then
                                    If .DO_Type = DocumentType.DocumentTypeVenteLivraison Or .DO_Type = DocumentType.DocumentTypeVenteFacture Then
                                        .DL_PUDevise = PrixUnitDevise
                                    End If
                                End If
                            Else
                                If BaseCpta.FactoryDevise.ExistIntitule(Trim(DeviseTiers)) = True Then
                                    If .DO_Type = DocumentType.DocumentTypeVenteLivraison Or .DO_Type = DocumentType.DocumentTypeVenteFacture Then
                                        .DL_PUDevise = PrixUnitDevise
                                    End If
                                End If
                            End If
                        End If
                    Else
                        OleDocAdapter = New OleDbDataAdapter("Select  * From F_COMPTET WHERE CT_Num ='" & Join(Split(LigneDocument.DocumentVente.Client.CT_Num, "'"), "''") & "'", OleSocieteConnect)
                        OleDocDataset = New DataSet
                        OleDocAdapter.Fill(OleDocDataset)
                        OleDocDatable = OleDocDataset.Tables(0)
                        If OleDocDatable.Rows.Count <> 0 Then
                            If IsNothing(.Article) = False Then
                                If Trim(LignePrixUnitaire) <> "" Then
                                    If Trim(EnteteIntituleDevise) <> "" Then
                                        If Trim(EnteteIntituleDevise) <> "" Then
                                            OleRecherAdapter = New OleDbDataAdapter("Select  * From F_ARTCLIENT WHERE AR_Ref='" & Join(Split(.Article.AR_Ref, "'"), "''") & "' And CT_Num='" & Join(Split(LigneDocument.DocumentVente.Client.CT_Num, "'"), "''") & "'", OleSocieteConnect)
                                            OleRecherDataset = New DataSet
                                            OleRecherAdapter.Fill(OleRecherDataset)
                                            OleRechDatable = OleRecherDataset.Tables(0)
                                            If OleRechDatable.Rows.Count <> 0 Then
                                                If Convert.IsDBNull(OleRechDatable.Rows(0).Item("AC_PrixDev")) = False Then
                                                    .DL_PUDevise = CDbl(OleRechDatable.Rows(0).Item("AC_PrixDev"))
                                                End If
                                            Else
                                                If IsNothing(Documents.CategorieTarif) = False Then
                                                    OleDocAdapter = New OleDbDataAdapter("Select  * From P_CATTARIF WHERE CT_Intitule ='" & Join(Split(Documents.CategorieTarif.CT_Intitule, "'"), "''") & "'", OleSocieteConnect)
                                                    OleDocDataset = New DataSet
                                                    OleDocAdapter.Fill(OleDocDataset)
                                                    OleDocDatable = OleDocDataset.Tables(0)
                                                    If OleDocDatable.Rows.Count <> 0 Then
                                                        If IsNothing(.Article) = False Then
                                                            If Trim(LignePrixUnitaire) <> "" Then
                                                                If Trim(EnteteIntituleDevise) <> "" Then
                                                                    If Trim(EnteteIntituleDevise) <> "" Then
                                                                        OleRecherAdapter = New OleDbDataAdapter("Select  * From F_ARTCLIENT WHERE (AR_Ref='" & Join(Split(.Article.AR_Ref, "'"), "''") & "' And  AC_Categorie =" & OleDocDatable.Rows(0).Item("cbIndice") & ") And (CT_Num IS NULL OR CT_Num='')", OleSocieteConnect)
                                                                        OleRecherDataset = New DataSet
                                                                        OleRecherAdapter.Fill(OleRecherDataset)
                                                                        OleRechDatable = OleRecherDataset.Tables(0)
                                                                        If OleRechDatable.Rows.Count <> 0 Then
                                                                            If Convert.IsDBNull(OleRechDatable.Rows(0).Item("AC_PrixDev")) = False Then
                                                                                .DL_PUDevise = CDbl(OleRechDatable.Rows(0).Item("AC_PrixDev"))
                                                                            End If
                                                                        Else
                                                                            .DL_PUDevise = .Article.AR_PrixVen
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        If BaseCpta.FactoryDevise.ExistIntitule(Trim(DeviseTiers)) = True Then
                                            OleRecherAdapter = New OleDbDataAdapter("Select  * From F_ARTCLIENT WHERE AR_Ref='" & Join(Split(.Article.AR_Ref, "'"), "''") & "' And CT_Num='" & Join(Split(LigneDocument.DocumentVente.Client.CT_Num, "'"), "''") & "'", OleSocieteConnect)
                                            OleRecherDataset = New DataSet
                                            OleRecherAdapter.Fill(OleRecherDataset)
                                            OleRechDatable = OleRecherDataset.Tables(0)
                                            If OleRechDatable.Rows.Count <> 0 Then
                                                If Convert.IsDBNull(OleRechDatable.Rows(0).Item("AC_PrixDev")) = False Then
                                                    .DL_PUDevise = CDbl(OleRechDatable.Rows(0).Item("AC_PrixDev"))
                                                End If
                                            Else
                                                If IsNothing(Documents.CategorieTarif) = False Then
                                                    OleDocAdapter = New OleDbDataAdapter("Select  * From P_CATTARIF WHERE CT_Intitule ='" & Join(Split(Documents.CategorieTarif.CT_Intitule, "'"), "''") & "'", OleSocieteConnect)
                                                    OleDocDataset = New DataSet
                                                    OleDocAdapter.Fill(OleDocDataset)
                                                    OleDocDatable = OleDocDataset.Tables(0)
                                                    If OleDocDatable.Rows.Count <> 0 Then
                                                        If IsNothing(.Article) = False Then
                                                            If Trim(LignePrixUnitaire) <> "" Then
                                                                If Trim(EnteteIntituleDevise) <> "" Then
                                                                    OleRecherAdapter = New OleDbDataAdapter("Select  * From F_ARTCLIENT WHERE (AR_Ref='" & Join(Split(.Article.AR_Ref, "'"), "''") & "' And  AC_Categorie =" & OleDocDatable.Rows(0).Item("cbIndice") & ") And (CT_Num IS NULL OR CT_Num='')", OleSocieteConnect)
                                                                    OleRecherDataset = New DataSet
                                                                    OleRecherAdapter.Fill(OleRecherDataset)
                                                                    OleRechDatable = OleRecherDataset.Tables(0)
                                                                    If OleRechDatable.Rows.Count <> 0 Then
                                                                        If Convert.IsDBNull(OleRechDatable.Rows(0).Item("AC_PrixDev")) = False Then
                                                                            .DL_PUDevise = CDbl(OleRechDatable.Rows(0).Item("AC_PrixDev"))
                                                                        End If
                                                                    Else
                                                                        .DL_PUDevise = .Article.AR_PrixVen
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
                        End If
                    End If
                End If
                Dim ExisteRemise As Boolean = False
                Dim RemiseDescriptif As Boolean = False
                If Trim(LigneTauxRemise1) <> "" Then
                    RemiseDescriptif = True
                    If Trim(LigneTauxRemise1) <> "" Then
                        If EstNumeric(Trim(LigneTauxRemise1), DecimalNomb, DecimalMone) = True Then
                            If CDbl(RenvoiTaux(Trim(LigneTauxRemise1), DecimalNomb, DecimalMone)) <> 0 Or CDbl(RenvoiTaux(Trim(LigneTauxRemise1), DecimalNomb, DecimalMone)) = 0 Then
                                ExisteRemise = True
                            End If
                        End If
                    End If
                End If
                If Trim(LigneTauxRemise2) <> "" Then
                    RemiseDescriptif = True
                    If Trim(LigneTauxRemise2) <> "" Then
                        If EstNumeric(Trim(LigneTauxRemise2), DecimalNomb, DecimalMone) = True Then
                            If CDbl(RenvoiTaux(Trim(LigneTauxRemise2), DecimalNomb, DecimalMone)) <> 0 Or CDbl(RenvoiTaux(Trim(LigneTauxRemise2), DecimalNomb, DecimalMone)) = 0 Then
                                ExisteRemise = True
                            End If
                        End If
                    End If
                End If
                If Trim(LigneTauxRemise3) <> "" Then
                    RemiseDescriptif = True
                    If Trim(LigneTauxRemise3) <> "" Then
                        If EstNumeric(Trim(LigneTauxRemise3), DecimalNomb, DecimalMone) = True Then
                            If CDbl(RenvoiTaux(Trim(LigneTauxRemise3), DecimalNomb, DecimalMone)) <> 0 Or CDbl(RenvoiTaux(Trim(LigneTauxRemise3), DecimalNomb, DecimalMone)) = 0 Then
                                ExisteRemise = True
                            End If
                        End If
                    End If
                End If
                If Trim(LigneTypeRemise1) <> "" Then
                    If Trim(LigneTauxRemise1) <> "" Then
                        If Trim(LigneTypeRemise1) = "1" Then
                            If EstNumeric(Trim(LigneTauxRemise1), DecimalNomb, DecimalMone) = True Then
                                .Remise.Remise(1).REM_Type = RemiseType.RemiseTypePourcent
                                .Remise.Remise(1).REM_Valeur = CDbl(RenvoiTaux(Trim(LigneTauxRemise1), DecimalNomb, DecimalMone))
                            End If
                        Else
                            If Trim(LigneTypeRemise1) = "2" Then
                                If EstNumeric(Trim(LigneTauxRemise1), DecimalNomb, DecimalMone) = True Then
                                    .Remise.Remise(1).REM_Type = RemiseType.RemiseTypeUnite
                                    .Remise.Remise(1).REM_Valeur = CDbl(RenvoiTaux(Trim(LigneTauxRemise1), DecimalNomb, DecimalMone))
                                End If
                            Else
                                If Trim(LigneTypeRemise1) = "0" Then
                                    If EstNumeric(Trim(LigneTauxRemise1), DecimalNomb, DecimalMone) = True Then
                                        .Remise.Remise(1).REM_Type = RemiseType.RemiseTypeMontant
                                        .Remise.Remise(1).REM_Valeur = CDbl(RenvoiTaux(Trim(LigneTauxRemise1), DecimalNomb, DecimalMone))
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If RemiseDescriptif = False And ExisteCommande = True Then
                            If TypRemise1 = 0 Then
                                .Remise.Remise(1).REM_Valeur = ValRemise1
                                .Remise.Remise(1).REM_Type = RemiseType.RemiseTypeMontant
                            Else
                                If TypRemise1 = 1 Then
                                    .Remise.Remise(1).REM_Valeur = ValRemise1
                                    .Remise.Remise(1).REM_Type = RemiseType.RemiseTypePourcent
                                Else
                                    If TypRemise1 = 2 Then
                                        .Remise.Remise(1).REM_Valeur = ValRemise1
                                        .Remise.Remise(1).REM_Type = RemiseType.RemiseTypeUnite
                                    Else
                                        .Remise.Remise(1).REM_Valeur = ValRemise1
                                        .Remise.Remise(1).REM_Type = RemiseType.RemiseTypePourcent
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else
                    If Trim(LigneTauxRemise1) <> "" Then
                        If EstNumeric(Trim(LigneTauxRemise1), DecimalNomb, DecimalMone) = True Then
                            .Remise.Remise(1).REM_Type = RemiseType.RemiseTypePourcent
                            .Remise.Remise(1).REM_Valeur = CDbl(RenvoiTaux(Trim(LigneTauxRemise1), DecimalNomb, DecimalMone))
                        End If
                    Else
                        If RemiseDescriptif = False And ExisteCommande = True Then
                            If TypRemise1 = 0 Then
                                .Remise.Remise(1).REM_Valeur = ValRemise1
                                .Remise.Remise(1).REM_Type = RemiseType.RemiseTypeMontant
                            Else
                                If TypRemise1 = 1 Then
                                    .Remise.Remise(1).REM_Valeur = ValRemise1
                                    .Remise.Remise(1).REM_Type = RemiseType.RemiseTypePourcent
                                Else
                                    If TypRemise1 = 2 Then
                                        .Remise.Remise(1).REM_Valeur = ValRemise1
                                        .Remise.Remise(1).REM_Type = RemiseType.RemiseTypeUnite
                                    Else
                                        .Remise.Remise(1).REM_Valeur = ValRemise1
                                        .Remise.Remise(1).REM_Type = RemiseType.RemiseTypePourcent
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                If Trim(LigneTypeRemise2) <> "" Then
                    If Trim(LigneTauxRemise2) <> "" Then
                        If Trim(LigneTypeRemise2) = "1" Then
                            If EstNumeric(Trim(LigneTauxRemise2), DecimalNomb, DecimalMone) = True Then
                                .Remise.Remise(2).REM_Type = RemiseType.RemiseTypePourcent
                                .Remise.Remise(2).REM_Valeur = CDbl(RenvoiTaux(Trim(LigneTauxRemise2), DecimalNomb, DecimalMone))
                            End If
                            If Trim(LigneTypeRemise2) = "2" Then
                                If EstNumeric(Trim(LigneTauxRemise2), DecimalNomb, DecimalMone) = True Then
                                    .Remise.Remise(2).REM_Type = RemiseType.RemiseTypeUnite
                                    .Remise.Remise(2).REM_Valeur = CDbl(RenvoiTaux(Trim(LigneTauxRemise2), DecimalNomb, DecimalMone))
                                End If
                            Else
                                If Trim(LigneTypeRemise2) = "0" Then
                                    If EstNumeric(Trim(LigneTauxRemise2), DecimalNomb, DecimalMone) = True Then
                                        .Remise.Remise(2).REM_Type = RemiseType.RemiseTypeMontant
                                        .Remise.Remise(2).REM_Valeur = CDbl(RenvoiTaux(Trim(LigneTauxRemise2), DecimalNomb, DecimalMone))
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If RemiseDescriptif = False And ExisteCommande = True Then
                            If TypRemise2 = 0 Then
                                .Remise.Remise(2).REM_Valeur = ValRemise2
                                .Remise.Remise(2).REM_Type = RemiseType.RemiseTypeMontant
                            Else
                                If TypRemise1 = 1 Then
                                    .Remise.Remise(2).REM_Valeur = ValRemise2
                                    .Remise.Remise(2).REM_Type = RemiseType.RemiseTypePourcent
                                Else
                                    If TypRemise2 = 2 Then
                                        .Remise.Remise(2).REM_Valeur = ValRemise2
                                        .Remise.Remise(2).REM_Type = RemiseType.RemiseTypeUnite
                                    Else
                                        .Remise.Remise(2).REM_Valeur = ValRemise2
                                        .Remise.Remise(2).REM_Type = RemiseType.RemiseTypePourcent
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else
                    If Trim(LigneTauxRemise2) <> "" Then
                        If EstNumeric(Trim(LigneTauxRemise2), DecimalNomb, DecimalMone) = True Then
                            .Remise.Remise(2).REM_Type = RemiseType.RemiseTypePourcent
                            .Remise.Remise(2).REM_Valeur = CDbl(RenvoiTaux(Trim(LigneTauxRemise2), DecimalNomb, DecimalMone))
                        End If
                    Else
                        If RemiseDescriptif = False And ExisteCommande = True Then
                            If TypRemise2 = 0 Then
                                .Remise.Remise(2).REM_Valeur = ValRemise2
                                .Remise.Remise(2).REM_Type = RemiseType.RemiseTypeMontant
                            Else
                                If TypRemise2 = 1 Then
                                    .Remise.Remise(2).REM_Valeur = ValRemise2
                                    .Remise.Remise(2).REM_Type = RemiseType.RemiseTypePourcent
                                Else
                                    If TypRemise2 = 2 Then
                                        .Remise.Remise(2).REM_Valeur = ValRemise2
                                        .Remise.Remise(2).REM_Type = RemiseType.RemiseTypeUnite
                                    Else
                                        .Remise.Remise(2).REM_Valeur = ValRemise2
                                        .Remise.Remise(2).REM_Type = RemiseType.RemiseTypePourcent
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

                If Trim(LigneTypeRemise3) <> "" Then
                    If Trim(LigneTauxRemise3) <> "" Then
                        If Trim(LigneTypeRemise3) = "1" Then
                            If EstNumeric(Trim(LigneTauxRemise3), DecimalNomb, DecimalMone) = True Then
                                .Remise.Remise(3).REM_Type = RemiseType.RemiseTypePourcent
                                .Remise.Remise(3).REM_Valeur = CDbl(RenvoiTaux(Trim(LigneTauxRemise3), DecimalNomb, DecimalMone))
                            End If
                        Else
                            If Trim(LigneTypeRemise3) = "2" Then
                                If EstNumeric(Trim(LigneTauxRemise3), DecimalNomb, DecimalMone) = True Then
                                    .Remise.Remise(3).REM_Type = RemiseType.RemiseTypeUnite
                                    .Remise.Remise(3).REM_Valeur = CDbl(RenvoiTaux(Trim(LigneTauxRemise3), DecimalNomb, DecimalMone))
                                End If
                            Else
                                If Trim(LigneTypeRemise3) = "0" Then
                                    If EstNumeric(Trim(LigneTauxRemise3), DecimalNomb, DecimalMone) = True Then
                                        .Remise.Remise(3).REM_Type = RemiseType.RemiseTypeMontant
                                        .Remise.Remise(3).REM_Valeur = CDbl(RenvoiTaux(Trim(LigneTauxRemise3), DecimalNomb, DecimalMone))
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If RemiseDescriptif = False And ExisteCommande = True Then
                            If TypRemise3 = 0 Then
                                .Remise.Remise(3).REM_Valeur = ValRemise3
                                .Remise.Remise(3).REM_Type = RemiseType.RemiseTypeMontant
                            Else
                                If TypRemise3 = 1 Then
                                    .Remise.Remise(3).REM_Valeur = ValRemise3
                                    .Remise.Remise(3).REM_Type = RemiseType.RemiseTypePourcent
                                Else
                                    If TypRemise3 = 2 Then
                                        .Remise.Remise(3).REM_Valeur = ValRemise3
                                        .Remise.Remise(3).REM_Type = RemiseType.RemiseTypeUnite
                                    Else
                                        .Remise.Remise(3).REM_Valeur = ValRemise3
                                        .Remise.Remise(3).REM_Type = RemiseType.RemiseTypePourcent
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else
                    If Trim(LigneTauxRemise3) <> "" Then
                        If EstNumeric(Trim(LigneTauxRemise3), DecimalNomb, DecimalMone) = True Then
                            .Remise.Remise(3).REM_Type = RemiseType.RemiseTypePourcent
                            .Remise.Remise(3).REM_Valeur = CDbl(RenvoiTaux(Trim(LigneTauxRemise3), DecimalNomb, DecimalMone))
                        End If
                    Else
                        If RemiseDescriptif = False And ExisteCommande = True Then
                            If TypRemise3 = 0 Then
                                .Remise.Remise(3).REM_Valeur = ValRemise3
                                .Remise.Remise(3).REM_Type = RemiseType.RemiseTypeMontant
                            Else
                                If TypRemise3 = 1 Then
                                    .Remise.Remise(3).REM_Valeur = ValRemise3
                                    .Remise.Remise(3).REM_Type = RemiseType.RemiseTypePourcent
                                Else
                                    If TypRemise3 = 2 Then
                                        .Remise.Remise(3).REM_Valeur = ValRemise3
                                        .Remise.Remise(3).REM_Type = RemiseType.RemiseTypeUnite
                                    Else
                                        .Remise.Remise(3).REM_Valeur = ValRemise3
                                        .Remise.Remise(3).REM_Type = RemiseType.RemiseTypePourcent
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                If ExisteRemise = False And RemiseDescriptif = False And ExisteCommande = False Then
                    .SetDefaultRemise()
                End If
                .Write()
                '-------------------------Application du Prix de revient à null Hermann ---------------
                If Trim(LignePrixdeRevientUnitaire) <> "" Then
                    If EstNumeric(Trim(LignePrixdeRevientUnitaire), DecimalNomb, DecimalMone) = False Then
                        .DL_PrixRU = 0
                        .Write()
                    End If
                End If
                '----------------------------------------------------------------------------
                EstBLReliquat = True
                If EstBLReliquat = True Then
                    If ErreurCreationEntete = False Then
                        If Trim(EnteteCodeTiers) <> "" Then
                            If BaseCpta.FactoryClient.ExistNumero(Trim(EnteteCodeTiers)) = True Then
                                If Trim(IdentifiantArticle) <> "" Then
                                    If InStr(IdentifiantCommande, ",") <> 0 Then
                                        If Trim(IdentifiantCommande) <> "" Then
                                            If Trim(LigneQuantite) <> "" Then
                                                Dim OleAdaptaterLe As OleDbDataAdapter
                                                'a etudier
                                                Dim OleLeDataset As DataSet
                                                Dim OledatableLe As DataTable
                                                OleAdaptaterLe = New OleDbDataAdapter("select * from COLIMPMOUV WHERE Libelle='" & Trim(IdentifiantArticle) & "' And Fichier='F_DOCLIGNE'", OleConnenection)
                                                OleLeDataset = New DataSet
                                                OleAdaptaterLe.Fill(OleLeDataset)
                                                OledatableLe = OleLeDataset.Tables(0)
                                                If OledatableLe.Rows.Count <> 0 Then
                                                    Dim OleAdaptaterCa As OleDbDataAdapter
                                                    Dim OleCaDataset As DataSet
                                                    Dim OledatableCa As DataTable
                                                    OleAdaptaterCa = New OleDbDataAdapter("select * from COLIMPMOUV WHERE Libelle='" & Trim(Strings.Left(IdentifiantCommande, InStr(IdentifiantCommande, ",") - 1)) & "' And Fichier='" & Trim(Strings.Right(IdentifiantCommande, Len(IdentifiantCommande) - InStr(IdentifiantCommande, ","))) & "'", OleConnenection)
                                                    OleCaDataset = New DataSet
                                                    OleAdaptaterCa.Fill(OleCaDataset)
                                                    OledatableCa = OleCaDataset.Tables(0)
                                                    If OledatableCa.Rows.Count <> 0 Then
                                                        Try
                                                            If EnteteTyPeDocument = "3" Then
                                                                OleDocAdapter = New OleDbDataAdapter("Select  * From F_DOCENTETE WHERE " & OledatableCa.Rows(0).Item("Champ") & " ='" & Join(Split(Trim(PieceCommande), "'"), "''") & "' And DO_Type=3 And DO_Domaine=0", OleSocieteConnect)
                                                            ElseIf EnteteTyPeDocument = "14" Then
                                                                OleDocAdapter = New OleDbDataAdapter("Select  * From F_DOCENTETE WHERE " & OledatableCa.Rows(0).Item("Champ") & " ='" & Join(Split(Trim(PieceCommande), "'"), "''") & "' And DO_Type=14 And DO_Domaine=1", OleSocieteConnect)
                                                            End If
                                                            OleDocDataset = New DataSet
                                                            OleDocAdapter.Fill(OleDocDataset)
                                                            OleDocDatable = OleDocDataset.Tables(0)
                                                            If OleDocDatable.Rows.Count = 1 Then
                                                                If EnteteTyPeDocument = "3" Then
                                                                    OleRecherAdapter = New OleDbDataAdapter("Select  * From F_DOCLIGNE WHERE DO_Piece='" & Join(Split(Trim(OleDocDatable.Rows(0).Item("DO_Piece")), "'"), "''") & "' And  " & OledatableLe.Rows(0).Item("Champ") & " ='" & Join(Split(Trim(PieceArticle), "'"), "''") & "' And CT_Num='" & Join(Split(Trim(EnteteCodeTiers), "'"), "''") & "' And DO_Type=3 And DO_Domaine=0", OleSocieteConnect)
                                                                ElseIf EnteteTyPeDocument = "14" Then
                                                                    OleRecherAdapter = New OleDbDataAdapter("Select  * From F_DOCLIGNE WHERE DO_Piece='" & Join(Split(Trim(OleDocDatable.Rows(0).Item("DO_Piece")), "'"), "''") & "' And  " & OledatableLe.Rows(0).Item("Champ") & " ='" & Join(Split(Trim(PieceArticle), "'"), "''") & "' And CT_Num='" & Join(Split(Trim(EnteteCodeTiers), "'"), "''") & "' And DO_Type=14 And DO_Domaine=1", OleSocieteConnect)
                                                                End If
                                                                OleRecherDataset = New DataSet
                                                                OleRecherAdapter.Fill(OleRecherDataset)
                                                                OleRechDatable = OleRecherDataset.Tables(0)
                                                                If OleRechDatable.Rows.Count = 1 Then
                                                                    Dim OleAdaptaterCa1 As OleDbDataAdapter
                                                                    Dim OleCaDataset1 As DataSet
                                                                    Dim OledatableCa1 As DataTable
                                                                    OleAdaptaterCa1 = New OleDbDataAdapter("Select  * From cbSysLibre WHERE CB_File='F_DOCLIGNE' And CB_Name ='" & Trim(IdentifiantArticle) & "'", OleSocieteConnect)
                                                                    OleCaDataset1 = New DataSet
                                                                    OleAdaptaterCa1.Fill(OleCaDataset1)
                                                                    OledatableCa1 = OleCaDataset1.Tables(0)
                                                                    If OledatableCa1.Rows.Count <> 0 Then
                                                                        PieceReliquat = OleRechDatable.Rows(0).Item("DO_Piece")
                                                                        If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                                                                            If (OleRechDatable.Rows(0).Item("DL_Qte") - CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone))) <= 0 Then
                                                                                Try
                                                                                    Try
                                                                                        If IsNothing(DocumentReliquat) = False Then
                                                                                            For Each LigneReliquat In DocumentReliquat.FactoryDocumentLigne.List
                                                                                                With LigneReliquat
                                                                                                    If IsNothing(.Article) = False Then
                                                                                                        If Trim(OleRechDatable.Rows(0).Item("AR_Ref")) = Trim(.Article.AR_Ref) And LigneReliquat.InfoLibre.Item("" & IdentifiantArticle & "") = Trim(OleRechDatable.Rows(0).Item("" & IdentifiantArticle & "")) Then
                                                                                                            LigneReliquatInfo = LigneReliquat
                                                                                                            CbMarq = OleRechDatable.Rows(0).Item("cbMarq")
                                                                                                            .Remove()
                                                                                                            EstLivraisonTotal = True
                                                                                                        End If
                                                                                                    End If
                                                                                                End With
                                                                                            Next
                                                                                        End If
                                                                                    Catch ex As Exception

                                                                                    End Try
                                                                                    ErreurJrn.WriteLine("N°Article Reliquat  :" & OleRechDatable.Rows(0).Item("AR_Ref") & " Supprimé ! N°Pièce :" & OleRechDatable.Rows(0).Item("DO_Piece") & " Le reliquat est nul")
                                                                                    If IsNothing(DocumentReliquat) = False Then
                                                                                        If DocumentReliquat.FactoryDocumentLigne.List.Count = 0 Then
                                                                                            Try
                                                                                                DocumentReliquat.Read()
                                                                                                DocumentReliquat.Remove()
                                                                                                ErreurJrn.WriteLine("Document Reliquat N°Pièce Sage :" & OleRechDatable.Rows(0).Item("DO_Piece") & " supprimé !")
                                                                                            Catch ex As Exception
                                                                                                ErreurJrn.WriteLine("Document Reliquat N°Pièce Sage :" & OleRechDatable.Rows(0).Item("DO_Piece") & " Erreur de suppression :" & ex.Message)
                                                                                            End Try
                                                                                        End If
                                                                                    End If
                                                                                Catch ex As Exception
                                                                                    exceptionTrouve = True
                                                                                    ErreurJrn.WriteLine("Erreur de Suppression de l'Article Reliquat  :" & OleRechDatable.Rows(0).Item("AR_Ref") & " Supprimé ! N°Pièce :" & OleRechDatable.Rows(0).Item("DO_Piece") & " Le reliquat est nul. Erreur Système :" & ex.Message)
                                                                                End Try
                                                                            Else
                                                                                Try
                                                                                    If IsNothing(DocumentReliquat) = False Then
                                                                                        For Each LigneReliquat In DocumentReliquat.FactoryDocumentLigne.List
                                                                                            With LigneReliquat
                                                                                                If IsNothing(.Article) = False Then
                                                                                                    If Trim(OleRechDatable.Rows(0).Item("AR_Ref")) = Trim(.Article.AR_Ref) And LigneReliquat.InfoLibre.Item("" & IdentifiantArticle & "") = Trim(OleRechDatable.Rows(0).Item("" & IdentifiantArticle & "")) Then
                                                                                                        LigneReliquatInfo = LigneReliquat
                                                                                                        .DL_QtePL = CDbl(Join(Split((OleRechDatable.Rows(0).Item("DL_Qte") - CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone))), ","), "."))
                                                                                                        .DL_Qte = CDbl(Join(Split((OleRechDatable.Rows(0).Item("DL_Qte") - CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone))), ","), "."))
                                                                                                        '.DL_QteBL = CDbl(Join(Split((OleRechDatable.Rows(0).Item("DL_Qte") - CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone))), ","), "."))
                                                                                                        If LigneReliquat.Article.FactoryArticleCond.List.Count = 0 Then
                                                                                                            .EU_Qte = CDbl(Join(Split((OleRechDatable.Rows(0).Item("DL_Qte") - CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone))), ","), "."))
                                                                                                        Else
                                                                                                            For Each ArtCond As IBOArticleCond3 In LigneReliquat.Article.FactoryArticleCond.List
                                                                                                                If ArtCond.EC_Enumere = .EU_Enumere Then
                                                                                                                    If ArtCond.EC_Quantite <> 0 Then
                                                                                                                        .EU_Qte = CDbl(Join(Split((OleRechDatable.Rows(0).Item("DL_Qte") - CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone))), ","), ".")) / ArtCond.EC_Quantite
                                                                                                                    Else
                                                                                                                        .EU_Qte = CDbl(Join(Split((OleRechDatable.Rows(0).Item("DL_Qte") - CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone))), ","), "."))
                                                                                                                    End If
                                                                                                                End If
                                                                                                            Next
                                                                                                        End If
                                                                                                        .Write()
                                                                                                    End If
                                                                                                End If
                                                                                            End With
                                                                                        Next
                                                                                    End If
                                                                                    ErreurJrn.WriteLine("N°Article Reliquat  :" & OleRechDatable.Rows(0).Item("AR_Ref") & " N°Pièce :" & OleRechDatable.Rows(0).Item("DO_Piece") & " mise à jour ! Le reliquat est :" & (OleRechDatable.Rows(0).Item("DL_Qte") - CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone))))
                                                                                Catch ex As Exception
                                                                                    exceptionTrouve = True
                                                                                    ErreurJrn.WriteLine("Erreur de mise à jour de l'Article Reliquat :" & OleRechDatable.Rows(0).Item("AR_Ref") & " N°Pièce :" & OleRechDatable.Rows(0).Item("DO_Piece") & " mise à jour ! Le reliquat est :" & (OleRechDatable.Rows(0).Item("DL_Qte") - CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone))) & " Erreur Système :" & ex.Message)
                                                                                End Try
                                                                            End If
                                                                        Else
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        Catch ex As Exception
                                                            exceptionTrouve = False
                                                            ErreurJrn.WriteLine("< Erreur de Recherche de l'Article Reliquat : " & Trim(PieceArticle) & " - N°Pièce document du reliquat : " & Trim(PieceCommande))
                                                        End Try
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
                If Trim(LigneDesignationArticle) <> "" Then
                    .DL_Design = LigneDesignationArticle
                    .Write()
                End If
                If RbtG3.Checked = False Then
                    If Trim(LigneQuantite) <> "" Then
                        If EstNumeric(Trim(LigneQuantite), DecimalNomb, DecimalMone) = True Then
                            .DL_Qte = CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone))
                            .DL_QteBL = CDbl(RenvoiMontant(Trim(LigneQuantite), FormatQte, DecimalNomb, DecimalMone))
                            .Write()
                            LigneQuantiteConditionne = LigneQuantite
                        End If
                    End If
                    If Trim(LigneQuantiteConditionne) <> "" Then
                        If IsNothing(LigneDocument.Article) = False Then
                            For Each ArticleIBI As IBOArticleCond3 In LigneDocument.Article.FactoryArticleCond.List
                                Dim OleAdaptaterschemaFourssAR = New OleDbDataAdapter("SELECT EC_Enumere, EC_Quantite FROM F_ENUMCOND WHERE  EC_Enumere = '" & ArticleIBI.EC_Enumere & "'", OleSocieteConnect)
                                Dim OleSchemaDatasetFourssAR = New DataSet
                                Dim OledatableSchemaFourssAR As DataTable
                                OleAdaptaterschemaFourssAR.Fill(OleSchemaDatasetFourssAR)
                                OledatableSchemaFourssAR = OleSchemaDatasetFourssAR.Tables(0)
                                If OledatableSchemaFourssAR.Rows.Count <> 0 Then 'If ArticleIBI.EC_Enumere = LigneDocument.EU_Enumere Then
                                    If ArticleIBI.EC_Quantite <> 0 Then
                                        .EU_Qte = CDbl(RenvoiMontantConditionnement(.DL_Qte / ArticleIBI.EC_Quantite, FormatQte, DecimalNomb, DecimalMone))
                                        .Write()
                                    End If
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                End If
                
                If ExisteCommande = True Then
                    If Trim(EstPieceCommande) <> "" Then
                        ListeReliquat.Add(Trim(EstPieceCommande) & ControlChars.Tab & EstDL_NoRef & ControlChars.Tab & .Document.DO_Piece & ControlChars.Tab & PieceArticle & ControlChars.Tab & CbMarq & ControlChars.Tab & EstLivraisonTotal & ControlChars.Tab & DateCmde)
                    End If
                End If
                If IsNothing(LigneDocument.Article) = False Then
                    ErreurJrn.WriteLine("Code Article : " & Trim(LigneDocument.Article.AR_Ref) & " Créé Pour la pièce du fichier N° :" & Trim(EntetePieceInterne))
                Else
                    ErreurJrn.WriteLine("Code Article : " & Trim("Vide") & " Créé Pour la pièce du fichier N° :" & Trim(EntetePieceInterne))
                End If
                Try
                    If IsNothing(LigneReliquatInfo) = False Then
                        For i As Integer = 1 To LigneReliquatInfo.InfoLibre.Count
                            LigneDocument.InfoLibre.Item(i) = LigneReliquatInfo.InfoLibre(i)
                        Next i
                        LigneDocument.Write()
                        LigneReliquatInfo = Nothing
                    End If
                Catch ex As Exception

                End Try
                Try
                    statistAdap = New OleDbDataAdapter("select * from cbSysLibre where CB_File='F_DOCLIGNE' and CB_Name ='" & IdentifiantArticle & "'", OleSocieteConnect)
                    statistDs = New DataSet
                    statistAdap.Fill(statistDs)
                    statistTab = statistDs.Tables(0)
                    If statistTab.Rows.Count <> 0 Then
                        'Texte
                        If statistTab.Rows(0).Item("CB_Type") = 9 Then
                            LigneDocument.InfoLibre.Item("" & IdentifiantArticle & "") = Trim(PieceArticle)
                            LigneDocument.Write()
                        End If
                        'Table
                        If statistTab.Rows(0).Item("CB_Type") = 22 Then
                            LigneDocument.InfoLibre.Item("" & IdentifiantArticle & "") = Trim(PieceArticle)
                            LigneDocument.Write()
                        End If
                        'Montant
                        If statistTab.Rows(0).Item("CB_Type") = 20 Then
                            If EstNumeric(Trim(PieceArticle), DecimalNomb, DecimalMone) = True Then
                                LigneDocument.InfoLibre.Item("" & IdentifiantArticle & "") = CDbl(RenvoiTaux(Trim(PieceArticle), DecimalNomb, DecimalMone))
                                LigneDocument.Write()
                            End If
                        End If
                        'Valeur
                        If statistTab.Rows(0).Item("CB_Type") = 7 Then
                            If EstNumeric(Trim(PieceArticle), DecimalNomb, DecimalMone) = True Then
                                LigneDocument.InfoLibre.Item("" & IdentifiantArticle & "") = CDbl(RenvoiTaux(Trim(PieceArticle), DecimalNomb, DecimalMone))
                                LigneDocument.Write()
                            End If
                        End If
                    End If
                Catch ex As Exception

                End Try
                'Traitement des Infos Libres
                Try
                    If -1 > 0 Then 'infoLigne.Count
                        While infoLigne.Count <> 0
                            OleAdaptaterDelete = New OleDbDataAdapter("select * From COLIMPMOUV where Libelle='" & Trim(Datagridaffiche.Columns(infoLigne.Item(0)).Name) & "' And InfoLigne=True", OleConnenection)
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
                                                    If statistTab.Rows(0).Item("CB_Len") >= Strings.Len(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value)) Then
                                                        LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value)
                                                        LigneDocument.Write()
                                                    End If
                                                End If
                                                'Table
                                                If statistTab.Rows(0).Item("CB_Type") = 22 Then
                                                    LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value)
                                                    LigneDocument.Write()
                                                End If
                                                'Montant
                                                If statistTab.Rows(0).Item("CB_Type") = 20 Then
                                                    If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                                        If EstNumeric(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), DecimalNomb, DecimalMone) = True Then
                                                            LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = CDbl(RenvoiTaux(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), DecimalNomb, DecimalMone))
                                                            LigneDocument.Write()
                                                        End If
                                                    End If
                                                End If
                                                'Valeur
                                                If statistTab.Rows(0).Item("CB_Type") = 7 Then
                                                    If Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value) <> "" Then
                                                        OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document' And Ligne=True", OleConnenection)
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
                                                        OleRecherAdapter = New OleDbDataAdapter("select * from TRANSCODAGEIMPORT where Concerne='" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "' and Valeurlue ='" & Join(Split(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), "'"), "''") & "' And Menu='Importation' And IDDossier=" & TraitementID & "  And Categorie='Document' And Ligne=True", OleConnenection)
                                                        OleRecherDataset = New DataSet
                                                        OleRecherAdapter.Fill(OleRecherDataset)
                                                        OleRechDatable = OleRecherDataset.Tables(0)
                                                        If OleRechDatable.Rows.Count <> 0 Then
                                                            LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = RenvoieDateValide(Trim(OleRechDatable.Rows(0).Item("Correspond")), FormatDatefichier)
                                                            LigneDocument.Write()
                                                        Else
                                                            LigneDocument.InfoLibre.Item("" & Datagridaffiche.Columns(infoLigne.Item(0)).Name & "") = RenvoieDateValide(Trim(Datagridaffiche.Item(infoLigne.Item(0), numLigne).Value), FormatDatefichier)
                                                            LigneDocument.Write()
                                                        End If
                                                    End If
                                                End If
                                                'Date Longue
                                                If statistTab.Rows(0).Item("CB_Type") = 14 Then
                                                    If Trim(infoLigne.Item(0)) <> "" Then
                                                        LigneDocument.InfoLibre.Item("" & infoLigne.Item(0) & "") = RenvoieDateValide(Trim(infoLigne.Item(0)), FormatDatefichier)
                                                        LigneDocument.Write()
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
        End Try
        Try
            '-----------------------------------------------------Lien BL--------------------------BC
            If EntetePieceInterne <> "" And NLignePieceCommande <> "" Then
                Try
                    Dim OleComDeletAna As OleDbCommand
                    Dim DeleteEcriture As String
                    Documents.Read()
                    DeleteEcriture = "SET ARITHABORT ON" 'Optimisation de la requete Sql le 04/02/2015 part Hermann
                    OleComDeletAna = New OleDbCommand(DeleteEcriture)
                    OleComDeletAna.Connection = OleSocieteConnect
                    OleComDeletAna.ExecuteNonQuery()
                    Dim dd = Documents.DO_Piece
                    If EnteteTyPeDocument = "3" Then
                        DeleteEcriture = "UPDATE  F_DOCLIGNE SET DL_PieceBC='" & Trim(Documents.DO_Piece) & "', DL_NoRef=" & Trim(NLignePieceCommande) & ",DL_DateBC=CONVERT(DATETIME, " & RenvoieDateValide(EnteteDateDocument, ComboDate.Text) & ", 102)  WHERE DO_Piece='" & Trim(Documents.DO_Piece) & "' And DO_Type=3" 'And  " & IdentifiantArticle & " ='" & Join(Split(Trim(ListeBL(3)), ","), ".") & "' 
                    ElseIf EnteteTyPeDocument = "14" Then
                        DeleteEcriture = "UPDATE  F_DOCLIGNE SET DL_PieceBC='" & Trim(Documents.DO_Piece) & "', DL_NoRef=" & Trim(NLignePieceCommande) & ",DL_DateBC=CONVERT(DATETIME, " & RenvoieDateValide(EnteteDateDocument, ComboDate.Text) & ", 102)  WHERE DO_Piece='" & Trim(Documents.DO_Piece) & "' And DO_Type=14" 'And  " & IdentifiantArticle & " ='" & Join(Split(Trim(ListeBL(3)), ","), ".") & "' 
                    End If
                    OleComDeletAna = New OleDbCommand(DeleteEcriture)
                    OleComDeletAna.Connection = OleSocieteConnect
                    OleComDeletAna.ExecuteNonQuery()


                    DeleteEcriture = "SET ARITHABORT OFF" 'Optimisation de la requete Sql le 04/02/2015 part Hermann
                    OleComDeletAna = New OleDbCommand(DeleteEcriture)
                    OleComDeletAna.Connection = OleSocieteConnect
                    OleComDeletAna.ExecuteNonQuery()

                    'Documents.CouldModified()
                Catch ex As Exception
                End Try
            End If
            If ListeReliquat.Count <> 0 Then
                Dim OleComDeletAna As OleDbCommand
                Dim DeleteEcriture As String
                For i As Integer = 0 To ListeReliquat.Count - 1
                    Dim ListeBL() As String = Split(ListeReliquat.Item(i), ControlChars.Tab)
                    If Trim(ListeBL(0)) <> "" Then
                        If Trim(ListeBL(1)) <> "" Then
                            If IsDate(Trim(ListeBL(6))) = True Then
                                Try
                                    Documents.Read()
                                    DeleteEcriture = "SET ARITHABORT ON" 'Optimisation de la requete Sql le 04/02/2015 part Hermann
                                    OleComDeletAna = New OleDbCommand(DeleteEcriture)
                                    OleComDeletAna.Connection = OleSocieteConnect
                                    OleComDeletAna.ExecuteNonQuery()

                                    DeleteEcriture = "UPDATE  F_DOCLIGNE SET DL_PieceBC='" & Trim(ListeBL(0)) & "', DL_NoRef=" & Trim(ListeBL(1)) & ",DL_DateBC=CONVERT(DATETIME, '" & Format(CDate(Trim(ListeBL(6))), "yyyy/MM/dd") & "', 102)  WHERE DO_Piece='" & Trim(ListeBL(2)) & "' And  " & IdentifiantArticle & " ='" & Join(Split(Trim(ListeBL(3)), ","), ".") & "' And DO_Type=3"
                                    OleComDeletAna = New OleDbCommand(DeleteEcriture)
                                    OleComDeletAna.Connection = OleSocieteConnect
                                    OleComDeletAna.ExecuteNonQuery()


                                    DeleteEcriture = "SET ARITHABORT OFF" 'Optimisation de la requete Sql le 04/02/2015 part Hermann
                                    OleComDeletAna = New OleDbCommand(DeleteEcriture)
                                    OleComDeletAna.Connection = OleSocieteConnect
                                    OleComDeletAna.ExecuteNonQuery()

                                    Documents.CouldModified()
                                Catch ex As Exception
                                End Try
                            Else
                                Try
                                    Documents.Read()
                                    DeleteEcriture = "SET ARITHABORT ON UPDATE  F_DOCLIGNE SET DL_PieceBC='" & Trim(ListeBL(0)) & "', DL_NoRef=" & Trim(ListeBL(1)) & "  WHERE DO_Piece='" & Trim(ListeBL(2)) & "' And  " & IdentifiantArticle & " ='" & Join(Split(Trim(ListeBL(3)), ","), ".") & "' And DO_Type=3 SET ARITHABORT OFF"
                                    OleComDeletAna = New OleDbCommand(DeleteEcriture)
                                    OleComDeletAna.Connection = OleSocieteConnect
                                    OleComDeletAna.ExecuteNonQuery()
                                    Documents.CouldModified()
                                Catch ex As Exception
                                End Try
                            End If
                        Else
                            If IsDate(Trim(ListeBL(6))) = True Then
                                Try
                                    Documents.Read()
                                    DeleteEcriture = "SET ARITHABORT ON UPDATE  F_DOCLIGNE SET DL_PieceBC='" & Trim(ListeBL(0)) & "',DL_DateBC=CONVERT(DATETIME, '" & Format(CDate(Trim(ListeBL(6))), "yyyy/MM/dd") & "', 102)  WHERE DO_Piece='" & Trim(ListeBL(2)) & "' And  " & IdentifiantArticle & " ='" & Join(Split(Trim(ListeBL(3)), ","), ".") & "'  And DO_Type=3 SET ARITHABORT OFF"
                                    OleComDeletAna = New OleDbCommand(DeleteEcriture)
                                    OleComDeletAna.Connection = OleSocieteConnect
                                    OleComDeletAna.ExecuteNonQuery()
                                    Documents.CouldModified()
                                Catch ex As Exception
                                End Try
                            Else
                                Try
                                    Documents.Read()
                                    DeleteEcriture = "SET ARITHABORT ON UPDATE  F_DOCLIGNE SET DL_PieceBC='" & Trim(ListeBL(0)) & "'  WHERE DO_Piece='" & Trim(ListeBL(2)) & "' And  " & IdentifiantArticle & " ='" & Join(Split(Trim(ListeBL(3)), ","), ".") & "'  And DO_Type=3 SET ARITHABORT OFF"
                                    OleComDeletAna = New OleDbCommand(DeleteEcriture)
                                    OleComDeletAna.Connection = OleSocieteConnect
                                    OleComDeletAna.ExecuteNonQuery()
                                    Documents.CouldModified()
                                Catch ex As Exception
                                End Try
                            End If
                        End If
                    End If
                Next i
                ListeReliquat = New List(Of String)
            End If
            Documents.Write()
            LigneDocument.Write()
        Catch ex As Exception
        End Try
    End Sub
    Public Sub Traitement_Integration(ByVal MyChemin As String)
        Try
            Dim OleAdaptaterschema, OleAdaptaterschemaLigne, OleAdaptaterschemaDétailLigne As OleDbDataAdapter
            Dim OleSchemaDataset, OleSchemaDatasetLigne, OleSchemaDatasetDétailLigne As DataSet
            Dim OledatableSchema, OledatableSchemaLigne, OledatableSchemaDétailLigne As DataTable

            Dim iline As Integer = 0
            Dim iLigne As Integer = 0
            Dim iDetaiLigne As Integer = 0
            Dim NbDetaiLigne As Integer = 0
            Dim NbLigne As Integer = 0
            Dim NbEntete As Integer = 0
            Dim Relation As String = ""
            Dim Statut As Boolean = False

            OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='CSO' AND ENTETE=true ORDER BY ORDRE", OleConnenection)
            OleSchemaDataset = New DataSet
            OleAdaptaterschema.Fill(OleSchemaDataset)
            OledatableSchema = OleSchemaDataset.Tables(0)

            OleAdaptaterschemaLigne = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='CSO' AND ENTETE=False AND Ligne=True ORDER BY ORDRE", OleConnenection)
            OleSchemaDatasetLigne = New DataSet
            OleAdaptaterschemaLigne.Fill(OleSchemaDatasetLigne)
            OledatableSchemaLigne = OleSchemaDatasetLigne.Tables(0)

            OleAdaptaterschemaDétailLigne = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='CSO' AND ENTETE=False AND Ligne=False ORDER BY ORDRE", OleConnenection)
            OleSchemaDatasetDétailLigne = New DataSet
            OleAdaptaterschemaDétailLigne.Fill(OleSchemaDatasetDétailLigne)
            OledatableSchemaDétailLigne = OleSchemaDatasetDétailLigne.Tables(0)

            If OledatableSchema.Rows.Count <> 0 Then
                Dim aRows() As String = Nothing
                Dim Line_Count As Integer = 0
                Dim Detail_Count As Integer = 0
                Dim k As Integer = 1
                Dim k1 As Integer = 1
                Dim Cpteur As Integer = 0
                'Dim LigneQuantiteDemandé As Double
                Dim LigneQuantiteLivre As Double = 0
                Dim LigneCodeArt As String = ""
                Dim CountLigne As Integer = 1
                If GetArrayFile(MyChemin, aRows) IsNot Nothing Then
                    aRows = GetArrayFile(MyChemin, aRows)
                    For i As Integer = 0 To UBound(aRows)
                        Dim Ligne As String = aRows(i)
                        If GetNombreLigne(Ligne, 359, 10) <> 0 Then
                            If Not IsNumeric(Strings.Mid(Ligne, 1, 10)) Then
                                Line_Count = Strings.Mid(Ligne, 359, 10)
                                If Strings.Mid(Ligne, 61, 30).Trim <> "CANCELED" Then
                                    NbEntete += 1
                                    Statut = False
                                    Integrer_Ecriture_Entete(Ligne)

                                    Dim IdentifiantCommande As String ' = "Do_Piece,F_DOCENTETE"
                                    Dim IdentifiantArticle As String '= "DL_Ligne"

                                    IdentifiantArticle = "DL_Ligne"
                                    IdentifiantCommande = "EntetePieceInterne,F_DOCENTETE"
                                    PieceCommande = EntetePieceInterne
                                    PieceArticle = NLignePieceCommande
                                    If NLignePieceCommande <> "" Then
                                        PieceArticle = Convert.ToDecimal(NLignePieceCommande)
                                    End If

                                    'ici je dois traité les lignes de la confirmation de commande
                                    Dim Parcourir As Integer = 1
                                    For Parcourir = 1 To Line_Count
                                        NbLigne += 1
                                        Dim Lignes As String = aRows(k + i)
                                        If RbtG1.Checked Then ' si le client ne gère pas les lots et qu'il ne souhaite pas récupérer le détail dans ses BL
                                            Integrer_Ecriture_Ligne(Lignes, "F_DOCLIGNE")
                                            If Statut = False And StatutCreationEnteteDoc = True And ExisteLecture = True Then
                                                Creation_Entete_Document(EnteteTyPeDocument)
                                                Statut = True
                                            End If
                                            If ExisteLecture = True Then
                                                'Creation_Ligne_Article(ComboDate.Text, Nothing, Nothing, "", "", Nothing, False)
                                                Creation_Ligne_Article(ComboDate.Text, PieceCommande, PieceArticle, IdentifiantCommande, IdentifiantArticle, "", Er_cre_entete_doc)
                                            End If
                                        ElseIf RbtG3.Checked Then ' si le client souhaite récupérer les lots/ sous forme de commentaires
                                            Integrer_Ecriture_Ligne(Lignes, "F_DOCLIGNE")
                                            If Statut = False And StatutCreationEnteteDoc = True And ExisteLecture = True Then
                                                Creation_Entete_Document(EnteteTyPeDocument)
                                                Statut = True
                                            End If
                                            If ExisteLecture = True Then
                                                'Creation_Ligne_Article(ComboDate.Text, Nothing, Nothing, "", "", Nothing, False)
                                                Creation_Ligne_Article(ComboDate.Text, PieceCommande, PieceArticle, IdentifiantCommande, IdentifiantArticle, "", Er_cre_entete_doc)
                                            End If
                                        End If
                                        Detail_Count = Strings.Mid(Lignes, 445, 10)
                                        For j As Integer = 1 To Detail_Count 'ici je dois traite le detail de la ligne de la confirmation de la commande 
                                            Ligne = aRows(k1 + Parcourir)
                                            NbDetaiLigne += 1
                                            If RbtG2.Checked Then 'si le client gère les lots/séries
                                                Integrer_Ecriture_Ligne(Ligne, "F_DOCLIGNE")
                                                If Statut = False And StatutCreationEnteteDoc = True And ExisteLecture = True Then
                                                    Creation_Entete_Document(EnteteTyPeDocument)
                                                    Statut = True
                                                End If
                                                If ExisteLecture = True Then
                                                    'Creation_Ligne_Article(ComboDate.Text, Nothing, Nothing, "", "", Nothing, False)
                                                    Creation_Ligne_Article(ComboDate.Text, PieceCommande, PieceArticle, IdentifiantCommande, IdentifiantArticle, "", Er_cre_entete_doc)
                                                End If
                                            ElseIf RbtG4.Checked Then 'si le client ne gère pas les lots,mais il souhaite récupérer les lots/séries dans une info libre (son traitement est prevue ici)
                                                Integrer_Ecriture_Ligne(Ligne, "F_DOCLIGNE")
                                                If Statut = False And StatutCreationEnteteDoc = True And ExisteLecture = True Then
                                                    Creation_Entete_Document(EnteteTyPeDocument)
                                                    Statut = True
                                                End If
                                                If ExisteLecture = True Then
                                                    ' Creation_Ligne_Article(ComboDate.Text, Nothing, Nothing, "", "", Nothing, False)
                                                    Creation_Ligne_Article(ComboDate.Text, PieceCommande, PieceArticle, IdentifiantCommande, IdentifiantArticle, "", Er_cre_entete_doc)
                                                End If
                                            End If
                                            k1 += 1
                                        Next
                                        k = k + Detail_Count + 1
                                    Next
                                End If
                            End If
                        End If
                        If (NbLigne + NbDetaiLigne + NbEntete) = aRows.Length Then
                            If CountChecked > 1 Then
                                CountChecked -= 1
                            End If
                            If StatutCreationEnteteDoc = True And ExisteLecture = True Then 'And StatutCreationLigneDoc = True
                                infosExport.Text = "Intégration Terminer !"
                                If CheckFille.Checked Then
                                    File.Move(MyChemin, PathsfileSave & "" & Strings.Right(DateAndTime.Year(Now), 2) & "" & Format(DateAndTime.Month(Now), "00") & "" & Format(DateAndTime.Day(Now), "00") & "[" & "" & Format(DateAndTime.Hour(Now), "00") & "-" & Format(DateAndTime.Minute(Now), "00") & "-" & Format(DateAndTime.Second(Now), "00") & "]" & System.IO.Path.GetFileName(Trim(MyChemin)))
                                    infosExport.Refresh()
                                    infosExport.Text = "Integration Terminée! Suppression des Fichiers exécutée..."
                                    DataListeIntegrer.Rows.Clear()
                                    DataListeIntegrerLigne.Rows.Clear()
                                    DataListeIntegrerDétailLigne.Rows.Clear()
                                End If
                                Statut = False
                            End If
                            Exit For
                        Else
                            i = k - 1 + i
                            k = 1
                            k1 = NbLigne + NbDetaiLigne + NbEntete + 1
                            If StatutCreationEnteteDoc = True Then 'And StatutCreationLigneDoc = True
                                Statut = False
                            End If
                        End If
                    Next
                    '''''''''''''''''''''''''''''''''''''''''''
                End If
            End If
        Catch ex As Exception
            MsgBox("Fonction Aperçu Erreur :" & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub
    Public Statut As Boolean = False
    Public Seconde As Integer = 0
    Public Function RenvoiPachsFile(ByVal codeSociete As String, ByVal Categories As String) As String
        Try
            Dim OleAdaptaterschemaCheminIO As OleDbDataAdapter
            Dim OleSchemaDatasetCheminIO As DataSet
            Dim OledatableSchemaCheminIO As DataTable
            OleAdaptaterschemaCheminIO = New OleDbDataAdapter("select distinct CheminFilexport from SCHEMAS_IMPMOUV WHERE BaseCial='" & codeSociete & "' AND Categorie='" & Categories & "'", OleConnenectionClient)
            OleSchemaDatasetCheminIO = New DataSet
            OleAdaptaterschemaCheminIO.Fill(OleSchemaDatasetCheminIO)
            OledatableSchemaCheminIO = OleSchemaDatasetCheminIO.Tables(0)
            Return OledatableSchemaCheminIO.Rows(0).Item("CheminFilexport")
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function RenvoiPachsFileXfert(ByVal codeSociete As String) As String
        Try
            Dim OleAdaptaterschemaCheminIO As OleDbDataAdapter
            Dim OleSchemaDatasetCheminIO As DataSet
            Dim OledatableSchemaCheminIO As DataTable
            OleAdaptaterschemaCheminIO = New OleDbDataAdapter("select distinct CheminFilexport from WIT_SCHEMA WHERE BaseCial='" & codeSociete & "'", OleConnenectionBC)
            OleSchemaDatasetCheminIO = New DataSet
            OleAdaptaterschemaCheminIO.Fill(OleSchemaDatasetCheminIO)
            OledatableSchemaCheminIO = OleSchemaDatasetCheminIO.Tables(0)
            Return OledatableSchemaCheminIO.Rows(0).Item("CheminFilexport")
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public SecondeChampLotCommentaire As String = ""
    Public Sub Transformation(ByVal Chemin As String, Optional ByVal number As Integer = 0)
        Dim OleAdaptaterschema, OleAdaptaterschemaLigne, OleAdaptaterschemaDétailLigne, OleAdaptaterschemaInfosLibre, OleAdaptaterschemaLigneInfosLibre, OleAdaptaterschemaDétailLigneInfosLibre, OleAdaptaterschemaEU_ENUMERE As OleDbDataAdapter
        Dim OleSchemaDataset, OleSchemaDatasetLigne, OleSchemaDatasetDétailLigne, OleSchemaDatasetInfosLibre, OleSchemaDatasetLigneInfosLibre, OleSchemaDatasetDétailLigneInfosLibre, OleSchemaDatasetEU_ENUMERE As DataSet
        Dim OledatableSchema, OledatableSchemaLigne, OledatableSchemaDétailLigne, OledatableSchemaInfosLibre, OledatableSchemaLigneInfosLibre, OledatableSchemaDétailLigneInfosLibre, OledatableSchemaEU_ENUMERE As DataTable
        RegardeStatut = True
        Dim ArtAdaptater As OleDbDataAdapter
        Dim ArtDataset As DataSet
        Dim Artdatatable As DataTable
        Dim CptaAdaptater As OleDbDataAdapter
        Dim CptaDataset As DataSet
        Dim Cptadatatable As DataTable

        Dim Information As String = ""
        Dim InformationLigne As String = ""
        Dim InfosLot As String = ""
        Dim InfosDateExport As String = ""
        Dim iline As Integer = 0
        Dim iLigne As Integer = 0
        Dim iDetaiLigne As Integer = 0
        Dim NbDetaiLigne As Integer = 0
        Dim NbLigne As Integer = 0
        Dim NbEntete As Integer = 0
        Dim EstQuantiteVide As Boolean = False
        Dim Relation As String = ""
        Dim InfosQuantite As Integer = 0
        Dim ReferenceArticle As String = ""
        Dim NumeroLigne1 As Integer = 0
        Dim NumeroLigne2 As Integer = 0
        Dim Piece As String = ""
        Dim DO_Type As String = ""
        Statut = False
        Dim EstPasser As Boolean = False
        Try
            OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='CSO' AND ENTETE=true AND ChampSage<>'' ORDER BY ORDRE", OleConnenection)
            OleSchemaDataset = New DataSet
            OleAdaptaterschema.Fill(OleSchemaDataset)
            OledatableSchema = OleSchemaDataset.Tables(0)

            OleAdaptaterschemaLigne = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='CSO' AND ChampSage<>'' AND Ligne=True  ORDER BY ORDRE", OleConnenection)
            OleSchemaDatasetLigne = New DataSet
            OleAdaptaterschemaLigne.Fill(OleSchemaDatasetLigne)
            OledatableSchemaLigne = OleSchemaDatasetLigne.Tables(0)

            OleAdaptaterschemaDétailLigne = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='CSO' AND ENTETE=False AND ChampSage<>'' AND Ligne=False  ORDER BY ORDRE", OleConnenection)
            OleSchemaDatasetDétailLigne = New DataSet
            OleAdaptaterschemaDétailLigne.Fill(OleSchemaDatasetDétailLigne)
            OledatableSchemaDétailLigne = OleSchemaDatasetDétailLigne.Tables(0)

            'Traitement des infoslibre
            '       En Entete 
            OleAdaptaterschemaInfosLibre = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='CSO' AND ENTETE=true AND InfosLibre=true AND ChampSage<>'' ORDER BY ORDRE", OleConnenection)
            OleSchemaDatasetInfosLibre = New DataSet
            OleAdaptaterschemaInfosLibre.Fill(OleSchemaDatasetInfosLibre)
            OledatableSchemaInfosLibre = OleSchemaDatasetInfosLibre.Tables(0)
            '       En ligne
            OleAdaptaterschemaLigneInfosLibre = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='CSO' AND ENTETE=False AND ChampSage<>'' AND Ligne=True  AND InfosLibre=true  ORDER BY ORDRE", OleConnenection)
            OleSchemaDatasetLigneInfosLibre = New DataSet
            OleAdaptaterschemaLigneInfosLibre.Fill(OleSchemaDatasetLigneInfosLibre)
            OledatableSchemaLigneInfosLibre = OleSchemaDatasetLigneInfosLibre.Tables(0)
            '       Sous Ligne 
            OleAdaptaterschemaDétailLigneInfosLibre = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='CSO' AND ENTETE=False AND ChampSage<>'' AND Ligne=False AND InfosLibre=true ORDER BY ORDRE", OleConnenection)
            OleSchemaDatasetDétailLigneInfosLibre = New DataSet
            OleAdaptaterschemaDétailLigneInfosLibre.Fill(OleSchemaDatasetDétailLigneInfosLibre)
            OledatableSchemaDétailLigneInfosLibre = OleSchemaDatasetDétailLigneInfosLibre.Tables(0)
            '-----------------------------------------------------------------------------
            If OledatableSchema.Rows.Count <> 0 Then
                Dim aRows() As String = Nothing
                Dim Line_Count As Integer = 0
                Dim Detail_Count As Integer = 0
                Dim k As Integer = 1
                Dim k1 As Integer = 1
                Dim Cpteur As Integer = 0
                Seconde += 1
                Dim CountLigne As Integer = 1
                If GetArrayFile(Chemin, aRows) IsNot Nothing Then
                    aRows = GetArrayFile(Chemin, aRows)
                    For i As Integer = 0 To UBound(aRows)
                        Dim Ligne As String = aRows(i)
                        If GetNombreLigne(Ligne, 359, 10) <> 0 Then
                            If Not IsNumeric(Strings.Mid(Ligne, 1, 10)) Then
                                Line_Count = Strings.Mid(Ligne, 359, 10)
                                If Strings.Mid(Ligne, 61, 30).Trim <> "CANCELED" Then
                                    Piece = Strings.Mid(Ligne, 11, 50).Trim
                                    CodeSociete = Strings.Mid(Ligne, 11, 50).Trim
                                    NbEntete += 1
                                    Statut = False
                                    'ici je dois traite l'entete de la confirmation de commande 
                                    For LigneCols As Integer = 0 To OledatableSchema.Rows.Count - 1
                                        Dim Chaine As String = Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format")))
                                        If OledatableSchema.Rows(LigneCols).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Dim Valeurs As Object = Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format")))
                                            If Valeurs.ToString.Trim <> "" Then
                                                Dim Year As String = Valeurs.ToString.Substring(0, 4)
                                                Dim Mois As String = Valeurs.ToString.Substring(4, 2)
                                                Dim Jours As String = Valeurs.ToString.Substring(6, 2)
                                                Dim Heure As String = Valeurs.ToString.Substring(8, 2)
                                                Dim Minute As String = Valeurs.ToString.Substring(10, 2)
                                                Dim Seconde As String = Valeurs.ToString.Substring(12, 2)
                                                Dim MyDate As String = Year & "-" & Mois & "-" & Jours & " " '& Heure & ":" & Minute & ":" & Seconde
                                                Dim MyNewDate As Date = RenvoieDateValide(MyDate, ComboDate.Text)
                                                Information &= MyNewDate & ";"
                                            Else
                                                Information &= Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format"))).Trim & ";"
                                            End If
                                        ElseIf OledatableSchema.Rows(LigneCols).Item("Format").ToString = "N(10)" Then
                                            Dim Valeur As String = Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format"))).Trim
                                            InformationLigne &= Convert.ToDecimal(Valeur) & ";" ' 
                                        Else
                                            If OledatableSchema.Rows(LigneCols).Item("Cols").ToString = "ACCOUNT_CODE" Then
                                                If Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format"))).Trim <> "" Then
                                                    Information &= "3;"
                                                    DO_Type = "3"
                                                    Information &= Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format"))).Trim & ";"
                                                    If EstPasser = False And Achat.Trim.ToUpper = "VENTE" Then
                                                        PathsFileCRP = RenvoiPachsFile(CodeSociete, "Vente")
                                                        If File.Exists(Trim(PathsFileCRP & Seconde & "CW_" & System.IO.Path.GetFileName(Chemin))) = True Then
                                                            File.Delete(Trim(PathsFileCRP & Seconde & "CW_" & System.IO.Path.GetFileName(Chemin)))
                                                            FichierCSO = File.AppendText(PathsFileCRP & Seconde & "CW_" & System.IO.Path.GetFileName(Chemin))
                                                            EstPasser = True
                                                        Else
                                                            FichierCSO = File.AppendText(PathsFileCRP & Seconde & "CW_" & System.IO.Path.GetFileName(Chemin))
                                                            EstPasser = True
                                                        End If
                                                    ElseIf DO_Type = "3" And Achat.Trim.ToUpper = "ACHAT" Then
                                                        EstTrouverException = True
                                                        Exit Sub
                                                    End If
                                                End If
                                            Else
                                                If OledatableSchema.Rows(LigneCols).Item("Cols").ToString = "SUPPLIER_CODE" Then
                                                    If Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format"))).Trim <> "" Then
                                                        Information &= "14;"
                                                        DO_Type = "14"
                                                        Information &= Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format"))).Trim & ";"
                                                        If EstPasser = False And Achat.Trim.ToUpper = "ACHAT" Then
                                                            PathsFileCSO = RenvoiPachsFile(CodeSociete, "Achat")
                                                            If File.Exists(Trim(PathsFileCSO & Seconde & "CW_" & System.IO.Path.GetFileName(Chemin))) = True Then
                                                                File.Delete(Trim(PathsFileCSO & Seconde & "CW_" & System.IO.Path.GetFileName(Chemin)))
                                                                FichierCSO = File.AppendText(PathsFileCSO & Seconde & "CW_" & System.IO.Path.GetFileName(Chemin))
                                                                EstPasser = True
                                                            Else
                                                                FichierCSO = File.AppendText(PathsFileCSO & Seconde & "CW_" & System.IO.Path.GetFileName(Chemin))
                                                                EstPasser = True
                                                            End If
                                                        Else
                                                            If DO_Type = "14" And Achat.Trim.ToUpper = "VENTE" Then
                                                                EstTrouverException = True
                                                                Exit Sub
                                                            End If
                                                        End If
                                                    End If
                                                Else
                                                    If OledatableSchema.Rows(LigneCols).Item("Cols").ToString = "WAREHOUSE_CODE_TO" Then
                                                        If Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format"))).Trim <> "" Then
                                                            Information &= "23;"
                                                            DO_Type = "23"
                                                            Information &= Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format"))).Trim & ";"
                                                            'ici pour le transfert de depot à depot 
                                                            PathsFileXFERT = RenvoiPachsFileXfert(CodeSociete)
                                                            If File.Exists(Trim(PathsFileXFERT & Seconde & "CW_" & System.IO.Path.GetFileName(Chemin))) = True Then
                                                                File.Delete(Trim(PathsFileXFERT & Seconde & "CW_" & System.IO.Path.GetFileName(Chemin)))
                                                                FichierCSO = File.AppendText(PathsFileXFERT & Seconde & "CW_" & System.IO.Path.GetFileName(Chemin))
                                                                EstPasser = True
                                                            Else
                                                                FichierCSO = File.AppendText(PathsFileXFERT & Seconde & "CW_" & System.IO.Path.GetFileName(Chemin))
                                                                EstPasser = True
                                                            End If
                                                        End If
                                                    ElseIf Trim(OledatableSchema.Rows(LigneCols).Item("Cols").ToString.ToUpper) = "DOCUMENT" Then
                                                        CodeSociete = Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format"))).Trim
                                                        If CodeSociete = "" Then
                                                            CodeSociete = OledatableSchema.Rows(LigneCols).Item("DefaultValue").ToString
                                                            Information &= CodeSociete & ";"
                                                        Else
                                                            Information &= Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format"))).Trim & ";"
                                                        End If
                                                    Else
                                                        Information &= Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format"))).Trim & ";"
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Next
                                    If FichierCSO Is Nothing Then
                                        Exit Sub
                                    End If
                                    'Traitement des infos libre en entete de document 
                                    For LigneCols As Integer = 0 To OledatableSchemaInfosLibre.Rows.Count - 1
                                        Dim Tableau() As String = OledatableSchemaInfosLibre.Rows(LigneCols).Item("Format").ToString.Split("(")(1).Split(".")
                                        If Tableau.Length = 2 Then
                                            Dim Divisuer As Integer = Math.Pow(10, Tableau(1).Split(")")(0))
                                            Dim Valeurs As Object = Strings.Mid(Ligne, OledatableSchemaInfosLibre.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaInfosLibre.Rows(LigneCols).Item("Format")))
                                            If Valeurs.ToString.Trim <> "" Then
                                                Information &= CDbl(Valeurs / Divisuer).ToString.Trim & ";" ' 
                                            Else
                                                Information &= ";"
                                            End If
                                        Else
                                            If OledatableSchemaInfosLibre.Rows(LigneCols).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Dim Valeurs As Object = Strings.Mid(Ligne, OledatableSchemaInfosLibre.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaInfosLibre.Rows(LigneCols).Item("Format")))
                                                If Valeurs.ToString.Trim <> "" Then
                                                    Dim Year As String = Valeurs.ToString.Substring(0, 4)
                                                    Dim Mois As String = Valeurs.ToString.Substring(4, 2)
                                                    Dim Jours As String = Valeurs.ToString.Substring(6, 2)
                                                    Dim Heure As String = Valeurs.ToString.Substring(8, 2)
                                                    Dim Minute As String = Valeurs.ToString.Substring(10, 2)
                                                    Dim Seconde As String = Valeurs.ToString.Substring(12, 2)
                                                    Dim MyDate As String = Year & "-" & Mois & "-" & Jours & " " & Heure & ":" & Minute & ":" & Seconde
                                                    Dim MyNewDate As Date = RenvoieDateValide(MyDate, ComboDate.Text)
                                                    Information &= MyDate.Trim & ";" '
                                                Else
                                                    Information &= Strings.Mid(Ligne, OledatableSchemaInfosLibre.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaInfosLibre.Rows(LigneCols).Item("Format"))).Trim & ";" ' 
                                                End If
                                            ElseIf OledatableSchemaInfosLibre.Rows(LigneCols).Item("Format").ToString = "N(10)" Then
                                                Dim Valeur As String = Strings.Mid(Ligne, OledatableSchemaInfosLibre.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaInfosLibre.Rows(LigneCols).Item("Format"))).Trim
                                                If Valeur.ToString <> "" Then
                                                    Information &= Convert.ToDecimal(Valeur) & ";"
                                                Else
                                                    Information &= ";"
                                                End If
                                            Else
                                                Information &= Strings.Mid(Ligne, OledatableSchemaInfosLibre.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaInfosLibre.Rows(LigneCols).Item("Format"))).Trim & ";" ' 
                                            End If
                                        End If
                                    Next
                                    '---------------------FIN Traitement des infos Libre --------------------------
                                    'ici je dois traité les lignes de la confirmation de commande
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
                                    Dim Parcourir As Integer = 1
                                    For Parcourir = 1 To Line_Count
                                        NbLigne += 1
                                        Dim Lignes As String = aRows(k + i)
                                        If LOT = False Then
                                            For LigneCols As Integer = 0 To OledatableSchemaLigne.Rows.Count - 1
                                                Dim Tableau() As String = OledatableSchemaLigne.Rows(LigneCols).Item("Format").ToString.Split("(")(1).Split(".")
                                                If Tableau.Length = 2 Then
                                                    Dim Divisuer As Integer = Math.Pow(10, Tableau(1).Split(")")(0))
                                                    Dim Valeurs As Object = Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format")))
                                                    If OledatableSchemaLigne.Rows(LigneCols).Item("Cols").ToString.ToUpper = "LINE_NUMBER" Then
                                                        If Valeurs.ToString.Trim <> "" Then
                                                            NumeroLigne1 = CDbl(Valeurs / Divisuer)
                                                            InformationLigne &= CDbl(Valeurs / Divisuer).ToString.Trim & ";" ' 
                                                        Else
                                                            InformationLigne &= ";"
                                                        End If
                                                    ElseIf OledatableSchemaLigne.Rows(LigneCols).Item("Cols").ToString.ToUpper = "CUST_LINE_NUMBER" Then
                                                        If Valeurs.ToString.Trim <> "" Then
                                                            InformationLigne &= CDbl(Valeurs / Divisuer).ToString.Trim & ";" ' 
                                                            NumeroLigne2 = CDbl(Valeurs / Divisuer)
                                                        Else
                                                            InformationLigne &= ";"
                                                        End If
                                                    ElseIf OledatableSchemaLigne.Rows(LigneCols).Item("Cols").ToString.ToUpper = "SERVED_QUANTITY" Then
                                                        If Valeurs.ToString.Trim <> "" Then
                                                            If CDbl(Valeurs / Divisuer) <> 0 Then
                                                                InformationLigne &= CDbl(Valeurs / Divisuer).ToString.Trim & ";" ' 
                                                                EstQuantiteVide = False
                                                            Else
                                                                'InformationLigne &= CDbl(Valeurs / Divisuer).ToString.Trim & ";" ' 
                                                                EstQuantiteVide = True
                                                            End If
                                                        Else
                                                            InformationLigne &= ";"
                                                        End If
                                                    Else
                                                        If Valeurs.ToString.Trim <> "" Then
                                                            InformationLigne &= CDbl(Valeurs / Divisuer).ToString.Trim & ";" ' 
                                                        Else
                                                            InformationLigne &= ";"
                                                        End If
                                                    End If
                                                Else
                                                    If OledatableSchemaLigne.Rows(LigneCols).Item("Format").ToString.Split("(")(0) = "D" Then
                                                        Dim Valeurs As Object = Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format")))
                                                        If Valeurs.ToString.Trim <> "" Then
                                                            Dim Year As String = Valeurs.ToString.Substring(0, 4)
                                                            Dim Mois As String = Valeurs.ToString.Substring(4, 2)
                                                            Dim Jours As String = Valeurs.ToString.Substring(6, 2)
                                                            Dim Heure As String = Valeurs.ToString.Substring(8, 2)
                                                            Dim Minute As String = Valeurs.ToString.Substring(10, 2)
                                                            Dim Seconde As String = Valeurs.ToString.Substring(12, 2)
                                                            Dim MyDate As String = Year & "-" & Mois & "-" & Jours & " " & Heure & ":" & Minute & ":" & Seconde
                                                            Dim MyNewDate As Date = RenvoieDateValide(MyDate, ComboDate.Text)
                                                            InformationLigne &= MyDate.Trim & ";" '
                                                        Else
                                                            InformationLigne &= Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format"))).Trim & ";" ' 
                                                        End If
                                                    ElseIf OledatableSchemaLigne.Rows(LigneCols).Item("Format").ToString = "N(10)" Then
                                                        Dim Valeur As String = Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format"))).Trim
                                                        If Valeur.ToString <> "" Then
                                                            InformationLigne &= Convert.ToDecimal(Valeur) & ";"
                                                        Else
                                                            InformationLigne &= ";"
                                                        End If
                                                    Else
                                                        Select Case OledatableSchemaLigne.Rows(LigneCols).Item("Cols").ToString.ToUpper
                                                            Case "PRODUCT_CODE"
                                                                InformationLigne &= Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format"))).Trim & ";"   ' 
                                                                ReferenceArticle = Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format"))).Trim
                                                            Case "UOM_CODE"
                                                                Dim EU_ENUMERE As String = Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format"))).Trim
                                                                OleAdaptaterschemaEU_ENUMERE = New OleDbDataAdapter("SELECT EU_Enumere FROM dbo.F_DOCLIGNE WHERE (DO_Type = " & DO_Type & ") AND (DO_Piece = '" & Piece & "') AND (DL_Ligne = " & NumeroLigne1 & ")", OleSocieteConnect)
                                                                OleSchemaDatasetEU_ENUMERE = New DataSet
                                                                OleAdaptaterschemaEU_ENUMERE.Fill(OleSchemaDatasetEU_ENUMERE)
                                                                OledatableSchemaEU_ENUMERE = OleSchemaDatasetEU_ENUMERE.Tables(0)
                                                                If OledatableSchemaEU_ENUMERE.Rows.Count <> 0 Then
                                                                    InformationLigne &= OledatableSchemaEU_ENUMERE.Rows(0).Item("EU_Enumere") & ";"  ' 
                                                                Else
                                                                    Dim OleAdaptaterschemaEU_ENUMERE2 As OleDbDataAdapter
                                                                    Dim OleSchemaDatasetEU_ENUMERE2 As DataSet
                                                                    Dim OledatableSchemaEU_ENUMERE2 As DataTable
                                                                    OleAdaptaterschemaEU_ENUMERE2 = New OleDbDataAdapter("SELECT  dbo.F_CONDITION.EC_Enumere FROM dbo.P_CONDITIONNEMENT RIGHT OUTER JOIN dbo.F_ARTICLE ON dbo.P_CONDITIONNEMENT.cbMarq = dbo.F_ARTICLE.AR_Condition LEFT OUTER JOIN dbo.F_CONDITION ON dbo.F_ARTICLE.AR_Ref = dbo.F_CONDITION.AR_Ref WHERE (dbo.F_ARTICLE.AR_Ref = '" & ReferenceArticle & "') AND (dbo.P_CONDITIONNEMENT.P_Conditionnement = '" & EU_ENUMERE & "') AND (dbo.F_CONDITION.CO_Principal = 1)", OleSocieteConnect)
                                                                    OleSchemaDatasetEU_ENUMERE2 = New DataSet
                                                                    OleAdaptaterschemaEU_ENUMERE2.Fill(OleSchemaDatasetEU_ENUMERE2)
                                                                    OledatableSchemaEU_ENUMERE2 = OleSchemaDatasetEU_ENUMERE2.Tables(0)
                                                                    If OledatableSchemaEU_ENUMERE2.Rows.Count <> 0 Then
                                                                        InformationLigne &= OledatableSchemaEU_ENUMERE2.Rows(0).Item("EC_Enumere") & ";"  '
                                                                    Else
                                                                        InformationLigne &= ";"  '
                                                                    End If
                                                                End If
                                                            Case Else
                                                                InformationLigne &= Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format"))).Trim & ";"  ' 
                                                        End Select
                                                    End If
                                                End If
                                            Next
                                            '------------Traitement des Infos Libres en Lignes de Document 
                                            For LigneCols As Integer = 0 To OledatableSchemaLigneInfosLibre.Rows.Count - 1
                                                Dim Tableau() As String = OledatableSchemaLigneInfosLibre.Rows(LigneCols).Item("Format").ToString.Split("(")(1).Split(".")
                                                If Tableau.Length = 2 Then
                                                    Dim Divisuer As Integer = Math.Pow(10, Tableau(1).Split(")")(0))
                                                    Dim Valeurs As Object = Strings.Mid(Lignes, OledatableSchemaLigneInfosLibre.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigneInfosLibre.Rows(LigneCols).Item("Format")))
                                                    If Valeurs.ToString.Trim <> "" Then
                                                        InformationLigne &= CDbl(Valeurs / Divisuer).ToString.Trim & ";" ' 
                                                    Else
                                                        InformationLigne &= ";"
                                                    End If
                                                Else
                                                    If OledatableSchemaLigneInfosLibre.Rows(LigneCols).Item("Format").ToString.Split("(")(0) = "D" Then
                                                        Dim Valeurs As Object = Strings.Mid(Lignes, OledatableSchemaLigneInfosLibre.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigneInfosLibre.Rows(LigneCols).Item("Format")))
                                                        If Valeurs.ToString.Trim <> "" Then
                                                            Dim Year As String = Valeurs.ToString.Substring(0, 4)
                                                            Dim Mois As String = Valeurs.ToString.Substring(4, 2)
                                                            Dim Jours As String = Valeurs.ToString.Substring(6, 2)
                                                            Dim Heure As String = Valeurs.ToString.Substring(8, 2)
                                                            Dim Minute As String = Valeurs.ToString.Substring(10, 2)
                                                            Dim Seconde As String = Valeurs.ToString.Substring(12, 2)
                                                            Dim MyDate As String = Year & "-" & Mois & "-" & Jours & " " & Heure & ":" & Minute & ":" & Seconde
                                                            Dim MyNewDate As Date = RenvoieDateValide(MyDate, ComboDate.Text)
                                                            InformationLigne &= MyDate.Trim & ";" '
                                                        Else
                                                            InformationLigne &= Strings.Mid(Lignes, OledatableSchemaLigneInfosLibre.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigneInfosLibre.Rows(LigneCols).Item("Format"))).Trim & ";" ' 
                                                        End If
                                                    ElseIf OledatableSchemaLigne.Rows(LigneCols).Item("Format").ToString = "N(10)" Then
                                                        Dim Valeur As String = Strings.Mid(Lignes, OledatableSchemaLigneInfosLibre.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigneInfosLibre.Rows(LigneCols).Item("Format"))).Trim
                                                        If Valeur.ToString <> "" Then
                                                            InformationLigne &= Convert.ToDecimal(Valeur) & ";"
                                                        Else
                                                            InformationLigne &= ";"
                                                        End If
                                                    Else
                                                        If OledatableSchemaLigneInfosLibre.Rows(LigneCols).Item("ChampSage").ToString.ToUpper = "AR_REF" Then
                                                            InformationLigne &= Strings.Mid(Lignes, OledatableSchemaLigneInfosLibre.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigneInfosLibre.Rows(LigneCols).Item("Format"))).Trim & ";"   ' 
                                                        Else
                                                            InformationLigne &= Strings.Mid(Lignes, OledatableSchemaLigneInfosLibre.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigneInfosLibre.Rows(LigneCols).Item("Format"))).Trim & ";"  ' 
                                                        End If
                                                    End If
                                                End If
                                            Next
                                            '---------------------------Fin de Traitement ------------------------------------------------
                                            Detail_Count = Strings.Mid(Lignes, 445, 10)
                                            For j As Integer = 1 To Detail_Count 'ici je dois traite le detail de la ligne de la confirmation de la commande 
                                                Ligne = aRows(k1 + Parcourir)
                                                NbDetaiLigne += 1
                                                If j = 1 Then
                                                    For LigneCols As Integer = 0 To OledatableSchemaDétailLigne.Rows.Count - 1
                                                        Dim Tableau() As String = OledatableSchemaDétailLigne.Rows(LigneCols).Item("Format").ToString.Split("(")(1).Split(".")
                                                        If LigneCols = OledatableSchemaDétailLigne.Rows.Count - 1 Then
                                                            If Tableau.Length = 2 Then
                                                                Dim Divisuer As Integer = Math.Pow(10, Tableau(1).Split(")")(0))
                                                                Dim Valeurs As Object = Strings.Mid(Ligne, OledatableSchemaDétailLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaDétailLigne.Rows(LigneCols).Item("Format")))
                                                                If OledatableSchemaDétailLigne.Rows(LigneCols).Item("Cols").ToString.Trim = "SHIPPED_QUANTITY" Then
                                                                    If Valeurs.ToString.Trim <> "" Then
                                                                        For ilineD As Integer = 0 To DataListeIntegrerDétailLigne.Rows.Count - 1
                                                                            If DataListeIntegrerDétailLigne.Rows(ilineD).Cells("Liens").Value = Piece And DataListeIntegrerDétailLigne.Rows(ilineD).Cells("RE").Value = ReferenceArticle Then
                                                                                InfosQuantite += CDbl(DataListeIntegrerDétailLigne.Rows(ilineD).Cells("SHIPPED_QUANTITY").Value)
                                                                            End If
                                                                        Next
                                                                        InformationLigne &= InfosQuantite 'CDbl(Valeurs / Divisuer).ToString.Trim & ";" ' 
                                                                    Else
                                                                        InformationLigne &= ""
                                                                    End If
                                                                Else
                                                                    If Valeurs.ToString.Trim <> "" Then
                                                                        InformationLigne &= CDbl(Valeurs / Divisuer).ToString.Trim ' 
                                                                    Else
                                                                        InformationLigne &= ""
                                                                    End If
                                                                End If
                                                                
                                                            Else
                                                                If OledatableSchemaDétailLigne.Rows(LigneCols).Item("Format").ToString.Split("(")(0) = "D" Then
                                                                    Dim Valeurs As Object = Strings.Mid(Ligne, OledatableSchemaDétailLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaDétailLigne.Rows(LigneCols).Item("Format"))).Trim
                                                                    If Valeurs.ToString.Trim <> "" Then
                                                                        Dim Year As String = Valeurs.ToString.Substring(0, 4)
                                                                        Dim Mois As String = Valeurs.ToString.Substring(4, 2)
                                                                        Dim Jours As String = Valeurs.ToString.Substring(6, 2)
                                                                        Dim Heure As String = Valeurs.ToString.Substring(8, 2)
                                                                        Dim Minute As String = Valeurs.ToString.Substring(10, 2)
                                                                        Dim Seconde As String = Valeurs.ToString.Substring(12, 2)
                                                                        Dim MyDate As String = Year & "-" & Mois & "-" & Jours '& " " & Heure & ":" & Minute & ":" & Seconde
                                                                        Dim MyNewDate As Date = RenvoieDateValide(MyDate, ComboDate.Text)
                                                                        InformationLigne &= MyNewDate & ";"
                                                                    Else
                                                                        InformationLigne &= Strings.Mid(Ligne, OledatableSchemaDétailLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaDétailLigne.Rows(LigneCols).Item("Format"))).Trim & ";" ' 
                                                                    End If
                                                                ElseIf OledatableSchemaDétailLigne.Rows(LigneCols).Item("Format").ToString = "N(10)" Then
                                                                    Dim Valeur As String = Strings.Mid(Ligne, OledatableSchemaDétailLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaDétailLigne.Rows(LigneCols).Item("Format"))).Trim
                                                                    If Valeur.ToString <> "" Then
                                                                        InformationLigne &= Convert.ToDecimal(Valeur)
                                                                    Else
                                                                        InformationLigne &= ""
                                                                    End If
                                                                Else
                                                                    InformationLigne &= Strings.Mid(Ligne, OledatableSchemaDétailLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaDétailLigne.Rows(LigneCols).Item("Format"))).Trim  ' 
                                                                End If
                                                            End If
                                                            'Traitement des infos libres sur les details de ligne 
                                                            For LigneColss As Integer = 0 To OledatableSchemaDétailLigneInfosLibre.Rows.Count - 1
                                                                Dim Tableau1() As String = OledatableSchemaDétailLigneInfosLibre.Rows(LigneCols).Item("Format").ToString.Split("(")(1).Split(".")
                                                                If LigneCols = OledatableSchemaDétailLigneInfosLibre.Rows.Count - 1 Then
                                                                    If Tableau1.Length = 2 Then
                                                                        Dim Divisuer As Integer = Math.Pow(10, Tableau1(1).Split(")")(0))
                                                                        Dim Valeurs As Object = Strings.Mid(Ligne, OledatableSchemaDétailLigneInfosLibre.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaDétailLigneInfosLibre.Rows(LigneCols).Item("Format")))
                                                                        If OledatableSchemaDétailLigneInfosLibre.Rows(LigneCols).Item("Cols").ToString.Trim = "SHIPPED_QUANTITY" Then
                                                                            If Valeurs.ToString.Trim <> "" Then
                                                                                For ilineD As Integer = 0 To DataListeIntegrerDétailLigne.Rows.Count - 1
                                                                                    If DataListeIntegrerDétailLigne.Rows(ilineD).Cells("Liens").Value = Piece And DataListeIntegrerDétailLigne.Rows(ilineD).Cells("RE").Value = ReferenceArticle Then
                                                                                        InfosQuantite += CDbl(DataListeIntegrerDétailLigne.Rows(ilineD).Cells("SHIPPED_QUANTITY").Value)
                                                                                    End If
                                                                                Next
                                                                                InformationLigne &= InfosQuantite  ' 
                                                                            Else
                                                                                InformationLigne &= ""
                                                                            End If
                                                                        Else
                                                                            If Valeurs.ToString.Trim <> "" Then
                                                                                InformationLigne &= CDbl(Valeurs / Divisuer).ToString.Trim ' 
                                                                            Else
                                                                                InformationLigne &= ""
                                                                            End If
                                                                        End If
                                                                    Else
                                                                        If OledatableSchemaDétailLigneInfosLibre.Rows(LigneCols).Item("Format").ToString.Split("(")(0) = "D" Then
                                                                            Dim Valeurs As Object = Strings.Mid(Ligne, OledatableSchemaDétailLigneInfosLibre.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaDétailLigneInfosLibre.Rows(LigneCols).Item("Format"))).Trim
                                                                            If Valeurs.ToString.Trim <> "" Then
                                                                                Dim Year As String = Valeurs.ToString.Substring(0, 4)
                                                                                Dim Mois As String = Valeurs.ToString.Substring(4, 2)
                                                                                Dim Jours As String = Valeurs.ToString.Substring(6, 2)
                                                                                Dim Heure As String = Valeurs.ToString.Substring(8, 2)
                                                                                Dim Minute As String = Valeurs.ToString.Substring(10, 2)
                                                                                Dim Seconde As String = Valeurs.ToString.Substring(12, 2)
                                                                                Dim MyDate As String = Year & "-" & Mois & "-" & Jours '& " " & Heure & ":" & Minute & ":" & Seconde
                                                                                Dim MyNewDate As Date = RenvoieDateValide(MyDate, ComboDate.Text)
                                                                                InformationLigne &= MyNewDate & ";"
                                                                            Else
                                                                                InformationLigne &= Strings.Mid(Ligne, OledatableSchemaDétailLigneInfosLibre.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaDétailLigneInfosLibre.Rows(LigneCols).Item("Format"))).Trim & ";" ' 
                                                                            End If
                                                                        ElseIf OledatableSchemaDétailLigneInfosLibre.Rows(LigneCols).Item("Format").ToString = "N(10)" Then
                                                                            Dim Valeur As String = Strings.Mid(Ligne, OledatableSchemaDétailLigneInfosLibre.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaDétailLigneInfosLibre.Rows(LigneCols).Item("Format"))).Trim
                                                                            If Valeur.ToString <> "" Then
                                                                                InformationLigne &= Convert.ToDecimal(Valeur)
                                                                            Else
                                                                                InformationLigne &= ""
                                                                            End If
                                                                        Else
                                                                            InformationLigne &= Strings.Mid(Ligne, OledatableSchemaDétailLigneInfosLibre.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaDétailLigneInfosLibre.Rows(LigneCols).Item("Format"))).Trim  ' 
                                                                        End If
                                                                    End If
                                                                Else
                                                                    If Tableau1.Length = 2 Then
                                                                        Dim Divisuer As Integer = Math.Pow(10, Tableau1(1).Split(")")(0))
                                                                        Dim Valeurs As Object = Strings.Mid(Ligne, OledatableSchemaDétailLigneInfosLibre.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaDétailLigneInfosLibre.Rows(LigneCols).Item("Format")))
                                                                        If OledatableSchemaDétailLigneInfosLibre.Rows(LigneCols).Item("Cols").ToString.Trim = "SHIPPED_QUANTITY" Then
                                                                            If Valeurs.ToString.Trim <> "" Then
                                                                                For ilineD As Integer = 0 To DataListeIntegrerDétailLigne.Rows.Count - 1
                                                                                    If DataListeIntegrerDétailLigne.Rows(ilineD).Cells("Liens").Value = Piece And DataListeIntegrerDétailLigne.Rows(ilineD).Cells("RE").Value = ReferenceArticle Then
                                                                                        InfosQuantite += CDbl(DataListeIntegrerDétailLigne.Rows(ilineD).Cells("SHIPPED_QUANTITY").Value)
                                                                                    End If
                                                                                Next
                                                                                InformationLigne &= InfosQuantite
                                                                            Else
                                                                                InformationLigne &= ""
                                                                            End If
                                                                        Else
                                                                            If Valeurs.ToString.Trim <> "" Then
                                                                                InformationLigne &= CDbl(Valeurs / Divisuer).ToString.Trim ' 
                                                                            Else
                                                                                InformationLigne &= ""
                                                                            End If
                                                                        End If
                                                                    Else
                                                                        If OledatableSchemaDétailLigneInfosLibre.Rows(LigneCols).Item("Format").ToString.Split("(")(0) = "D" Then
                                                                            Dim Valeurs As Object = Strings.Mid(Ligne, OledatableSchemaDétailLigneInfosLibre.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaDétailLigneInfosLibre.Rows(LigneCols).Item("Format"))).Trim
                                                                            If Valeurs.ToString.Trim <> "" Then
                                                                                Dim Year As String = Valeurs.ToString.Substring(0, 4)
                                                                                Dim Mois As String = Valeurs.ToString.Substring(4, 2)
                                                                                Dim Jours As String = Valeurs.ToString.Substring(6, 2)
                                                                                Dim Heure As String = Valeurs.ToString.Substring(8, 2)
                                                                                Dim Minute As String = Valeurs.ToString.Substring(10, 2)
                                                                                Dim Seconde As String = Valeurs.ToString.Substring(12, 2)
                                                                                Dim MyDate As String = Year & "-" & Mois & "-" & Jours '& " " & Heure & ":" & Minute & ":" & Seconde
                                                                                Dim MyNewDate As Date = RenvoieDateValide(MyDate, ComboDate.Text)
                                                                                InformationLigne &= MyNewDate & ";"
                                                                            Else
                                                                                InformationLigne &= Strings.Mid(Ligne, OledatableSchemaDétailLigneInfosLibre.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaDétailLigneInfosLibre.Rows(LigneCols).Item("Format"))).Trim & ";" ' 
                                                                            End If
                                                                        ElseIf OledatableSchemaDétailLigneInfosLibre.Rows(LigneCols).Item("Format").ToString = "N(10)" Then
                                                                            Dim Valeur As String = Strings.Mid(Ligne, OledatableSchemaDétailLigneInfosLibre.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaDétailLigneInfosLibre.Rows(LigneCols).Item("Format"))).Trim
                                                                            If Valeur.ToString <> "" Then
                                                                                InformationLigne &= Convert.ToDecimal(Valeur) & ";"
                                                                            Else
                                                                                InformationLigne &= ";"
                                                                            End If
                                                                        Else
                                                                            InformationLigne &= Strings.Mid(Ligne, OledatableSchemaDétailLigneInfosLibre.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaDétailLigneInfosLibre.Rows(LigneCols).Item("Format"))).Trim & ";" ' 
                                                                        End If
                                                                    End If
                                                                End If
                                                            Next
                                                            '-----------------------Fin de Traitement Infos Libre SL ---------------------
                                                            Select Case Achat.Trim.ToUpper
                                                                Case "ACHAT"
                                                                    If EstQuantiteVide = False Then
                                                                        FichierCSO.WriteLine(Information & InformationLigne)
                                                                    End If
                                                                    InformationLigne = ""
                                                                    InfosQuantite = 0
                                                                Case "VENTE"
                                                                    If EstQuantiteVide = False Then
                                                                        FichierCSO.WriteLine(Information & InformationLigne)
                                                                    End If
                                                                    InformationLigne = ""
                                                                    InfosQuantite = 0
                                                                Case "TRANSFERT"
                                                                    If DO_Type = "14" Then
                                                                        FichierCSO.WriteLine(Information & InformationLigne)
                                                                    End If
                                                                    InformationLigne = ""
                                                                    InfosQuantite = 0
                                                            End Select
                                                        Else
                                                            If Tableau.Length = 2 Then
                                                                Dim Divisuer As Integer = Math.Pow(10, Tableau(1).Split(")")(0))
                                                                Dim Valeurs As Object = Strings.Mid(Ligne, OledatableSchemaDétailLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaDétailLigne.Rows(LigneCols).Item("Format")))
                                                                If OledatableSchemaDétailLigne.Rows(LigneCols).Item("Cols").ToString.Trim = "SHIPPED_QUANTITY" Then
                                                                    If Valeurs.ToString.Trim <> "" Then
                                                                        For ilineD As Integer = 0 To DataListeIntegrerDétailLigne.Rows.Count - 1
                                                                            If DataListeIntegrerDétailLigne.Rows(ilineD).Cells("Liens").Value = Piece And DataListeIntegrerDétailLigne.Rows(ilineD).Cells("RE").Value = ReferenceArticle Then
                                                                                InfosQuantite += CDbl(DataListeIntegrerDétailLigne.Rows(ilineD).Cells("SHIPPED_QUANTITY").Value)
                                                                            End If
                                                                        Next
                                                                        InformationLigne &= InfosQuantite & ";" 'CDbl(Valeurs / Divisuer).ToString.Trim & ";" ' 
                                                                    Else
                                                                        InformationLigne &= ";"
                                                                    End If
                                                                Else
                                                                    If Valeurs.ToString.Trim <> "" Then
                                                                        InformationLigne &= CDbl(Valeurs / Divisuer).ToString.Trim & ";" ' 
                                                                    Else
                                                                        InformationLigne &= ";"
                                                                    End If
                                                                End If
                                                            Else
                                                                If OledatableSchemaDétailLigne.Rows(LigneCols).Item("Format").ToString.Split("(")(0) = "D" Then
                                                                    Dim Valeurs As Object = Strings.Mid(Ligne, OledatableSchemaDétailLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaDétailLigne.Rows(LigneCols).Item("Format"))).Trim
                                                                    If Valeurs.ToString.Trim <> "" Then
                                                                        Dim Year As String = Valeurs.ToString.Substring(0, 4)
                                                                        Dim Mois As String = Valeurs.ToString.Substring(4, 2)
                                                                        Dim Jours As String = Valeurs.ToString.Substring(6, 2)
                                                                        Dim Heure As String = Valeurs.ToString.Substring(8, 2)
                                                                        Dim Minute As String = Valeurs.ToString.Substring(10, 2)
                                                                        Dim Seconde As String = Valeurs.ToString.Substring(12, 2)
                                                                        Dim MyDate As String = Year & "-" & Mois & "-" & Jours & " " & Heure & ":" & Minute & ":" & Seconde
                                                                        Dim MyNewDate As Date = RenvoieDateValide(MyDate, ComboDate.Text)
                                                                        InformationLigne &= MyNewDate & ";"
                                                                    Else
                                                                        InformationLigne &= Strings.Mid(Ligne, OledatableSchemaDétailLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaDétailLigne.Rows(LigneCols).Item("Format"))).Trim & ";" ' 
                                                                    End If
                                                                ElseIf OledatableSchemaDétailLigne.Rows(LigneCols).Item("Format").ToString = "N(10)" Then
                                                                    Dim Valeur As String = Strings.Mid(Ligne, OledatableSchemaDétailLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaDétailLigne.Rows(LigneCols).Item("Format"))).Trim
                                                                    If Valeur.ToString <> "" Then
                                                                        InformationLigne &= Convert.ToDecimal(Valeur) & ";"
                                                                    Else
                                                                        InformationLigne &= ";"
                                                                    End If
                                                                Else
                                                                    InformationLigne &= Strings.Mid(Ligne, OledatableSchemaDétailLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaDétailLigne.Rows(LigneCols).Item("Format"))).Trim & ";" ' 
                                                                End If
                                                            End If
                                                        End If
                                                    Next
                                                End If
                                                k1 += 1
                                            Next
                                            If Detail_Count = 0 Then
                                                For LigneCols As Integer = 0 To OledatableSchemaDétailLigne.Rows.Count - 1
                                                    If LigneCols = OledatableSchemaDétailLigne.Rows.Count - 1 Then
                                                        Select Case Achat.Trim.ToUpper
                                                            Case "ACHAT"
                                                                If EstQuantiteVide = False Then
                                                                    FichierCSO.WriteLine(Information & InformationLigne)
                                                                End If
                                                                InformationLigne = ""
                                                            Case "VENTE"
                                                                If EstQuantiteVide = False Then
                                                                    FichierCSO.WriteLine(Information & InformationLigne)
                                                                End If
                                                                InformationLigne = ""
                                                            Case "TRANSFERT"
                                                                If DO_Type = "14" Then
                                                                    FichierCSO.WriteLine(Information & InformationLigne)
                                                                End If
                                                                InformationLigne = ""
                                                        End Select
                                                    Else
                                                        InformationLigne &= ";"
                                                    End If
                                                Next
                                            End If
                                            k = k + Detail_Count + 1
                                        Else 'Gestion des Lot Commentaires --------------------------------------------------------------------------
                                            For LigneCols As Integer = 0 To OledatableSchemaLigne.Rows.Count - 1
                                                Dim Tableau() As String = OledatableSchemaLigne.Rows(LigneCols).Item("Format").ToString.Split("(")(1).Split(".")
                                                If LigneCols = OledatableSchemaLigne.Rows.Count - 1 Then
                                                    If Tableau.Length = 2 Then
                                                        Dim Divisuer As Integer = Math.Pow(10, Tableau(1).Split(")")(0))
                                                        Dim Valeurs As Object = Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format")))
                                                        If OledatableSchemaLigne.Rows(LigneCols).Item("Cols").ToString.ToUpper = "LINE_NUMBER" Then
                                                            If Valeurs.ToString.Trim <> "" Then
                                                                NumeroLigne1 = CDbl(Valeurs / Divisuer)
                                                                InformationLigne &= CDbl(Valeurs / Divisuer).ToString.Trim & "" ' 
                                                            Else
                                                                InformationLigne &= ""
                                                            End If
                                                        ElseIf OledatableSchemaLigne.Rows(LigneCols).Item("Cols").ToString.ToUpper = "CUST_LINE_NUMBER" Then
                                                            If Valeurs.ToString.Trim <> "" Then
                                                                InformationLigne &= CDbl(Valeurs / Divisuer).ToString.Trim ' 
                                                                NumeroLigne2 = CDbl(Valeurs / Divisuer)
                                                            Else
                                                                InformationLigne &= ""
                                                            End If
                                                        ElseIf OledatableSchemaLigne.Rows(LigneCols).Item("Cols").ToString.ToUpper = "SERVED_QUANTITY" Then
                                                            If Valeurs.ToString.Trim <> "" Then
                                                                If CDbl(Valeurs / Divisuer) <> 0 Then
                                                                    InformationLigne &= CDbl(Valeurs / Divisuer).ToString.Trim & ";" ' 
                                                                    EstQuantiteVide = False
                                                                Else
                                                                    EstQuantiteVide = True
                                                                End If
                                                            Else
                                                                InformationLigne &= ";"
                                                            End If
                                                        Else
                                                            If Valeurs.ToString.Trim <> "" Then
                                                                InformationLigne &= CDbl(Valeurs / Divisuer).ToString.Trim  ' 
                                                            Else
                                                                InformationLigne &= ""
                                                            End If
                                                        End If
                                                    Else
                                                        If OledatableSchemaLigne.Rows(LigneCols).Item("Format").ToString.Split("(")(0) = "D" Then
                                                            Dim Valeurs As Object = Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format")))
                                                            If Valeurs.ToString.Trim <> "" Then
                                                                Dim Year As String = Valeurs.ToString.Substring(0, 4)
                                                                Dim Mois As String = Valeurs.ToString.Substring(4, 2)
                                                                Dim Jours As String = Valeurs.ToString.Substring(6, 2)
                                                                Dim Heure As String = Valeurs.ToString.Substring(8, 2)
                                                                Dim Minute As String = Valeurs.ToString.Substring(10, 2)
                                                                Dim Seconde As String = Valeurs.ToString.Substring(12, 2)
                                                                Dim MyDate As String = Year & "-" & Mois & "-" & Jours & " " & Heure & ":" & Minute & ":" & Seconde
                                                                Dim MyNewDate As Date = RenvoieDateValide(MyDate, ComboDate.Text)
                                                                InformationLigne &= MyDate.Trim
                                                            Else
                                                                InformationLigne &= Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format"))).Trim  ' 
                                                            End If
                                                        ElseIf OledatableSchemaLigne.Rows(LigneCols).Item("Format").ToString = "N(10)" Then
                                                            Dim Valeur As String = Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format"))).Trim
                                                            If Valeur.ToString <> "" Then
                                                                InformationLigne &= Convert.ToDecimal(Valeur) & ""
                                                            Else
                                                                InformationLigne &= ""
                                                            End If
                                                        Else
                                                            Select Case OledatableSchemaLigne.Rows(LigneCols).Item("Cols").ToString.ToUpper
                                                                Case "PRODUCT_CODE"
                                                                    InformationLigne &= Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format"))).Trim    ' 
                                                                    ReferenceArticle = Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format"))).Trim
                                                                Case "UOM_CODE"
                                                                    Dim EU_ENUMERE As String = Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format"))).Trim
                                                                    OleAdaptaterschemaEU_ENUMERE = New OleDbDataAdapter("SELECT EU_Enumere FROM dbo.F_DOCLIGNE WHERE (DO_Type = " & DO_Type & ") AND (DO_Piece = '" & Piece & "') AND (DL_Ligne = " & NumeroLigne1 & ")", OleSocieteConnect)
                                                                    OleSchemaDatasetEU_ENUMERE = New DataSet
                                                                    OleAdaptaterschemaEU_ENUMERE.Fill(OleSchemaDatasetEU_ENUMERE)
                                                                    OledatableSchemaEU_ENUMERE = OleSchemaDatasetEU_ENUMERE.Tables(0)
                                                                    If OledatableSchemaEU_ENUMERE.Rows.Count <> 0 Then
                                                                        InformationLigne &= OledatableSchemaEU_ENUMERE.Rows(0).Item("EU_Enumere")   ' 
                                                                    Else
                                                                        Dim OleAdaptaterschemaEU_ENUMERE2 As OleDbDataAdapter
                                                                        Dim OleSchemaDatasetEU_ENUMERE2 As DataSet
                                                                        Dim OledatableSchemaEU_ENUMERE2 As DataTable
                                                                        OleAdaptaterschemaEU_ENUMERE2 = New OleDbDataAdapter("SELECT  dbo.F_CONDITION.EC_Enumere FROM dbo.P_CONDITIONNEMENT RIGHT OUTER JOIN dbo.F_ARTICLE ON dbo.P_CONDITIONNEMENT.cbMarq = dbo.F_ARTICLE.AR_Condition LEFT OUTER JOIN dbo.F_CONDITION ON dbo.F_ARTICLE.AR_Ref = dbo.F_CONDITION.AR_Ref WHERE (dbo.F_ARTICLE.AR_Ref = '" & ReferenceArticle & "') AND (dbo.P_CONDITIONNEMENT.P_Conditionnement = '" & EU_ENUMERE & "') AND (dbo.F_CONDITION.CO_Principal = 1)", OleSocieteConnect)
                                                                        OleSchemaDatasetEU_ENUMERE2 = New DataSet
                                                                        OleAdaptaterschemaEU_ENUMERE2.Fill(OleSchemaDatasetEU_ENUMERE2)
                                                                        OledatableSchemaEU_ENUMERE2 = OleSchemaDatasetEU_ENUMERE2.Tables(0)
                                                                        If OledatableSchemaEU_ENUMERE2.Rows.Count <> 0 Then
                                                                            InformationLigne &= OledatableSchemaEU_ENUMERE2.Rows(0).Item("EC_Enumere")   '
                                                                        Else
                                                                            InformationLigne &= ""  '
                                                                        End If
                                                                    End If
                                                                Case Else
                                                                    InformationLigne &= Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format"))).Trim & ";"  ' 
                                                            End Select
                                                        End If
                                                    End If
                                                    Dim VirgulSuplementaire As String = ";;;;;"
                                                    Select Case Achat.Trim.ToUpper
                                                        Case "ACHAT"
                                                            If EstQuantiteVide = False Then
                                                                FichierCSO.WriteLine(Information & InformationLigne & VirgulSuplementaire)
                                                            End If
                                                            InformationLigne = ""
                                                        Case "VENTE"
                                                            If EstQuantiteVide = False Then
                                                                FichierCSO.WriteLine(Information & InformationLigne & VirgulSuplementaire)
                                                            End If
                                                            InformationLigne = ""
                                                        Case "TRANSFERT"
                                                            If DO_Type = "14" Then
                                                                FichierCSO.WriteLine(Information & InformationLigne & ";")
                                                            End If
                                                            InformationLigne = ""
                                                    End Select
                                                Else
                                                    If Tableau.Length = 2 Then
                                                        Dim Divisuer As Integer = Math.Pow(10, Tableau(1).Split(")")(0))
                                                        Dim Valeurs As Object = Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format")))
                                                        If OledatableSchemaLigne.Rows(LigneCols).Item("Cols").ToString.ToUpper = "LINE_NUMBER" Then
                                                            If Valeurs.ToString.Trim <> "" Then
                                                                NumeroLigne1 = CDbl(Valeurs / Divisuer)
                                                                InformationLigne &= CDbl(Valeurs / Divisuer).ToString.Trim & ";" ' 
                                                            Else
                                                                InformationLigne &= ";"
                                                            End If
                                                        ElseIf OledatableSchemaLigne.Rows(LigneCols).Item("Cols").ToString.ToUpper = "CUST_LINE_NUMBER" Then
                                                            If Valeurs.ToString.Trim <> "" Then
                                                                InformationLigne &= CDbl(Valeurs / Divisuer).ToString.Trim & ";" ' 
                                                                NumeroLigne2 = CDbl(Valeurs / Divisuer)
                                                            Else
                                                                InformationLigne &= ";"
                                                            End If
                                                        ElseIf OledatableSchemaLigne.Rows(LigneCols).Item("Cols").ToString.ToUpper = "SERVED_QUANTITY" Then
                                                            If Valeurs.ToString.Trim <> "" Then
                                                                If CDbl(Valeurs / Divisuer) <> 0 Then
                                                                    InformationLigne &= CDbl(Valeurs / Divisuer).ToString.Trim & ";" ' 
                                                                    EstQuantiteVide = False
                                                                Else
                                                                    EstQuantiteVide = True
                                                                End If
                                                            Else
                                                                InformationLigne &= ";"
                                                            End If
                                                        Else

                                                            If Valeurs.ToString.Trim <> "" Then
                                                                InformationLigne &= CDbl(Valeurs / Divisuer).ToString.Trim & ";" ' 
                                                            Else
                                                                InformationLigne &= ";"
                                                            End If
                                                        End If
                                                    Else
                                                        If OledatableSchemaLigne.Rows(LigneCols).Item("Format").ToString.Split("(")(0) = "D" Then
                                                            Dim Valeurs As Object = Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format")))
                                                            If Valeurs.ToString.Trim <> "" Then
                                                                Dim Year As String = Valeurs.ToString.Substring(0, 4)
                                                                Dim Mois As String = Valeurs.ToString.Substring(4, 2)
                                                                Dim Jours As String = Valeurs.ToString.Substring(6, 2)
                                                                Dim Heure As String = Valeurs.ToString.Substring(8, 2)
                                                                Dim Minute As String = Valeurs.ToString.Substring(10, 2)
                                                                Dim Seconde As String = Valeurs.ToString.Substring(12, 2)
                                                                Dim MyDate As String = Year & "-" & Mois & "-" & Jours & " " & Heure & ":" & Minute & ":" & Seconde
                                                                Dim MyNewDate As Date = RenvoieDateValide(MyDate, ComboDate.Text)
                                                                InformationLigne &= MyDate.Trim & ";" '
                                                            Else
                                                                InformationLigne &= Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format"))).Trim & ";" ' 
                                                            End If
                                                        ElseIf OledatableSchemaLigne.Rows(LigneCols).Item("Format").ToString = "N(10)" Then
                                                            Dim Valeur As String = Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format"))).Trim
                                                            If Valeur.ToString <> "" Then
                                                                InformationLigne &= Convert.ToDecimal(Valeur) & ";"
                                                            Else
                                                                InformationLigne &= ";"
                                                            End If
                                                        Else
                                                            Select Case OledatableSchemaLigne.Rows(LigneCols).Item("Cols").ToString.ToUpper
                                                                Case "PRODUCT_CODE"
                                                                    InformationLigne &= Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format"))).Trim & ";"   ' 
                                                                    ReferenceArticle = Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format"))).Trim
                                                                Case "UOM_CODE"
                                                                    Dim EU_ENUMERE As String = Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format"))).Trim
                                                                    OleAdaptaterschemaEU_ENUMERE = New OleDbDataAdapter("SELECT EU_Enumere FROM dbo.F_DOCLIGNE WHERE (DO_Type = " & DO_Type & ") AND (DO_Piece = '" & Piece & "') AND (DL_Ligne = " & NumeroLigne1 & ")", OleSocieteConnect)
                                                                    OleSchemaDatasetEU_ENUMERE = New DataSet
                                                                    OleAdaptaterschemaEU_ENUMERE.Fill(OleSchemaDatasetEU_ENUMERE)
                                                                    OledatableSchemaEU_ENUMERE = OleSchemaDatasetEU_ENUMERE.Tables(0)
                                                                    If OledatableSchemaEU_ENUMERE.Rows.Count <> 0 Then
                                                                        InformationLigne &= OledatableSchemaEU_ENUMERE.Rows(0).Item("EU_Enumere") & ";"  ' 
                                                                    Else
                                                                        Dim OleAdaptaterschemaEU_ENUMERE2 As OleDbDataAdapter
                                                                        Dim OleSchemaDatasetEU_ENUMERE2 As DataSet
                                                                        Dim OledatableSchemaEU_ENUMERE2 As DataTable

                                                                        OleAdaptaterschemaEU_ENUMERE2 = New OleDbDataAdapter("SELECT  dbo.F_CONDITION.EC_Enumere FROM dbo.P_CONDITIONNEMENT RIGHT OUTER JOIN dbo.F_ARTICLE ON dbo.P_CONDITIONNEMENT.cbMarq = dbo.F_ARTICLE.AR_Condition LEFT OUTER JOIN dbo.F_CONDITION ON dbo.F_ARTICLE.AR_Ref = dbo.F_CONDITION.AR_Ref WHERE (dbo.F_ARTICLE.AR_Ref = '" & ReferenceArticle & "') AND (dbo.P_CONDITIONNEMENT.P_Conditionnement = '" & EU_ENUMERE & "') AND (dbo.F_CONDITION.CO_Principal = 1)", OleSocieteConnect)
                                                                        OleSchemaDatasetEU_ENUMERE2 = New DataSet
                                                                        OleAdaptaterschemaEU_ENUMERE2.Fill(OleSchemaDatasetEU_ENUMERE2)
                                                                        OledatableSchemaEU_ENUMERE2 = OleSchemaDatasetEU_ENUMERE2.Tables(0)

                                                                        If OledatableSchemaEU_ENUMERE2.Rows.Count <> 0 Then
                                                                            InformationLigne &= OledatableSchemaEU_ENUMERE2.Rows(0).Item("EC_Enumere") & ";"  '
                                                                        Else
                                                                            InformationLigne &= ";"  '
                                                                        End If
                                                                    End If
                                                                Case Else
                                                                    InformationLigne &= Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format"))).Trim & ";"  ' 
                                                            End Select
                                                        End If
                                                    End If
                                                End If
                                            Next
                                            Detail_Count = Strings.Mid(Lignes, 445, 10)
                                            For j As Integer = 1 To Detail_Count 'ici je dois traite le detail de la ligne de la confirmation de la commande 
                                                Ligne = aRows(k1 + Parcourir)
                                                NbDetaiLigne += 1
                                                Dim Divisuer As Integer
                                                For LigneCols As Integer = 0 To OledatableSchemaDétailLigne.Rows.Count - 1
                                                    If LigneCols = OledatableSchemaDétailLigne.Rows.Count - 1 Then
                                                        Dim Tableau() As String = OledatableSchemaDétailLigne.Rows(LigneCols).Item("Format").ToString.Split("(")(1).Split(".")
                                                        Dim Valeurs As Object = Strings.Mid(Ligne, OledatableSchemaDétailLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaDétailLigne.Rows(LigneCols).Item("Format")))
                                                        Select Case OledatableSchemaDétailLigne.Rows(LigneCols).Item("Cols").ToString.Trim
                                                            Case "LOT_CODE"
                                                                InfosLot = " Lot " & Valeurs.ToString.Trim
                                                            Case "EXPIRATION_DATE"
                                                                If Valeurs.ToString.Trim <> "" Then
                                                                    Dim Year As String = Valeurs.ToString.Substring(0, 4)
                                                                    Dim Mois As String = Valeurs.ToString.Substring(4, 2)
                                                                    Dim Jours As String = Valeurs.ToString.Substring(6, 2)
                                                                    Dim Heure As String = Valeurs.ToString.Substring(8, 2)
                                                                    Dim Minute As String = Valeurs.ToString.Substring(10, 2)
                                                                    Dim Seconde As String = Valeurs.ToString.Substring(12, 2)
                                                                    Dim MyDate As String = Year & "-" & Mois & "-" & Jours & " " & Heure & ":" & Minute & ":" & Seconde
                                                                    Dim MyNewDate As Date = RenvoieDateValide(MyDate, ComboDate.Text)
                                                                    InfosDateExport = " DLUO " & MyNewDate
                                                                End If
                                                            Case "BEST_BEFORE_DATE"
                                                                If Valeurs.ToString.Trim <> "" Then
                                                                    Dim Year As String = Valeurs.ToString.Substring(0, 4)
                                                                    Dim Mois As String = Valeurs.ToString.Substring(4, 2)
                                                                    Dim Jours As String = Valeurs.ToString.Substring(6, 2)
                                                                    Dim Heure As String = Valeurs.ToString.Substring(8, 2)
                                                                    Dim Minute As String = Valeurs.ToString.Substring(10, 2)
                                                                    Dim Seconde As String = Valeurs.ToString.Substring(12, 2)
                                                                    Dim MyDate As String = Year & "-" & Mois & "-" & Jours & " " & Heure & ":" & Minute & ":" & Seconde
                                                                    Dim MyNewDate As Date = RenvoieDateValide(MyDate, ComboDate.Text)
                                                                    InfosDateExport = " DLUO " & MyNewDate
                                                                End If
                                                            Case "SHIPPED_QUANTITY"
                                                                Divisuer = Math.Pow(10, Tableau(1).Split(")")(0))
                                                                If Tableau.Length = 2 Then
                                                                    If Valeurs.ToString.Trim <> "" Then
                                                                        InfosQuantite = Valeurs / Divisuer
                                                                        'For ilineD As Integer = 0 To DataListeIntegrerDétailLigne.Rows.Count - 1
                                                                        '    If DataListeIntegrerDétailLigne.Rows(ilineD).Cells("Liens").Value = Piece And DataListeIntegrerDétailLigne.Rows(ilineD).Cells("RE").Value = ReferenceArticle Then
                                                                        '        InfosQuantite += CDbl(DataListeIntegrerDétailLigne.Rows(ilineD).Cells("SHIPPED_QUANTITY").Value)
                                                                        '    End If
                                                                        'Next
                                                                    End If
                                                                End If
                                                            Case "SOURCE_CONTAINER_NO"
                                                                SecondeChampLotCommentaire = Trim(Valeurs)
                                                        End Select
                                                        Dim Virgule As String = "" '
                                                        Dim AutreChamp1 As Date = CDate(InfosDateExport.Replace(" DLUO ", ""))
                                                        Dim AutreChamp2 As String = InfosLot.Replace(" Lot ", "")
                                                        Dim NewDate As String = Format(AutreChamp1, "yyyy MM dd").Replace(" ", "-")
                                                        Dim NewChaine As String = InfosLot & InfosDateExport & " Qte " & InfosQuantite & ";" & NewDate & ";" & SecondeChampLotCommentaire & ";" & AutreChamp2 & ";" & InfosQuantite ' "REF " & ReferenceArticle &
                                                        For LigneSecondaire As Integer = 0 To OledatableSchemaLigne.Rows.Count - 1
                                                            Virgule &= ";"
                                                        Next

                                                        Select Case Achat.Trim.ToUpper
                                                            Case "ACHAT"
                                                                If EstQuantiteVide = False Then
                                                                    FichierCSO.WriteLine(Information & Virgule & NewChaine)
                                                                End If
                                                            Case "VENTE"
                                                                If EstQuantiteVide = False Then
                                                                    FichierCSO.WriteLine(Information & Virgule & NewChaine)
                                                                End If
                                                            Case "TRANSFERT"
                                                                If DO_Type = "14" Then
                                                                    FichierCSO.WriteLine(Information & Virgule & NewChaine)
                                                                End If
                                                        End Select
                                                        InfosDateExport = ""
                                                        InfosLot = ""
                                                        InfosQuantite = 0
                                                        'If ReferenceArticle <> "" Then

                                                        '    'ReferenceArticle = ""
                                                        'End If
                                                    Else
                                                        Dim Tableau() As String = OledatableSchemaDétailLigne.Rows(LigneCols).Item("Format").ToString.Split("(")(1).Split(".")
                                                        Dim Valeurs As Object = Strings.Mid(Ligne, OledatableSchemaDétailLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaDétailLigne.Rows(LigneCols).Item("Format")))
                                                        Select Case OledatableSchemaDétailLigne.Rows(LigneCols).Item("Cols").ToString.Trim
                                                            Case "LOT_CODE"
                                                                InfosLot = " Lot " & Valeurs.ToString.Trim
                                                            Case "EXPIRATION_DATE"
                                                                If Valeurs.ToString.Trim <> "" Then
                                                                    Dim Year As String = Valeurs.ToString.Substring(0, 4)
                                                                    Dim Mois As String = Valeurs.ToString.Substring(4, 2)
                                                                    Dim Jours As String = Valeurs.ToString.Substring(6, 2)
                                                                    Dim Heure As String = Valeurs.ToString.Substring(8, 2)
                                                                    Dim Minute As String = Valeurs.ToString.Substring(10, 2)
                                                                    Dim Seconde As String = Valeurs.ToString.Substring(12, 2)
                                                                    Dim MyDate As String = Year & "-" & Mois & "-" & Jours & " " & Heure & ":" & Minute & ":" & Seconde
                                                                    Dim MyNewDate As Date = RenvoieDateValide(MyDate, ComboDate.Text)
                                                                    InfosDateExport = " DLUO " & MyNewDate
                                                                End If
                                                            Case "BEST_BEFORE_DATE"
                                                                If Valeurs.ToString.Trim <> "" Then
                                                                    Dim Year As String = Valeurs.ToString.Substring(0, 4)
                                                                    Dim Mois As String = Valeurs.ToString.Substring(4, 2)
                                                                    Dim Jours As String = Valeurs.ToString.Substring(6, 2)
                                                                    Dim Heure As String = Valeurs.ToString.Substring(8, 2)
                                                                    Dim Minute As String = Valeurs.ToString.Substring(10, 2)
                                                                    Dim Seconde As String = Valeurs.ToString.Substring(12, 2)
                                                                    Dim MyDate As String = Year & "-" & Mois & "-" & Jours & " " & Heure & ":" & Minute & ":" & Seconde
                                                                    Dim MyNewDate As Date = RenvoieDateValide(MyDate, ComboDate.Text)
                                                                    InfosDateExport = " DLUO " & MyNewDate
                                                                End If
                                                            Case "SHIPPED_QUANTITY"
                                                                Divisuer = Math.Pow(10, Tableau(1).Split(")")(0))
                                                                If Tableau.Length = 2 Then
                                                                    If Valeurs.ToString.Trim <> "" Then
                                                                        InfosQuantite = Valeurs / Divisuer
                                                                        'For ilineD As Integer = 0 To DataListeIntegrerDétailLigne.Rows.Count - 1
                                                                        '    If DataListeIntegrerDétailLigne.Rows(ilineD).Cells("Liens").Value = Piece And DataListeIntegrerDétailLigne.Rows(ilineD).Cells("RE").Value = ReferenceArticle Then
                                                                        '        InfosQuantite += CDbl(DataListeIntegrerDétailLigne.Rows(ilineD).Cells("SHIPPED_QUANTITY").Value)
                                                                        '    End If
                                                                        'Next
                                                                    End If
                                                                End If
                                                            Case "SOURCE_CONTAINER_NO"
                                                                SecondeChampLotCommentaire = Trim(Valeurs)
                                                        End Select
                                                    End If
                                                Next
                                                k1 += 1
                                            Next
                                            k = k + Detail_Count + 1
                                        End If
                                    Next
                                Else
                                    For Parcourir As Integer = 1 To Line_Count
                                        NbLigne += 1
                                        Dim Lignes As String = aRows(k + i)
                                        Detail_Count = Strings.Mid(Lignes, 445, 10)
                                        For j As Integer = 1 To Detail_Count 'ici je dois traite le detail de la ligne de la confirmation de la commande 
                                            k1 += 1
                                        Next
                                        k = k + Detail_Count + 1
                                    Next
                                    k1 += NbLigne + i + 1
                                    i = k - 1
                                    k = 1
                                    Continue For
                                End If
                            End If
                        End If
                        If (NbLigne + NbDetaiLigne + NbEntete) = aRows.Length Then
                            Information = ""
                            InformationLigne = ""
                            InfosQuantite = 0
                            InfosDateExport = ""
                            InfosLot = ""
                            ReferenceArticle = ""
                            If ChEncapsuler.Checked Then
                                FichierCSO.Close()
                            End If
                            infosExport.Text = "Export Terminer !"
                            Exit For
                        Else
                            i = k - 1 + i
                            k = 1
                            k1 = NbLigne + NbDetaiLigne + NbEntete + 1
                            Information = ""
                            InformationLigne = ""
                            InfosQuantite = 0
                            InfosDateExport = ""
                            InfosLot = ""
                            ReferenceArticle = ""
                        End If
                    Next
                    '''''''''''''''''''''''''''''''''''''''''''
                End If
            End If
        Catch ex As Exception
            MsgBox("Fonction Aperçu Erreur :" & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub
    Public Sub AperçuElement(ByVal Chemin As String, Optional ByVal EstExecution As String = "")
        Dim OleAdaptaterschema, OleAdaptaterschemaLigne, OleAdaptaterschemaDétailLigne As OleDbDataAdapter
        Dim OleSchemaDataset, OleSchemaDatasetLigne, OleSchemaDatasetDétailLigne As DataSet
        Dim OledatableSchema, OledatableSchemaLigne, OledatableSchemaDétailLigne As DataTable
        RegardeStatut = True
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
        Statut = False

        OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='CSO' AND ENTETE=true ORDER BY ORDRE", OleConnenection)
        OleSchemaDataset = New DataSet
        OleAdaptaterschema.Fill(OleSchemaDataset)
        OledatableSchema = OleSchemaDataset.Tables(0)

        OleAdaptaterschemaLigne = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='CSO' AND ENTETE=False AND Ligne=True ORDER BY ORDRE", OleConnenection)
        OleSchemaDatasetLigne = New DataSet
        OleAdaptaterschemaLigne.Fill(OleSchemaDatasetLigne)
        OledatableSchemaLigne = OleSchemaDatasetLigne.Tables(0)

        OleAdaptaterschemaDétailLigne = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='CSO' AND ENTETE=False AND Ligne=False ORDER BY ORDRE", OleConnenection)
        OleSchemaDatasetDétailLigne = New DataSet
        OleAdaptaterschemaDétailLigne.Fill(OleSchemaDatasetDétailLigne)
        OledatableSchemaDétailLigne = OleSchemaDatasetDétailLigne.Tables(0)

        If EstExecution <> "" Then
            Try
                DataListeIntegrer.Rows.Clear()
                DataListeIntegrerLigne.Rows.Clear()
                DataListeIntegrerDétailLigne.Rows.Clear()

                If OledatableSchema.Rows.Count <> 0 Then
                    Dim aRows() As String = Nothing
                    Dim Line_Count As Integer = 0
                    Dim Detail_Count As Integer = 0
                    Dim k As Integer = 1
                    Dim k1 As Integer = 1
                    Dim Cpteur As Integer = 0
                    Dim LigneQuantiteDemandé As Double
                    Dim LigneQuantiteLivre As Double = 0
                    Dim LigneCodeArt As String = ""
                    Dim CountLigne As Integer = 1
                    If GetArrayFile(Chemin, aRows) IsNot Nothing Then
                        aRows = GetArrayFile(Chemin, aRows)
                        For i As Integer = 0 To UBound(aRows)
                            Dim Ligne As String = aRows(i)
                            If GetNombreLigne(Ligne, 359, 10) <> 0 Then
                                If Not IsNumeric(Strings.Mid(Ligne, 1, 10)) Then
                                    Line_Count = Strings.Mid(Ligne, 359, 10)
                                    If Strings.Mid(Ligne, 61, 30).Trim <> "CANCELED" Then
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
                                        'Integration_Du_Fichier(Ligne)
                                        Integrer_Ecriture_Entete(Ligne)
                                        Dim Pathfichierjournal As String = Pathsfilejournal & "CSO_BL_" & Format(DateAndTime.Year(Now), "0000") & "" & Format(DateAndTime.Month(Now), "00") & "" & Format(DateAndTime.Day(Now), "00") & "_" & "" & Format(DateAndTime.Hour(Now), "00") & "_" & Format(DateAndTime.Minute(Now), "00") & "_" & Format(DateAndTime.Second(Now), "00") & ".txt"
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
                                                                infosExport.Text = "Connexion à la Société " & Trim(NomBaseCpta) & " : Echec"
                                                            End If
                                                        Else
                                                            RegardeStatut = False
                                                            ErreurJrn.WriteLine("Echec de Connexion à SQL de base de données :" & System.IO.Path.GetFileNameWithoutExtension(Trim(Cptadatatable.Rows(0).Item("Chemin1"))) & " Serveur : " & LireChaine(Trim(Cptadatatable.Rows(0).Item("Chemin1")), "CBASE", "ServeurSQL"))
                                                            infosExport.Text = "Echec de Connexion SQL à la Société " & Trim(NomBaseCpta) & " : Echec de traitement"
                                                        End If
                                                    Else
                                                        RegardeStatut = False
                                                        ErreurJrn.WriteLine("Chemin du fichier Comptable : " & Trim(Cptadatatable.Rows(0).Item("Chemin1")) & " inexistant")
                                                        infosExport.Text = "Chemin du fichier Comptable : " & Trim(Cptadatatable.Rows(0).Item("Chemin1")) & " inexistant"
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
                                            infosExport.Text = "Le Répertoire Journal :" & Pathsfilejournal & " n'est pas valide "
                                        End If
                                        If RegardeStatut = False Then
                                            Exit Sub
                                        End If
                                        If ExisteSOCIETE_ROUTAGE(CodeSociete) = True Then
                                            If ExisteTiers(EnteteCodeTiers, EnteteCodeFournisseur, EnteteCodeTransfertDepot) = True Then
                                                If ExistePiece(EntetePieceInterne, EnteteTyPeDocument) Then
                                                    If ExisteDepot(IDDepotEntete, EntetePieceInterne) = True Then
                                                        If ExisteModeExpedition(EnteteDoExpedition, EntetePieceInterne) Then
                                                            StatutCreationEnteteDoc = True
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                        For LigneCols As Integer = 0 To OledatableSchema.Rows.Count - 4
                                            Dim Chaine As String = Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format")))
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
                                                    DataListeIntegrer.Rows(iline - 1).Cells(LigneCols).Value = MyDate 'Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format"))) ' 
                                                Else
                                                    DataListeIntegrer.Rows(iline - 1).Cells(LigneCols).Value = Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format"))).Trim  ' 
                                                End If
                                            Else
                                                DataListeIntegrer.Rows(iline - 1).Cells(LigneCols).Value = Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format"))).Trim  ' 
                                            End If
                                        Next
                                        'ici je dois traité les lignes de la confirmation de commande
                                        Dim Parcourir As Integer = 1
                                        For Parcourir = 1 To Line_Count
                                            NbLigne += 1
                                            Dim Lignes As String = aRows(k + i)
                                            Integrer_Ecriture_Ligne(Lignes, "F_DOCLIGNE")
                                            LigneCodeArt = Strings.Mid(Lignes, 61, 50)
                                            LigneQuantiteDemandé = CDbl(Strings.Mid(Lignes, 161, 12) / Math.Pow(10, 5))
                                            If RenvoieStockNegatif() = False Then
                                                If VerificationStockDispoDepot(IDDepotEntete, LigneCodeArt, LigneQuantiteDemandé, EntetePieceInterne) Then
                                                    StatutCreationLigneDoc = True
                                                End If
                                            End If
                                            If DataListeIntegrerLigne.RowCount = 0 Then
                                                iLigne += 1
                                                DataListeIntegrerLigne.RowCount = iLigne
                                            Else
                                                If DataListeIntegrerLigne.Rows(DataListeIntegrerLigne.RowCount - 1).Cells(0).Value <> "" Then
                                                    If NbEntete > 1 And Statut = False Then
                                                        DataListeIntegrerLigne.Rows.Add()
                                                        iLigne += 2
                                                        DataListeIntegrerLigne.RowCount = iLigne
                                                        Statut = True
                                                    Else
                                                        iLigne += 1
                                                        DataListeIntegrerLigne.RowCount = iLigne
                                                    End If
                                                Else
                                                    iLigne = DataListeIntegrerLigne.RowCount
                                                    iLigne += 1
                                                    DataListeIntegrerLigne.RowCount = iLigne
                                                End If
                                            End If

                                            For LigneCols As Integer = 0 To DataListeIntegrerLigne.ColumnCount - 1 ' OledatableSchemaLigne.Rows.Count - 2
                                                If LigneCols = DataListeIntegrerLigne.ColumnCount - 1 Then
                                                    DataListeIntegrerLigne.Rows(iLigne - 1).Cells(LigneCols).Value = DataListeIntegrer.Rows(iline - 1).Cells("SORDER_CODE").Value.ToString.Trim
                                                Else
                                                    Dim Tableau() As String = OledatableSchemaLigne.Rows(LigneCols).Item("Format").ToString.Split("(")(1).Split(".")
                                                    If Tableau.Length = 2 Then
                                                        Dim Divisuer As Integer = Math.Pow(10, Tableau(1).Split(")")(0))
                                                        Dim Valeurs As Object = Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format")))
                                                        If Valeurs.ToString.Trim <> "" Then
                                                            DataListeIntegrerLigne.Rows(iLigne - 1).Cells(LigneCols).Value = CDbl(Valeurs / Divisuer).ToString.Trim  ' 
                                                        Else
                                                            DataListeIntegrerLigne.Rows(iLigne - 1).Cells(LigneCols).Value = 0
                                                        End If
                                                    Else
                                                        If OledatableSchemaLigne.Rows(LigneCols).Item("Format").ToString.Split("(")(0) = "D" Then
                                                            Dim Valeurs As Object = Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format")))
                                                            If Valeurs.ToString.Trim <> "" Then
                                                                Dim Year As String = Valeurs.ToString.Substring(0, 4)
                                                                Dim Mois As String = Valeurs.ToString.Substring(4, 2)
                                                                Dim Jours As String = Valeurs.ToString.Substring(6, 2)
                                                                Dim Heure As String = Valeurs.ToString.Substring(8, 2)
                                                                Dim Minute As String = Valeurs.ToString.Substring(10, 2)
                                                                Dim Seconde As String = Valeurs.ToString.Substring(12, 2)
                                                                Dim MyDate As String = Year & "-" & Mois & "-" & Jours & " " & Heure & ":" & Minute & ":" & Seconde
                                                                Dim MyNewDate As Date = RenvoieDateValide(MyDate, ComboDate.Text)
                                                                DataListeIntegrerLigne.Rows(iLigne - 1).Cells(LigneCols).Value = MyDate.Trim 'Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format"))) ' 
                                                            Else
                                                                DataListeIntegrerLigne.Rows(iLigne - 1).Cells(LigneCols).Value = Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format"))).Trim  ' 
                                                            End If
                                                        Else
                                                            DataListeIntegrerLigne.Rows(iLigne - 1).Cells(LigneCols).Value = Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format"))).Trim  ' 
                                                        End If
                                                    End If
                                                End If
                                            Next
                                            Detail_Count = Strings.Mid(Lignes, 445, 10)
                                            For j As Integer = 1 To Detail_Count 'ici je dois traite le detail de la ligne de la confirmation de la commande 
                                                Ligne = aRows(k1 + Parcourir)
                                                NbDetaiLigne += 1
                                                If DataListeIntegrerDétailLigne.RowCount = 0 Then
                                                    iDetaiLigne += 1
                                                    DataListeIntegrerDétailLigne.RowCount = iDetaiLigne
                                                Else
                                                    If DataListeIntegrerDétailLigne.Rows(DataListeIntegrerDétailLigne.RowCount - 1).Cells(0).Value <> "" Then
                                                        If NbEntete > 1 And Statut = False Then
                                                            DataListeIntegrerDétailLigne.Rows.Add()
                                                            iDetaiLigne += 2
                                                            DataListeIntegrerDétailLigne.RowCount = iDetaiLigne
                                                            Statut = True
                                                        Else
                                                            iDetaiLigne += 1
                                                            DataListeIntegrerDétailLigne.RowCount = iDetaiLigne
                                                        End If
                                                    Else
                                                        iDetaiLigne = DataListeIntegrerDétailLigne.RowCount
                                                        iDetaiLigne += 1
                                                        DataListeIntegrerDétailLigne.RowCount = iDetaiLigne
                                                    End If
                                                End If
                                                For LigneCols As Integer = 0 To DataListeIntegrerDétailLigne.ColumnCount - 1  'OledatableSchemaLigne.Rows.Count - 2
                                                    If LigneCols = DataListeIntegrerDétailLigne.ColumnCount - 1 Then
                                                        DataListeIntegrerDétailLigne.Rows(iDetaiLigne - 1).Cells(LigneCols).Value = DataListeIntegrerLigne.Rows(iLigne - 1).Cells("PRODUCT_CODE").Value.ToString.Trim
                                                    Else
                                                        Dim Tableau() As String = OledatableSchemaDétailLigne.Rows(LigneCols).Item("Format").ToString.Split("(")(1).Split(".")
                                                        If Tableau.Length = 2 Then
                                                            Dim Divisuer As Integer = Math.Pow(10, Tableau(1).Split(")")(0))
                                                            Dim Valeurs As Object = Strings.Mid(Ligne, OledatableSchemaDétailLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaDétailLigne.Rows(LigneCols).Item("Format")))
                                                            If Valeurs.ToString.Trim <> "" Then
                                                                DataListeIntegrerDétailLigne.Rows(iDetaiLigne - 1).Cells(LigneCols).Value = CDbl(Valeurs / Divisuer).ToString.Trim  ' 
                                                            Else
                                                                DataListeIntegrerDétailLigne.Rows(iDetaiLigne - 1).Cells(LigneCols).Value = 0
                                                            End If
                                                        Else
                                                            If OledatableSchemaDétailLigne.Rows(LigneCols).Item("Format").ToString.Split("(")(0) = "D" Then
                                                                Dim Valeurs As Object = Strings.Mid(Ligne, OledatableSchemaDétailLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaDétailLigne.Rows(LigneCols).Item("Format"))).Trim
                                                                If Valeurs.ToString.Trim <> "" Then
                                                                    Dim Year As String = Valeurs.ToString.Substring(0, 4)
                                                                    Dim Mois As String = Valeurs.ToString.Substring(4, 2)
                                                                    Dim Jours As String = Valeurs.ToString.Substring(6, 2)
                                                                    Dim Heure As String = Valeurs.ToString.Substring(8, 2)
                                                                    Dim Minute As String = Valeurs.ToString.Substring(10, 2)
                                                                    Dim Seconde As String = Valeurs.ToString.Substring(12, 2)
                                                                    Dim MyDate As String = Year & "-" & Mois & "-" & Jours & " " & Heure & ":" & Minute & ":" & Seconde
                                                                    DataListeIntegrerDétailLigne.Rows(iDetaiLigne - 1).Cells(LigneCols).Value = MyDate
                                                                Else
                                                                    DataListeIntegrerDétailLigne.Rows(iDetaiLigne - 1).Cells(LigneCols).Value = Strings.Mid(Ligne, OledatableSchemaDétailLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaDétailLigne.Rows(LigneCols).Item("Format"))).Trim  ' 
                                                                End If
                                                            Else
                                                                DataListeIntegrerDétailLigne.Rows(iDetaiLigne - 1).Cells(LigneCols).Value = Strings.Mid(Ligne, OledatableSchemaDétailLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaDétailLigne.Rows(LigneCols).Item("Format"))).Trim  ' 
                                                            End If
                                                        End If
                                                    End If
                                                Next
                                                k1 += 1
                                            Next
                                            k = k + Detail_Count + 1
                                        Next
                                    End If
                                End If
                            End If
                            If (NbLigne + NbDetaiLigne + NbEntete) = aRows.Length Then
                                If CountChecked > 1 Then
                                    DataListeIntegrer.Rows.Add()
                                    DataListeIntegrerLigne.Rows.Add()
                                    If NbDetaiLigne <> 0 Then
                                        DataListeIntegrerDétailLigne.Rows.Add()
                                    End If
                                    CountChecked -= 1
                                End If
                                If StatutCreationEnteteDoc = True And ExisteLecture = True Then 'And StatutCreationLigneDoc = True
                                    Traitement_Integration(Chemin)
                                Else
                                    infosExport.Text = "Verification Integration <[Consulter le fichier Journal]> Ctrl + J."
                                    Statut = False
                                    RegardeStatut = False
                                End If
                                Exit For
                            Else
                                i = k - 1 + i
                                k = 1
                                k1 = NbLigne + NbDetaiLigne + NbEntete + 1
                                If NbDetaiLigne <> 0 Then
                                    DataListeIntegrerDétailLigne.Rows.Add()
                                End If
                                If StatutCreationEnteteDoc = True Then 'And StatutCreationLigneDoc = True
                                    Dim list = 2
                                    'Creation_Entete_Document(EnteteTyPeDocument)
                                    Statut = False
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
                    Dim Parcourir As Integer = 1
                    Dim CompteurElseLigne As Integer = 0
                    Dim CompteurElseSsLigne As Integer = 0
                    Dim CountLigne As Integer = 1
                    If GetArrayFile(Chemin, aRows) IsNot Nothing Then
                        aRows = GetArrayFile(Chemin, aRows)
                        For i As Integer = 0 To UBound(aRows)
                            Dim Ligne As String = aRows(i)
                            If GetNombreLigne(Ligne, 359, 10) <> 0 Then
                                If Not IsNumeric(Strings.Mid(Ligne, 1, 10)) Then
                                    Line_Count = Strings.Mid(Ligne, 359, 10)
                                    If Strings.Mid(Ligne, 61, 30).Trim <> "CANCELED" Then
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
                                        For LigneCols As Integer = 0 To OledatableSchema.Rows.Count - 4
                                            Dim Chaine As String = Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format")))
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
                                                    DataListeIntegrer.Rows(iline - 1).Cells(LigneCols).Value = MyDate 'Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format"))) ' 
                                                Else
                                                    DataListeIntegrer.Rows(iline - 1).Cells(LigneCols).Value = Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format"))).Trim  ' 
                                                End If
                                            Else
                                                DataListeIntegrer.Rows(iline - 1).Cells(LigneCols).Value = Strings.Mid(Ligne, OledatableSchema.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchema.Rows(LigneCols).Item("Format"))).Trim  ' 
                                            End If
                                        Next
                                        'ici je dois traité les lignes de la confirmation de commande

                                        For Parcourir = 1 To Line_Count
                                            NbLigne += 1
                                            Dim Lignes As String = aRows(k + i)
                                            If DataListeIntegrerLigne.RowCount = 0 Then
                                                iLigne += 1
                                                DataListeIntegrerLigne.RowCount = iLigne
                                            Else
                                                If DataListeIntegrerLigne.Rows(DataListeIntegrerLigne.RowCount - 1).Cells(0).Value <> "" Then
                                                    If NbEntete > 1 And Statut = False Then
                                                        DataListeIntegrerLigne.Rows.Add()
                                                        iLigne += 2
                                                        DataListeIntegrerLigne.RowCount = iLigne
                                                        Statut = True
                                                    Else
                                                        iLigne += 1
                                                        DataListeIntegrerLigne.RowCount = iLigne
                                                    End If
                                                Else
                                                    iLigne = DataListeIntegrerLigne.RowCount
                                                    iLigne += 1
                                                    DataListeIntegrerLigne.RowCount = iLigne
                                                End If
                                            End If
                                            For LigneCols As Integer = 0 To DataListeIntegrerLigne.ColumnCount - 1 ' OledatableSchemaLigne.Rows.Count - 2
                                                If LigneCols = DataListeIntegrerLigne.ColumnCount - 1 Then
                                                    DataListeIntegrerLigne.Rows(iLigne - 1).Cells(LigneCols).Value = DataListeIntegrer.Rows(iline - 1).Cells("SORDER_CODE").Value.ToString.Trim
                                                Else
                                                    Dim Tableau() As String = OledatableSchemaLigne.Rows(LigneCols).Item("Format").ToString.Split("(")(1).Split(".")
                                                    If Tableau.Length = 2 Then
                                                        Dim Divisuer As Integer = Math.Pow(10, Tableau(1).Split(")")(0))
                                                        Dim Valeurs As Object = Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format")))
                                                        If Valeurs.ToString.Trim <> "" Then
                                                            DataListeIntegrerLigne.Rows(iLigne - 1).Cells(LigneCols).Value = CDbl(Valeurs / Divisuer).ToString.Trim  ' 
                                                        Else
                                                            DataListeIntegrerLigne.Rows(iLigne - 1).Cells(LigneCols).Value = 0
                                                        End If
                                                    Else
                                                        If OledatableSchemaLigne.Rows(LigneCols).Item("Format").ToString.Split("(")(0) = "D" Then
                                                            Dim Valeurs As Object = Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format")))
                                                            If Valeurs.ToString.Trim <> "" Then
                                                                Dim Year As String = Valeurs.ToString.Substring(0, 4)
                                                                Dim Mois As String = Valeurs.ToString.Substring(4, 2)
                                                                Dim Jours As String = Valeurs.ToString.Substring(6, 2)
                                                                Dim Heure As String = Valeurs.ToString.Substring(8, 2)
                                                                Dim Minute As String = Valeurs.ToString.Substring(10, 2)
                                                                Dim Seconde As String = Valeurs.ToString.Substring(12, 2)
                                                                Dim MyDate As String = Year & "-" & Mois & "-" & Jours & " " & Heure & ":" & Minute & ":" & Seconde
                                                                Dim MyNewDate As Date = RenvoieDateValide(MyDate, ComboDate.Text)
                                                                DataListeIntegrerLigne.Rows(iLigne - 1).Cells(LigneCols).Value = MyDate.Trim 'Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format"))) ' 
                                                            Else
                                                                DataListeIntegrerLigne.Rows(iLigne - 1).Cells(LigneCols).Value = Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format"))).Trim  ' 
                                                            End If
                                                        Else
                                                            DataListeIntegrerLigne.Rows(iLigne - 1).Cells(LigneCols).Value = Strings.Mid(Lignes, OledatableSchemaLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaLigne.Rows(LigneCols).Item("Format"))).Trim  ' 
                                                        End If
                                                    End If
                                                End If
                                            Next
                                            Detail_Count = Strings.Mid(Lignes, 445, 10)
                                            For j As Integer = 1 To Detail_Count 'ici je dois traite le detail de la ligne de la confirmation de la commande 
                                                Ligne = aRows(k1 + Parcourir)
                                                NbDetaiLigne += 1
                                                If DataListeIntegrerDétailLigne.RowCount = 0 Then
                                                    iDetaiLigne += 1
                                                    DataListeIntegrerDétailLigne.RowCount = iDetaiLigne
                                                Else
                                                    If DataListeIntegrerDétailLigne.Rows(DataListeIntegrerDétailLigne.RowCount - 1).Cells(0).Value <> "" Then
                                                        If NbEntete > 1 And Statut = False Then
                                                            DataListeIntegrerDétailLigne.Rows.Add()
                                                            iDetaiLigne += 2
                                                            DataListeIntegrerDétailLigne.RowCount = iDetaiLigne
                                                            Statut = True
                                                        Else
                                                            iDetaiLigne += 1
                                                            DataListeIntegrerDétailLigne.RowCount = iDetaiLigne
                                                        End If
                                                    Else
                                                        iDetaiLigne = DataListeIntegrerDétailLigne.RowCount
                                                        iDetaiLigne += 1
                                                        DataListeIntegrerDétailLigne.RowCount = iDetaiLigne
                                                    End If
                                                End If
                                                For LigneCols As Integer = 0 To DataListeIntegrerDétailLigne.ColumnCount - 2  'OledatableSchemaLigne.Rows.Count - 2PRODUCT_CODE
                                                    If LigneCols = DataListeIntegrerDétailLigne.ColumnCount - 2 Then
                                                        DataListeIntegrerDétailLigne.Rows(iDetaiLigne - 1).Cells(LigneCols).Value = DataListeIntegrerLigne.Rows(iLigne - 1).Cells("lien").Value.ToString.Trim
                                                        DataListeIntegrerDétailLigne.Rows(iDetaiLigne - 1).Cells(LigneCols + 1).Value = DataListeIntegrerLigne.Rows(iLigne - 1).Cells("PRODUCT_CODE").Value.ToString.Trim
                                                    Else
                                                        Dim Tableau() As String = OledatableSchemaDétailLigne.Rows(LigneCols).Item("Format").ToString.Split("(")(1).Split(".")
                                                        If Tableau.Length = 2 Then
                                                            Dim Divisuer As Integer = Math.Pow(10, Tableau(1).Split(")")(0))
                                                            Dim Valeurs As Object = Strings.Mid(Ligne, OledatableSchemaDétailLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaDétailLigne.Rows(LigneCols).Item("Format")))
                                                            If Valeurs.ToString.Trim <> "" Then
                                                                DataListeIntegrerDétailLigne.Rows(iDetaiLigne - 1).Cells(LigneCols).Value = CDbl(Valeurs / Divisuer).ToString.Trim  ' 
                                                            Else
                                                                DataListeIntegrerDétailLigne.Rows(iDetaiLigne - 1).Cells(LigneCols).Value = 0
                                                            End If
                                                        Else
                                                            If OledatableSchemaDétailLigne.Rows(LigneCols).Item("Format").ToString.Split("(")(0) = "D" Then
                                                                Dim Valeurs As Object = Strings.Mid(Ligne, OledatableSchemaDétailLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaDétailLigne.Rows(LigneCols).Item("Format"))).Trim
                                                                If Valeurs.ToString.Trim <> "" Then
                                                                    Dim Year As String = Valeurs.ToString.Substring(0, 4)
                                                                    Dim Mois As String = Valeurs.ToString.Substring(4, 2)
                                                                    Dim Jours As String = Valeurs.ToString.Substring(6, 2)
                                                                    Dim Heure As String = "" ' Valeurs.ToString.Substring(8, 2)
                                                                    Dim Minute As String = Valeurs.ToString.Substring(10, 2)
                                                                    Dim Seconde As String = Valeurs.ToString.Substring(12, 2)
                                                                    Dim MyDate As String = Year & "-" & Mois & "-" & Jours & " " & Heure & ":" & Minute & ":" & Seconde
                                                                    DataListeIntegrerDétailLigne.Rows(iDetaiLigne - 1).Cells(LigneCols).Value = MyDate
                                                                Else
                                                                    DataListeIntegrerDétailLigne.Rows(iDetaiLigne - 1).Cells(LigneCols).Value = Strings.Mid(Ligne, OledatableSchemaDétailLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaDétailLigne.Rows(LigneCols).Item("Format"))).Trim  ' 
                                                                End If
                                                            Else
                                                                DataListeIntegrerDétailLigne.Rows(iDetaiLigne - 1).Cells(LigneCols).Value = Strings.Mid(Ligne, OledatableSchemaDétailLigne.Rows(LigneCols).Item("PositionG"), GetLongueurChaine(OledatableSchemaDétailLigne.Rows(LigneCols).Item("Format"))).Trim  ' 
                                                            End If
                                                        End If
                                                    End If
                                                Next
                                                k1 += 1
                                            Next
                                            k = k + Detail_Count + 1
                                        Next
                                    Else
                                        'Faire avance le pointeur d'entete et ligne et ss-ligne ici
                                        For Parcourir = 1 To Line_Count
                                            NbLigne += 1
                                            Dim Lignes As String = aRows(k + i)
                                            Detail_Count = Strings.Mid(Lignes, 445, 10)
                                            For j As Integer = 1 To Detail_Count 'ici je dois traite le detail de la ligne de la confirmation de la commande 
                                                k1 += 1
                                            Next
                                            k = k + Detail_Count + 1
                                        Next
                                        k1 += NbLigne + i + 1
                                        i = k - 1
                                        k = 1
                                        Continue For
                                    End If
                                End If
                            End If
                            If (NbLigne + NbDetaiLigne + NbEntete) = aRows.Length Then
                                If CountChecked > 1 Then
                                    DataListeIntegrer.Rows.Add()
                                    DataListeIntegrerLigne.Rows.Add()
                                    If NbDetaiLigne <> 0 Then
                                        DataListeIntegrerDétailLigne.Rows.Add()
                                    End If
                                    CountChecked -= 1
                                End If
                                Exit For
                            Else
                                i = k - 1 + i
                                k = 1
                                k1 = NbLigne + NbDetaiLigne + NbEntete + 1
                                If NbDetaiLigne <> 0 Then
                                    DataListeIntegrerDétailLigne.Rows.Add()
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
            OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='CSO' AND Entete=" & Entete & " AND InfosLibre=" & InfosLibre & " AND Ligne=" & Ligne, OleConnenection)
            OleSchemaDataset = New DataSet
            OleAdaptaterschema.Fill(OleSchemaDataset)
            OledatableSchema = OleSchemaDataset.Tables(0)
            If OledatableSchema.Rows.Count <> 0 Then
                Dim Pathfichierjournal As String = Pathsfilejournal & "LOGIMP_INFOSLIBRE_CSO_" & Format(DateAndTime.Year(Now), "0000") & "" & Format(DateAndTime.Month(Now), "00") & "" & Format(DateAndTime.Day(Now), "00") & "_" & "" & Format(DateAndTime.Hour(Now), "00") & "_" & Format(DateAndTime.Minute(Now), "00") & "_" & Format(DateAndTime.Second(Now), "00") & ".txt"
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
            infoListe = New List(Of Integer)
            infoLigne = New List(Of Integer)

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
                    infoLigne.Add(InfosLibre)
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
        End Try
    End Function
    Public Sub OledbInitialiseur()
        Try
            OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='CSO' AND ENTETE=true ORDER BY ORDRE", OleConnenection)
            OleSchemaDataset = New DataSet
            OleAdaptaterschema.Fill(OleSchemaDataset)
            OledatableSchema = OleSchemaDataset.Tables(0)

            OleAdaptaterschemaLigne = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='CSO' AND ENTETE=False AND Ligne=True ORDER BY ORDRE", OleConnenection)
            OleSchemaDatasetLigne = New DataSet
            OleAdaptaterschemaLigne.Fill(OleSchemaDatasetLigne)
            OledatableSchemaLigne = OleSchemaDatasetLigne.Tables(0)

            OleAdaptaterschemaDétailLigne = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='CSO' AND ENTETE=False AND Ligne=False ORDER BY ORDRE", OleConnenection)
            OleSchemaDatasetDétailLigne = New DataSet
            OleAdaptaterschemaDétailLigne.Fill(OleSchemaDatasetDétailLigne)
            OledatableSchemaDétailLigne = OleSchemaDatasetDétailLigne.Tables(0)

        Catch ex As Exception
        End Try
    End Sub
    Public Function EcritureFlux(ByVal Chemin As Object) As Boolean
        Dim iline As Integer = 0
        Dim iLigne As Integer = 0
        Dim iDetaiLigne As Integer = 0
        Dim NbDetaiLigne As Integer = 0
        Dim NbLigne As Integer = 0
        Dim NbEntete As Integer = 0
        Dim Relation As String = ""
        Dim Statut As Boolean = False
        Try
            If Chemin <> "" Then
                Dim aRows() As String = Nothing
                Dim Line_Count As Integer = 0
                Dim Detail_Count As Integer = 0
                Dim k As Integer = 1
                Dim k1 As Integer = 1
                Dim Cpteur As Integer = 0
                'Dim TCpteur As Integer = 0
                Dim CountLigne As Integer = 1
                If GetArrayFile(Chemin, aRows) IsNot Nothing Then
                    aRows = GetArrayFile(Chemin, aRows)
                    For i As Integer = 0 To UBound(aRows)
                        Dim Ligne As String = aRows(i)
                        If GetNombreLigne(Ligne, 359, 10) <> 0 Then
                            If Not IsNumeric(Strings.Mid(Ligne, 1, 10)) Then
                                Line_Count = Strings.Mid(Ligne, 359, 10)
                                If Strings.Mid(Ligne, 61, 30).Trim <> "CANCELED" Then
                                    NbEntete += 1
                                    Statut = False
                                    'ici je dois traite l'entete de la confirmation de commande 
                                    'CreationEnteteDocument(Ligne, OledatableSchema)
                                    Dim Parcourir As Integer = 1
                                    For Parcourir = 1 To Line_Count
                                        NbLigne += 1
                                        Dim Lignes As String = aRows(k + i)
                                        'ici je dois traité les lignes de la confirmation de commande
                                        Detail_Count = Strings.Mid(Lignes, 445, 10)
                                        For j As Integer = 1 To Detail_Count 'ici je dois traite le detail de la ligne de la confirmation de la commande 
                                            Ligne = aRows(k1 + Parcourir)
                                            NbDetaiLigne += 1
                                            k1 += 1
                                        Next
                                        k = k + Detail_Count + 1
                                    Next
                                End If
                            End If
                        End If
                        If (NbLigne + NbDetaiLigne + NbEntete) = aRows.Length Then
                            Exit For
                        Else
                            i = k - 1 + i
                            k = 1
                            k1 = NbLigne + NbDetaiLigne + NbEntete + 1
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox("Fonction Aperçu Erreur :" & ex.Message)
        End Try
    End Function
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
            .ShowsForm = "1"
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
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_integrer.Click
        ExisteLecture = True
        OledbInitialiseur()
        vidage()
        infosExport.Text = ""
        IfrowErreur = 0
        Dim i As Integer
        Try
            CountChecked = IsChecked()
            If IsChecked() Then
                If RbtG1.Checked Then
                    EstInfosLibre(False, True, True)
                ElseIf RbtG2.Checked Then
                    EstInfosLibre(False, False, True)
                End If
                If NbInfosLibre = NbInfosLibreVue Then 'NbInfosLibreVue
                    For i = 0 To Datagridaffiche.RowCount - 1
                        If Datagridaffiche.Rows(i).Cells("C6").Value = True Then
                            IfrowErreur = i
                            MonFichier = Datagridaffiche.Rows(i).Cells("C8").Value
                            AperçuElement(Datagridaffiche.Rows(i).Cells("C8").Value, "Execution")
                            If RegardeStatut = True And ExisteLectures = True Then
                                Datagridaffiche.Rows(i).Cells("C7").Value = My.Resources.accepter
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
                MessageBox.Show("Un choix de traitement doit être fait Merci de faire votre choix", "Infos Choix Traitement", MessageBoxButtons.OK, MessageBoxIcon.Information)
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
        DataListeIntegrerLigne.Rows.Clear()
        DataListeIntegrerDétailLigne.Rows.Clear()
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
    Public Sub BtnListe_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnListe.Click
        Try
            LirefichierConfig()
            OuvreLaListedeFichier(PathsfileExport)
            Connected()
        Catch ex As Exception
        End Try
    End Sub 'liste des fichier à traité
    Public Sub Frm_FluxEntrantCritére_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            Me.WindowState = FormWindowState.Maximized
            ComboDate.SelectedIndex = 4
            CmbStatut.SelectedIndex = 3
            Connected()
            BackgroundWorker4.RunWorkerAsync()
        Catch ex As Exception
        End Try
        Try
            BtnListe_Click(sender, e)
        Catch ex As Exception
        End Try
    End Sub
    Public ExisteLectures As Boolean = True
    Public Function ExistePiece(ByVal EntetePieceInterne As String, ByVal TypePiece As String) As Boolean
        ExisteLectures = True
        If ChkPieceAuto.Checked = False Then
            Select Case TypePiece
                Case "3"
                    If BaseCial.FactoryDocumentVente.ExistPiece(DocumentType.DocumentTypeVenteLivraison, EntetePieceInterne.Trim) = True Then
                        ExisteLectures = False
                        ErreurJrn.WriteLine("Bon de Livraison N° : " & EntetePieceInterne & " Existe déja ")
                    End If
                Case "14"
                    If BaseCial.FactoryDocumentAchat.ExistPiece(DocumentType.DocumentTypeAchatReprise, EntetePieceInterne) = True Then
                        ExisteLectures = False
                        ErreurJrn.WriteLine("Bon de Retour N° : " & EntetePieceInterne & " Existe déja ")
                    End If
                Case "23"
                    If BaseCial.FactoryDocumentStock.ExistPiece(DocumentType.DocumentTypeStockVirement, EntetePieceInterne) = True Then
                        ExisteLectures = False
                        ErreurJrn.WriteLine("Transfert de Dépot à dépot - N°Pièce du Fichier : " & EntetePieceInterne & " Existe déja ")
                    End If
            End Select
        End If
        ExistePiece = ExisteLectures
    End Function
    Public Function ExisteTiers(ByVal Client As String, ByVal Fournisseurs As String, ByVal TransfertDepot As String) As Boolean
        Dim Etat As Boolean = True
        If Client.Trim <> "" Then
            If BaseCpta.FactoryClient.ExistNumero(Trim(Client)) = False Then
                ErreurJrn.WriteLine("Le Client [" & Client.Trim & "] n'existe pas dans Sage et dans le parametrage ")
                Etat = False
            Else
                Etat = True
                EnteteTyPeDocument = 3
            End If
        ElseIf Fournisseurs.Trim <> "" Then
            If BaseCpta.FactoryFournisseur.ExistNumero(Trim(Fournisseurs.Trim)) = False Then
                ErreurJrn.WriteLine("Le Fournisseur [" & Fournisseurs.Trim & "] n'existe pas dans Sage et dans le parametrage ")
                Etat = False
            Else
                Etat = True
                EnteteTyPeDocument = 14
            End If
        ElseIf TransfertDepot.Trim <> "" Then
            EnteteTyPeDocument = 23
            Etat = True
        End If
        ExisteTiers = Etat
    End Function
    Public Function ExisteModeExpedition(ByVal EnteteDoExpedition As String, ByVal EntetePieceInterne As String) As Boolean
        If IsDBNull(EnteteDoExpedition) = False Then
            If EnteteDoExpedition <> Nothing Then
                If BaseCial.FactoryExpedition.ExistIntitule(Trim(EnteteDoExpedition)) = False Then
                    ErreurJrn.WriteLine("Le Mode d' Expedition [" & EnteteDoExpedition & "] N°Pièce du Fichier : " & EntetePieceInterne & " m'existe dans Sage ")
                    Return False
                Else
                    Return True
                End If
            Else
                Return True
            End If
        Else
            ExisteModeExpedition = True
        End If
    End Function
    Public Function ExisteCondition(ByVal EnteteConditiondeLivraison As String, ByVal EntetePieceInterne As String) As Boolean
        If BaseCial.FactoryConditionLivraison.ExistIntitule(Trim(EnteteConditiondeLivraison)) = False Then
            ErreurJrn.WriteLine("< L'Intitulé Condition de livraison : " & Trim(EnteteConditiondeLivraison) & " - N°Pièce du Fichier : " & Trim(EntetePieceInterne) & " n'existe pas dans la table de paramétrage et dans Sage >")
            Return False
        End If
    End Function
    Public Function ExisteSOCIETE_ROUTAGE(ByVal CodeSociete As String) As Boolean
        Dim OleRecherAdapter, OleAdaptaterEnreg As OleDbDataAdapter
        Dim OleRecherDataset, OleEnregDataset As New DataSet
        Dim OleRechDatable, OledatableEnreg As DataTable

        Try
            If CodeSociete <> Nothing Then
                OleRecherAdapter = New OleDbDataAdapter("SELECT D_RaisonSoc FROM P_DOSSIER WHERE D_RaisonSoc='" & CodeSociete & "'", OleSocieteConnect)
                OleRecherAdapter.Fill(OleRecherDataset)
                OleRechDatable = OleRecherDataset.Tables(0)

                If OleRechDatable.Rows.Count <> 0 Then
                    OleAdaptaterEnreg = New OleDbDataAdapter("select * From PARAMETRE WHERE Societe='" & OleRechDatable.Rows(0).Item("D_RaisonSoc") & "' And  nomtype ='COMMERCIAL'", OleConnenection)
                    OleEnregDataset = New DataSet
                    OleAdaptaterEnreg.Fill(OleEnregDataset)
                    OledatableEnreg = OleEnregDataset.Tables(0)
                    If OleRechDatable.Rows.Count = 0 Then
                        ErreurJrn.WriteLine("< Cette Société [<" & CodeSociete & ">] n'existe pas dans la table de paramétrage>")
                        Return False
                    Else
                        Return True
                    End If
                Else
                    ErreurJrn.WriteLine("< Cette Société [<" & CodeSociete & ">] n'existe pas dans la Sage >")
                    Return False
                End If
            Else
                Return False
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Existe Société")
        End Try
    End Function
    Private Function RenvoieDepotPrincipal(ByVal DE_NO As String) As String
        Dim DossierAdap As OleDbDataAdapter
        Dim DossierDs As DataSet
        Dim DossierTab As DataTable
        If IsNumeric(DE_NO) Then
            DossierAdap = New OleDbDataAdapter("select * from F_DEPOT WHERE DE_NO=" & CInt(DE_NO), OleExcelConnect)
        Else
            DossierAdap = New OleDbDataAdapter("select * from F_DEPOT WHERE DE_INTITULE='" & DE_NO & "'", OleExcelConnect)
        End If
        DossierDs = New DataSet
        DossierAdap.Fill(DossierDs)
        DossierTab = DossierDs.Tables(0)
        If DossierTab.Rows.Count <> 0 Then
            Return DossierTab.Rows(0).Item("DE_Intitule")
        Else
            Return Nothing
        End If
    End Function
    Dim valeurDefaultDepot As String = ""
    Public Function ExisteDepot(ByVal IntituleDepot As String, ByVal EntetePieceInterne As String) As Boolean
        If IntituleDepot.Trim <> "" Then
            If BaseCial.FactoryDepot.ExistIntitule(Trim(IntituleDepot)) = False Then
                If valeurDefaultDepot.Trim <> "" Then
                    If BaseCial.FactoryDepot.ExistIntitule(Trim(valeurDefaultDepot.Trim)) = False Then
                        Return False
                    Else
                        IDDepotEntete = valeurDefaultDepot.Trim
                        Return True
                    End If
                End If
                ErreurJrn.WriteLine("Le Dépôt [ " & IntituleDepot & " ] Correspondant au N°Pièce du Fichier : " & EntetePieceInterne.Trim & " n'existe pas dans Sage")

                Return False
            Else
                Return True
            End If
        End If
    End Function
    Private Function VerificationStockDispoDepot(ByRef IDDepot As String, ByRef ArticleRef As String, ByVal LigneQuantite As Double, ByVal EntetePieceInterne As String) As Boolean
        Dim DossierAdap As OleDbDataAdapter
        Dim DossierDs As DataSet
        Dim DossierTab As DataTable
        Try
            ListeStock = New List(Of String)
            If IsNumeric(IDDepot) Then
                DossierAdap = New OleDbDataAdapter("select * from F_ARTSTOCK where DE_No ='" & CInt(Trim(IDDepot)) & "' And AR_Ref='" & Trim(ArticleRef) & "'", OleSocieteConnect)
            Else
                DossierAdap = New OleDbDataAdapter("select * from F_ARTSTOCK INNER JOIN F_DEPOT ON  F_DEPOT.DE_Intitule ='" & Trim(IDDepot) & "' And F_ARTSTOCK.AR_Ref='" & Trim(ArticleRef) & "' AND F_ARTSTOCK.DE_No=F_DEPOT.DE_No", OleSocieteConnect)
            End If
            If BaseCial.FactoryArticle.ExistReference(Trim(ArticleRef)) = False Then
                ErreurJrn.WriteLine("<-La Référence de l'article[<" & ArticleRef & ">]  present dans le Fichier :" & EntetePieceInterne & " n'existe pas dans Sage->")
                ErreurJrn.WriteLine("")
                ExisteLecture = False
            Else
                DossierDs = New DataSet
                DossierAdap.Fill(DossierDs)
                DossierTab = DossierDs.Tables(0)
                If DossierTab.Rows.Count <> 0 Then
                    If Convert.IsDBNull(DossierTab.Rows(0).Item("AS_QteSto")) = False Then
                        If Convert.IsDBNull(DossierTab.Rows(0).Item("AS_QtePrepa")) = False Then
                            If (DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) < LigneQuantite Then
                                ErreurJrn.WriteLine("")
                                ErreurJrn.WriteLine("Le Stock : " & (DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) & " dans le dépôt : " & Trim(IDDepot) & "   ne permet pas de Créer l'article : " & ArticleRef.Trim & " de Quantité : " & LigneQuantite & " ,  N°Pièce du Fichier :" & Trim(EntetePieceInterne))
                                ExisteLecture = False
                            Else
                                If ListeStock.Count <> 0 Then
                                    Dim ExisteArtStock As Boolean = False
                                    For i As Integer = 0 To ListeStock.Count - 1
                                        Dim StockListe() As String = Split(ListeStock.Item(i), ControlChars.Tab)
                                        If StockListe(0) = CInt(Trim(IDDepot)) And Trim(StockListe(1)) = Trim(ArticleRef) Then
                                            ExisteArtStock = True
                                            If StockListe(2) < LigneQuantite Then
                                                ExisteLecture = False
                                                ErreurJrn.WriteLine("Le Stock Calculé " & StockListe(2) & " dans le dépôt : " & StockListe(0) & "   ne permet pas de créer l'article : " & ArticleRef.Trim & " de Quantité : " & LigneQuantite & " ,  N° :" & Trim(EntetePieceInterne))
                                                ErreurJrn.WriteLine("")
                                            Else
                                                ListeStock.RemoveAt(i)
                                                ListeStock.Add(Trim(IDDepot) & ControlChars.Tab & Trim(ArticleRef) & ControlChars.Tab & (StockListe(2) - LigneQuantite))
                                                ErreurJrn.WriteLine("Le Stock : " & (StockListe(2) - LigneQuantite & " dans le dépôt : " & Trim(IDDepot) & "  l'article : " & ArticleRef.Trim & " de Quantité à Importer : " & LigneQuantite & " ,  N° :" & Trim(EntetePieceInterne)))
                                                Exit For
                                            End If
                                        End If
                                    Next i
                                    If ExisteArtStock = False Then
                                        ListeStock.Add(Trim(IDDepot) & ControlChars.Tab & Trim(ArticleRef) & ControlChars.Tab & ((DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) - LigneQuantite))
                                        ErreurJrn.WriteLine("Le Stock : " & ((DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) - LigneQuantite & " dans le dépôt : " & Trim(IDDepot) & "  l'article : " & ArticleRef.Trim & " de Quantité à Importer : " & LigneQuantite & " ,  N° :" & Trim(EntetePieceInterne)))
                                    End If
                                Else
                                    ListeStock.Add(Trim(IDDepot) & ControlChars.Tab & Trim(ArticleRef) & ControlChars.Tab & ((DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) - LigneQuantite))
                                    ErreurJrn.WriteLine("Le Stock : " & ((DossierTab.Rows(0).Item("AS_QteSto") - DossierTab.Rows(0).Item("AS_QtePrepa")) - LigneQuantite) & " dans le dépôt : " & Trim(IDDepot) & "  l'article : " & ArticleRef.Trim & " de Quantité à Importer : " & LigneQuantite & " ,  N° :" & Trim(EntetePieceInterne))
                                End If
                            End If
                        Else
                            If DossierTab.Rows(0).Item("AS_QteSto") < LigneQuantite Then
                                ExisteLecture = False
                                ErreurJrn.WriteLine("Le Stock : " & DossierTab.Rows(0).Item("AS_QteSto") & " dans le dépôt : " & Trim(IDDepot) & "   ne permet pas de créer l'article : " & ArticleRef.Trim & " de Quantité : " & LigneQuantite & " ,  N° :" & Trim(EntetePieceInterne))
                            Else
                                If ListeStock.Count <> 0 Then
                                    Dim ExisteArtStock As Boolean = False
                                    For i As Integer = 0 To ListeStock.Count - 1
                                        Dim StockListe() As String = Split(ListeStock.Item(i), ControlChars.Tab)
                                        If StockListe(0) = Trim(IDDepot) And Trim(StockListe(1)) = Trim(ArticleRef) Then
                                            ExisteArtStock = True
                                            If StockListe(2) < CDbl(Trim(LigneQuantite)) Then
                                                ExisteLecture = False
                                                ErreurJrn.WriteLine("")
                                                ErreurJrn.WriteLine("Le Stock Calculé " & StockListe(2) & " dans le dépôt : " & StockListe(0) & "   ne permet pas de créer l'article : " & ArticleRef.Trim & " de Quantité : " & CDbl(Trim(LigneQuantite)) & " ,  N° :" & Trim(EntetePieceInterne))
                                            Else
                                                ListeStock.RemoveAt(i)
                                                ListeStock.Add(Trim(IDDepot) & ControlChars.Tab & Trim(ArticleRef) & ControlChars.Tab & (StockListe(2) - LigneQuantite))
                                                ErreurJrn.WriteLine("Le Stock : " & (StockListe(2) - LigneQuantite) & " dans le dépôt : " & Trim(IDDepot) & "  l'article : " & ArticleRef.Trim & " de Quantité à Importer : " & LigneQuantite & " ,  N° :" & Trim(EntetePieceInterne))
                                                Exit For
                                            End If
                                        End If
                                    Next i
                                    If ExisteArtStock = False Then
                                        ListeStock.Add(Trim(IDDepot) & ControlChars.Tab & Trim(ArticleRef) & ControlChars.Tab & (DossierTab.Rows(0).Item("AS_QteSto") - LigneQuantite))
                                        ErreurJrn.WriteLine("Le Stock : " & (DossierTab.Rows(0).Item("AS_QteSto") - LigneQuantite) & " dans le dépôt : " & Trim(IDDepot) & "  l'article : " & ArticleRef.Trim & " de Quantité à Importer : " & LigneQuantite & " ,  N° :" & Trim(EntetePieceInterne))
                                    End If
                                Else
                                    ListeStock.Add(Trim(IDDepot) & ControlChars.Tab & Trim(ArticleRef) & ControlChars.Tab & (DossierTab.Rows(0).Item("AS_QteSto") - LigneQuantite))
                                    ErreurJrn.WriteLine("Le Stock : " & (DossierTab.Rows(0).Item("AS_QteSto") - LigneQuantite) & " dans le dépôt : " & Trim(IDDepot) & "  l'article : " & ArticleRef.Trim & " de Quantité à Importer : " & LigneQuantite & " ,  N° :" & Trim(EntetePieceInterne))
                                End If
                            End If
                        End If
                    Else
                        ExisteLecture = False
                        ErreurJrn.WriteLine("")
                        ErreurJrn.WriteLine("Le Stock dans le dépôt : " & Trim(IDDepot) & " est NULL et  ne permet pas de créer l'article : " & ArticleRef.Trim & " ,  N° :" & Trim(EntetePieceInterne))
                    End If
                Else
                    ExisteLecture = False
                    ErreurJrn.WriteLine("")
                    ErreurJrn.WriteLine("Le Stock dans le dépôt : " & Trim(IDDepot) & "   ne permet pas de créer l'article : " & ArticleRef.Trim & " ,  N° :" & Trim(EntetePieceInterne))
                End If
            End If
        Catch ex As Exception
            ExisteLecture = False
        End Try
        If ExisteLecture = False Then
            infosExport.Text = "Verification du Stock du dépôt <[" & Trim(IDDepot) & "]>"
        End If
        VerificationStockDispoDepot = ExisteLecture
    End Function
    Public Function ExisteMappingSage(ByVal NameColonne As String, Optional ByVal TableLie As String = "") As String
        Dim OleAdaptaterschema As OleDbDataAdapter
        Dim OleSchemaDataset As DataSet
        Dim OledatableSchema As DataTable
        Dim ColonneMappe As String = ""
        Try
            If TableLie <> "" Then
                OleAdaptaterschema = New OleDbDataAdapter("select Champ,Libelle from COLIMPMOUV WHERE Fichier='F_DOCLIGNE' AND Champ='" & NameColonne & "'", OleConnenection)
            Else
                OleAdaptaterschema = New OleDbDataAdapter("select Champ,Libelle from COLIMPMOUV WHERE Fichier='F_DOCENTETE' AND Champ='" & NameColonne & "'", OleConnenection)
            End If
            OleSchemaDataset = New DataSet
            OleAdaptaterschema.Fill(OleSchemaDataset)
            OledatableSchema = OleSchemaDataset.Tables(0)
            If OledatableSchema.Rows.Count <> 0 Then
                If TableLie <> "" Then
                    ColonneMappe = OledatableSchema.Rows(0).Item("Libelle")
                Else
                    ColonneMappe = OledatableSchema.Rows(0).Item("Champ")
                End If
            End If
        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Critical, "Recherche Correspondance Sage ")
            ColonneMappe = ""
        End Try
        ExisteMappingSage = ColonneMappe
    End Function
    ' Public PieceCommande As String = ""

    Private Sub Integrer_Ecriture_Ligne(ByVal Document_Ligne As String, Optional ByVal TableLie As String = "")
        Dim fournisseurAdap As OleDbDataAdapter
        Dim fournisseurDs As DataSet
        Dim fournisseurTab As DataTable
        Dim PieceAutoma As String = ""
        'Dim TableLie As String = "F_DOCLIGNE"
        infosExport.Refresh()
        infosExport.Text = "Integration des Ligne En Cours..."
        If RbtG1.Checked Or RbtG3.Checked Then
            fournisseurAdap = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='CSO' AND Entete=" & False & " AND InfosLibre=" & False & " AND Ligne=" & True & " AND ChampSage<>'' ORDER BY ORDRE", OleConnenection)
        End If
        If RbtG2.Checked Then
            fournisseurAdap = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='CSO' AND Entete=" & False & " AND InfosLibre=" & False & " AND Ligne=" & False & " AND ChampSage<>'' ORDER BY ORDRE", OleConnenection)
        End If
        If RbtG4.Checked Then
            fournisseurAdap = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='CSO' AND Entete=" & False & " AND InfosLibre=" & True & " AND Ligne=" & True & " AND ChampSage<>'' ORDER BY ORDRE", OleConnenection)
        End If
        'RbtG3.Checked
        'fournisseurAdap = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='CSO' AND Entete=" & False & " AND InfosLibre=" & False & " AND Ligne=" & True & " AND ChampSage<>'' ORDER BY ORDRE", OleConnenection)
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
                        Dim g = fournisseurTab.Rows(numColonne).Item("ChampSage").ToString.Trim
                        'Ligne Document
                        If fournisseurTab.Rows(numColonne).Item("Cols").ToString.Trim = "LINE_NUMBER" Then
                            If fournisseurTab.Rows(numColonne).Item("ChampSage").ToString.Trim <> "" Then
                                NLignePieceCommande = Strings.Mid(Document_Ligne, PositionG, Longueur)
                                If Trim(NLignePieceCommande) = "" Then
                                    NLignePieceCommande = DefaultValeur.Trim
                                End If
                                Continue For
                            End If
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LigneCodeAffaire" Then
                            LigneCodeAffaire = Strings.Mid(Document_Ligne, PositionG, Longueur)
                            If Trim(LigneCodeAffaire) = "" Then
                                LigneCodeAffaire = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LigneDatedeFabrication" Then
                            LigneDatedeFabrication = Strings.Mid(Document_Ligne, PositionG, Longueur)
                            If Trim(LigneDatedeFabrication) = "" Then
                                LigneDatedeFabrication = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LigneDatedeLivraison" Then
                            LigneDatedeLivraison = Strings.Mid(Document_Ligne, PositionG, Longueur)
                            If Trim(LigneDatedeLivraison) = "" Then
                                LigneDatedeLivraison = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LigneDatedePeremption" Then
                            LigneDatedePeremption = Strings.Mid(Document_Ligne, PositionG, Longueur).Trim
                            If Trim(LigneDatedePeremption) = "" Then
                                LigneDatedePeremption = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LigneDesignationArticle" Then
                            LigneDesignationArticle = Strings.Mid(Document_Ligne, PositionG, Longueur).Trim
                            If Strings.Len(Trim(LigneDesignationArticle)) <= 69 Then
                                LigneDesignationArticle = Strings.Mid(Document_Ligne, PositionG, Longueur).Trim
                                If Trim(LigneDesignationArticle) = "" Then
                                    LigneDesignationArticle = DefaultValeur.Trim
                                End If
                            Else
                                LigneDesignationArticle = Strings.Left(Trim(LigneDesignationArticle), 69)
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LigneLibelleComplementaire" Then
                            LigneLibelleComplementaire = Strings.Mid(Document_Ligne, PositionG, Longueur).Trim
                            If Trim(LigneLibelleComplementaire) = "" Then
                                LigneLibelleComplementaire = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LigneEnumereConditionnement" Then
                            LigneEnumereConditionnement = Strings.Mid(Document_Ligne, PositionG, Longueur).Trim
                            If Trim(LigneEnumereConditionnement) = "" Then
                                LigneEnumereConditionnement = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LigneFraisApproche" Then
                            LigneFraisApproche = Strings.Mid(Document_Ligne, PositionG, Longueur).Trim
                            If Trim(LigneFraisApproche) = "" Then
                                LigneFraisApproche = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LigneIntituleDepot" Then
                            LigneIntituleDepot = Strings.Mid(Document_Ligne, PositionG, Longueur).Trim
                            If Trim(LigneIntituleDepot) = "" Then
                                LigneIntituleDepot = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LigneNomRepresentant" Then
                            LigneNomRepresentant = Strings.Mid(Document_Ligne, PositionG, Longueur).Trim
                            If Trim(LigneNomRepresentant) = "" Then
                                LigneNomRepresentant = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LigneNSerieLot" Then
                            LigneNSerieLot = Strings.Mid(Document_Ligne, PositionG, Longueur).Trim
                            If Strings.Len(Trim(LigneNSerieLot)) <= 30 Then
                                LigneNSerieLot = Trim(LigneNSerieLot)
                            Else
                                LigneNSerieLot = Strings.Left(Trim(LigneNSerieLot), 30)
                            End If
                            If Trim(LigneNomRepresentant) = "" Then
                                If Strings.Len(Trim(DefaultValeur.Trim)) <= 30 Then
                                    LigneNSerieLot = Trim(DefaultValeur.Trim)
                                Else
                                    LigneNSerieLot = Strings.Left(Trim(DefaultValeur.Trim), 30)
                                End If
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LignePlanAnalytique" Then
                            LignePlanAnalytique = Strings.Mid(Document_Ligne, PositionG, Longueur).Trim
                            If Trim(LignePlanAnalytique) = "" Then
                                LignePlanAnalytique = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LignePoidsBrut" Then
                            If Strings.Mid(Document_Ligne, PositionG, Longueur).Trim <> "" Then
                                LignePoidsBrut = Strings.Mid(Document_Ligne, PositionG, Longueur).Trim / Math.Pow(10, 5)
                            Else
                                LignePoidsBrut = ""
                            End If
                            If Trim(LignePoidsBrut) = "" Then
                                LignePoidsBrut = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LignePoidsNet" Then
                            If Strings.Mid(Document_Ligne, PositionG, Longueur).Trim <> "" Then
                                LignePoidsNet = Strings.Mid(Document_Ligne, PositionG, Longueur).Trim / Math.Pow(10, 5)
                            Else
                                LignePoidsNet = ""
                            End If
                            If Trim(LignePoidsNet) = "" Then
                                LignePoidsNet = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LignePrenomRepresentant" Then
                            LignePrenomRepresentant = Strings.Mid(Document_Ligne, PositionG, Longueur).Trim
                            If Trim(LignePrenomRepresentant) = "" Then
                                LignePrenomRepresentant = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LignePrixdeRevientUnitaire" Then
                            LignePrixdeRevientUnitaire = Strings.Mid(Document_Ligne, PositionG, Longueur).Trim
                            If Trim(LignePrixdeRevientUnitaire) = "" Then
                                LignePrixdeRevientUnitaire = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LignePrixUnitaire" Then
                            LignePrixUnitaire = Strings.Mid(Document_Ligne, PositionG, Longueur).Trim
                            If Trim(LignePrixUnitaire) = "" Then
                                LignePrixUnitaire = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LignePrixUnitaireDevise" Then
                            LignePrixUnitaireDevise = Strings.Mid(Document_Ligne, PositionG, Longueur).Trim
                            If Trim(LignePrixUnitaireDevise) = "" Then
                                LignePrixUnitaireDevise = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LigneValorisé" Then
                            LigneValorisé = Strings.Mid(Document_Ligne, PositionG, Longueur).Trim
                            If Trim(LigneValorisé) = "" Then
                                LigneValorisé = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LigneTypePrixUnitaire" Then
                            LigneTypePrixUnitaire = Strings.Mid(Document_Ligne, PositionG, Longueur).Trim
                            If Trim(LigneTypePrixUnitaire) = "" Then
                                LigneTypePrixUnitaire = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LigneQuantite" Then
                            If Strings.Mid(Document_Ligne, PositionG, Longueur).Trim <> "" Then
                                LigneQuantite = Strings.Mid(Document_Ligne, PositionG, Longueur).Trim / Math.Pow(10, 5)
                            Else
                                LigneQuantite = ""
                            End If
                            If Trim(LigneQuantite) = "" Then
                                LigneQuantite = DefaultValeur.Trim
                            End If
                            Continue For
                        End If

                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LigneQuantiteConditionne" Then
                            If Strings.Mid(Document_Ligne, PositionG, Longueur).Trim <> "" Then
                                LigneQuantiteConditionne = Strings.Mid(Document_Ligne, PositionG, Longueur).Trim / Math.Pow(10, 5)
                            Else
                                LigneQuantiteConditionne = ""
                            End If
                            If Trim(LigneQuantiteConditionne) = "" Then
                                If DefaultValeur.Trim <> "" Then
                                    LigneQuantite = DefaultValeur.Trim
                                End If
                                Continue For
                            End If
                            LigneQuantite = LigneQuantiteConditionne
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LigneReference" Then
                            LigneReference = Strings.Mid(Document_Ligne, PositionG, Longueur).Trim
                            If Strings.Len(Trim(LigneReference)) <= 17 Then
                                LigneReference = Trim(LigneReference)
                            Else
                                LigneReference = Strings.Left(Trim(LigneReference), 17)
                            End If
                            If LigneReference = "" Then
                                If Strings.Len(Trim(DefaultValeur.Trim)) <= 17 Then
                                    LigneReference = Trim(DefaultValeur.Trim)
                                Else
                                    LigneReference = Strings.Left(Trim(DefaultValeur.Trim), 17)
                                End If
                            End If
                            Continue For
                        End If

                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LigneArticleCompose" Then
                            LigneArticleCompose = Strings.Mid(Document_Ligne, PositionG, Longueur).Trim
                            LigneArticleCompose = Formatage_Article(Trim(LigneArticleCompose))
                            If LigneArticleCompose = "" Then
                                LigneArticleCompose = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LigneCodeArticle" Then
                            LigneCodeArticle = Strings.Mid(Document_Ligne, PositionG, Longueur).Trim
                            LigneCodeArticle = Formatage_Article(Trim(LigneCodeArticle))
                            If LigneCodeArticle = "" Then
                                LigneCodeArticle = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LigneReferenceArticleTiers" Then
                            LigneReferenceArticleTiers = Strings.Mid(Document_Ligne, PositionG, Longueur).Trim
                            If Strings.Len(Trim(LigneReferenceArticleTiers)) <= 18 Then
                                LigneReferenceArticleTiers = Trim(LigneReferenceArticleTiers)
                            Else
                                LigneReferenceArticleTiers = Strings.Left(Trim(LigneReferenceArticleTiers), 18)
                            End If
                            If LigneReferenceArticleTiers = "" Then
                                LigneReferenceArticleTiers = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LigneTauxRemise1" Then
                            LigneTauxRemise1 = Strings.Mid(Document_Ligne, PositionG, Longueur).Trim
                            If LigneTauxRemise1 = "" Then
                                LigneTauxRemise1 = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LigneTauxRemise2" Then
                            LigneTauxRemise2 = Strings.Mid(Document_Ligne, PositionG, Longueur).Trim
                            If LigneTauxRemise2 = "" Then
                                LigneTauxRemise2 = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LigneTauxRemise3" Then
                            LigneTauxRemise3 = Strings.Mid(Document_Ligne, PositionG, Longueur).Trim
                            If LigneTauxRemise3 = "" Then
                                LigneTauxRemise3 = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LigneTypeRemise1" Then
                            LigneTypeRemise1 = Strings.Mid(Document_Ligne, PositionG, Longueur).Trim
                            If LigneTypeRemise1 = "" Then
                                LigneTypeRemise1 = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LigneTypeRemise2" Then
                            LigneTypeRemise2 = Strings.Mid(Document_Ligne, PositionG, Longueur).Trim
                            If LigneTypeRemise2 = "" Then
                                LigneTypeRemise2 = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString, TableLie).Trim = "LigneTypeRemise3" Then
                            LigneTypeRemise3 = Strings.Mid(Document_Ligne, PositionG, Longueur).Trim
                            If LigneTypeRemise3 = "" Then
                                LigneTypeRemise3 = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                    Else
                        Continue For
                    End If
                Next
                Dim r = 1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Fonction d'integration")
        End Try
    End Sub
    Private Sub Integrer_Ecriture_Entete(ByVal Document_Entete As String, Optional ByVal PieceArticle As String = "")
        Dim fournisseurAdap As OleDbDataAdapter
        Dim fournisseurDs As DataSet
        Dim fournisseurTab As DataTable
        Dim PieceAutoma As String = ""
        infosExport.Refresh()
        infosExport.Text = "Integration En Cours..."
        fournisseurAdap = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='CSO' AND Entete=" & True & " AND InfosLibre=" & False & " AND Ligne=" & False & " AND ChampSage<>'' ORDER BY ORDRE", OleConnenection)
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
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "DO_Piece" Then
                            EntetePieceInterne = Strings.Mid(Document_Entete, PositionG, Longueur)
                            If EntetePieceInterne.ToString.Trim <> "" Then
                                If Strings.Len(Trim(EntetePieceInterne.ToString.Trim)) <= 8 Then
                                    EntetePieceInterne = Formatage_Chaine(Trim(EntetePieceInterne.ToString.Trim))
                                End If
                            ElseIf DefaultValeur.Trim <> "" Then
                                EntetePieceInterne = DefaultValeur.Trim
                                If Strings.Len(Trim(DefaultValeur.Trim)) <= 8 Then
                                    EntetePieceInterne = Formatage_Chaine(Trim(DefaultValeur.Trim))
                                End If
                            End If
                            Continue For
                        End If

                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "AR_REF" Then
                            PieceArticle = Trim(PieceArticle)
                            Continue For
                        End If

                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "MR_No" Then
                            EcheanceConditionPaiement = Strings.Mid(Document_Entete, PositionG, Longueur)
                            If Trim(EcheanceConditionPaiement) = "" Then
                                EcheanceConditionPaiement = DefaultValeur.Trim
                            End If
                            Continue For
                        End If

                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "MR_No" Then
                            EcheanceModeleReglement = Strings.Mid(Document_Entete, PositionG, Longueur)
                            If Trim(EcheanceModeleReglement) = "" Then
                                EcheanceModeleReglement = DefaultValeur.Trim
                            End If
                            Continue For
                        End If

                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "N_Reglement" Then
                            EcheanceModeleReglement = Strings.Mid(Document_Entete, PositionG, Longueur)
                            If Trim(EcheanceModeleReglement) = "" Then
                                EcheanceModeleReglement = DefaultValeur.Trim
                            End If
                            Continue For
                        End If

                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "DR_Date" Then
                            EcheanceDatePied = Strings.Mid(Document_Entete, PositionG, Longueur)
                            If Trim(EcheanceDatePied) = "" Then
                                EcheanceDatePied = DefaultValeur.Trim
                            End If
                            Continue For
                        End If

                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "DO_Statut" Then
                            EnteteStatutdocument = Strings.Mid(Document_Entete, PositionG, Longueur)
                            If Trim(EnteteStatutdocument) = "" Then
                                EnteteStatutdocument = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "DO_ValFrais" Then
                            EnteteFraisExpedition = Strings.Mid(Document_Entete, PositionG, Longueur)
                            If Trim(EnteteFraisExpedition) = "" Then
                                EnteteFraisExpedition = DefaultValeur.Trim
                            End If
                            Continue For
                        End If

                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "DO_Provenance" Then
                            ProvenanceFacture = Strings.Mid(Document_Entete, PositionG, Longueur)
                            If Trim(ProvenanceFacture) = "" Then
                                ProvenanceFacture = DefaultValeur.Trim
                            End If
                            Continue For
                        End If

                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "DO_TypeColis" Then
                            EnteteUniteColis = Strings.Mid(Document_Entete, PositionG, Longueur)
                            If Trim(EnteteUniteColis) = "" Then
                                EnteteUniteColis = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "DO_Type" Then
                            EnteteTyPeDocument = Strings.Mid(Document_Entete, PositionG, Longueur)
                            If Trim(EnteteTyPeDocument) = "" Then
                                EnteteTyPeDocument = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        'si ChampSage Do_Tiers 
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "DO_Tiers" Then
                            EnteteCodeTiers = Strings.Mid(Document_Entete, PositionG, Longueur)
                            If Trim(EnteteCodeTiers) = "" Then
                                EnteteCodeTiers = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        'si ChampSage CT_Num
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "CT_NUM" Then
                            'traitement autres à faire ici pour determine le type de piece 
                            EnteteCodeFournisseur = Strings.Mid(Document_Entete, 307, 20).Trim
                            EnteteCodeTiers = Strings.Mid(Document_Entete, 327, 12).Trim
                            EnteteCodeTransfertDepot = Strings.Mid(Document_Entete, 339, 10).Trim

                            If Trim(EnteteCodeFournisseur) = "" Then
                                EnteteCodeFournisseur = DefaultValeur.Trim
                            Else
                                EnteteTyPeDocument = "14"
                                lblType.Text = "Bon de Retour"
                            End If
                            If Trim(EnteteCodeTiers) = "" Then
                                EnteteCodeTiers = DefaultValeur.Trim
                            Else
                                EnteteTyPeDocument = "3"
                                lblType.Text = "Bon de Livraison"
                            End If
                            If Trim(EnteteCodeTransfertDepot) = "" Then
                                EnteteCodeTransfertDepot = DefaultValeur.Trim
                            Else
                                EnteteTyPeDocument = "23"
                                lblType.Text = "Transfért dépôts"
                            End If
                            Continue For
                        End If

                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "DO_Contact" Then
                            EnteteContact = Strings.Mid(Document_Entete, PositionG, Longueur)
                            If Strings.Len(Trim(EnteteContact)) <= 35 Then
                                EnteteContact = Trim(EnteteContact)
                            Else
                                EnteteContact = Strings.Left(Trim(EnteteContact), 35)
                            End If
                            If Trim(EnteteContact) = "" Then
                                EnteteContact = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        'ici on renseigne le cour qui ce trouve dans notre fichier a importe 
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "DO_Cours" Then
                            EnteteCours = Strings.Mid(Document_Entete, PositionG, Longueur)
                            If Trim(EnteteCours) = "" Then
                                EnteteCours = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "DO_Coord01" Then
                            Entete1 = Strings.Mid(Document_Entete, PositionG, Longueur)
                            If Trim(Entete1) = "" Then
                                Entete1 = DefaultValeur.Trim
                            End If
                            If Strings.Len(Trim(Entete1)) <= 25 Then
                                Entete1 = Trim(Entete1)
                            Else
                                Entete1 = Strings.Left(Trim(Entete1), 25)
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "DO_Coord02" Then
                            Entete2 = Strings.Mid(Document_Entete, PositionG, Longueur)
                            If Trim(Entete2) = "" Then
                                Entete2 = DefaultValeur.Trim
                            End If
                            If Strings.Len(Trim(Entete2)) <= 25 Then
                                Entete2 = Trim(Entete2)
                            Else
                                Entete2 = Strings.Left(Trim(Entete2), 25)
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "DO_Coord03" Then
                            Entete3 = Strings.Mid(Document_Entete, PositionG, Longueur)
                            If Trim(Entete3) = "" Then
                                Entete3 = DefaultValeur.Trim
                            End If
                            If Strings.Len(Trim(Entete3)) <= 25 Then
                                Entete3 = Trim(Entete3)
                            Else
                                Entete3 = Strings.Left(Trim(Entete3), 25)
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "DO_Coord04" Then
                            Entete4 = Strings.Mid(Document_Entete, PositionG, Longueur)
                            If Trim(Entete4) = "" Then
                                Entete4 = DefaultValeur.Trim
                            End If
                            If Strings.Len(Trim(Entete4)) <= 25 Then
                                Entete4 = Trim(Entete4)
                            Else
                                Entete4 = Strings.Left(Trim(Entete4), 25)
                            End If
                            Continue For
                        End If

                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "DO_BLFact" Then
                            EnteteBLFacture = Strings.Mid(Document_Entete, PositionG, Longueur)
                            If Trim(EnteteBLFacture) = "" Then
                                EnteteBLFacture = DefaultValeur.Trim
                            End If
                            Continue For
                        End If

                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "N_CatCompta" Then
                            EnteteCatégorieComptable = Strings.Mid(Document_Entete, PositionG, Longueur)
                            If Trim(EnteteCatégorieComptable) = "" Then
                                EnteteCatégorieComptable = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "DO_Tarif" Then
                            EnteteCatégorietarifaire = Strings.Mid(Document_Entete, PositionG, Longueur)
                            If Trim(EnteteCatégorietarifaire) = "" Then
                                EnteteCatégorietarifaire = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "CA_Num" Then
                            EnteteCodeAffaire = Strings.Mid(Document_Entete, PositionG, Longueur)
                            If Trim(EnteteCodeAffaire) = "" Then
                                EnteteCodeAffaire = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "CT_NumPayeur" Then
                            EnteteCodeTiersPayeur = Strings.Mid(Document_Entete, PositionG, Longueur)
                            If Trim(EnteteCodeTiersPayeur) = "" Then
                                EnteteCodeTiersPayeur = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "DO_Colisage" Then
                            EnteteColisagedeLivraison = Strings.Mid(Document_Entete, PositionG, Longueur)
                            If Trim(EnteteColisagedeLivraison) = "" Then
                                EnteteColisagedeLivraison = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "CG_Num" Then
                            EnteteCompteGeneral = Strings.Mid(Document_Entete, PositionG, Longueur)
                            If Trim(EnteteCompteGeneral) = "" Then
                                EnteteCompteGeneral = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "DO_Condition" Then
                            EnteteConditiondeLivraison = Strings.Mid(Document_Entete, PositionG, Longueur)
                            If Trim(EnteteConditiondeLivraison) = "" Then
                                EnteteConditiondeLivraison = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "DO_Date" Then
                            EnteteDateDocument = Trim(Strings.Mid(Document_Entete, PositionG, Longueur)).Substring(0, 8)
                            If Trim(EnteteDateDocument) = "" Then
                                EnteteDateDocument = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "DO_DateLivr" Then
                            EnteteDateLivraison = Trim(Strings.Mid(Document_Entete, PositionG, Longueur))
                            If Trim(EnteteDateDocument) = "" Then
                                EnteteDateLivraison = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "DO_Ecart" Then
                            EnteteEcartValorisation = Trim(Strings.Mid(Document_Entete, PositionG, Longueur))
                            If Trim(EnteteEcartValorisation) = "" Then
                                EnteteEcartValorisation = DefaultValeur.Trim
                            End If
                            Continue For
                        End If

                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "DE_No" Then
                            IDDepotEntete = Trim(Strings.Mid(Document_Entete, PositionG, Longueur))
                            If DefaultValeur.Trim <> "" Then
                                valeurDefaultDepot = DefaultValeur.Trim
                                IDDepotEntete = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If fournisseurTab.Rows(numColonne).Item("Cols").ToString = "DOCUMENT" Then
                            CodeSociete = Trim(Strings.Mid(Document_Entete, PositionG, Longueur))
                            If CodeSociete = "" Then
                                If DefaultValeur.Trim <> "" Then
                                    CodeSociete = DefaultValeur.Trim
                                End If
                            End If
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "DE_No" Then
                            EnteteIntituleDepot = Trim(Strings.Mid(Document_Entete, PositionG, Longueur))
                            If Trim(EnteteIntituleDepot) = "" Then
                                EnteteIntituleDepot = DefaultValeur.Trim
                            End If
                            EnteteIntituleDepot = Trim(Strings.Mid(Document_Entete, PositionG, Longueur))
                            If DefaultValeur.Trim <> "" Then
                                valeurDefaultDepot = DefaultValeur.Trim
                                EnteteIntituleDepot = DefaultValeur.Trim
                            End If
                            Continue For
                        End If

                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "LI_No" Then
                            EnteteIntituleDepotClient = Trim(Strings.Mid(Document_Entete, PositionG, Longueur))
                            If Trim(EnteteIntituleDepotClient) = "" Then
                                EnteteIntituleDepotClient = DefaultValeur.Trim
                            End If
                            Continue For
                        End If

                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "DO_Devise" Then
                            EnteteIntituleDevise = Trim(Strings.Mid(Document_Entete, PositionG, Longueur))
                            If Trim(EnteteIntituleDevise) = "" Then
                                EnteteIntituleDevise = DefaultValeur.Trim
                            End If
                            Continue For
                        End If

                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "DO_Expedit" Then
                            EnteteIntituleExpédition = Trim(Strings.Mid(Document_Entete, PositionG, Longueur))
                            If Trim(EnteteIntituleExpédition) = "" Then
                                EnteteIntituleExpédition = DefaultValeur.Trim
                            End If
                            Continue For
                        End If

                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "DO_Transaction" Then
                            EnteteNatureTransaction = Trim(Strings.Mid(Document_Entete, PositionG, Longueur))
                            If Trim(EnteteNatureTransaction) = "" Then
                                EnteteNatureTransaction = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "DO_NbFacture" Then
                            EnteteNombredeFacture = Trim(Strings.Mid(Document_Entete, PositionG, Longueur))
                            If Trim(EnteteNombredeFacture) = "" Then
                                EnteteNombredeFacture = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "CO_No" Then
                            EnteteNomReprésentant = Trim(Strings.Mid(Document_Entete, PositionG, Longueur))
                            If Trim(EnteteNomReprésentant) = "" Then
                                EnteteNomReprésentant = DefaultValeur.Trim
                            End If
                            Continue For
                        End If

                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "CO_No" Then
                            EntetePrenomReprésentant = Trim(Strings.Mid(Document_Entete, PositionG, Longueur))
                            If Trim(EntetePrenomReprésentant) = "" Then
                                EntetePrenomReprésentant = DefaultValeur.Trim
                            End If
                            Continue For
                        End If

                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "EntetePlanAnalytique" Then
                            EntetePlanAnalytique = Trim(Strings.Mid(Document_Entete, PositionG, Longueur))
                            If Trim(EntetePlanAnalytique) = "" Then
                                EntetePlanAnalytique = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "DO_Ref" Then
                            EnteteReference = Trim(Strings.Mid(Document_Entete, PositionG, Longueur))
                            If Strings.Len(Trim(EnteteReference)) <= 17 Then
                                EnteteReference = Trim(Strings.Mid(Document_Entete, PositionG, Longueur))
                            Else
                                EnteteReference = Strings.Left(Trim(EnteteReference), 17)
                            End If
                            If Trim(EnteteReference) = "" Then
                                EnteteReference = DefaultValeur.Trim
                            End If
                            Continue For
                        End If

                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "DO_Regime" Then
                            EnteteRegimeDocument = Trim(Strings.Mid(Document_Entete, PositionG, Longueur))
                            If Trim(EnteteRegimeDocument) = "" Then
                                EnteteRegimeDocument = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "DO_Souche" Then
                            EnteteSoucheDocument = Trim(Strings.Mid(Document_Entete, PositionG, Longueur))
                            If Trim(EnteteSoucheDocument) = "" Then
                                EnteteSoucheDocument = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                        If ExisteMappingSage(fournisseurTab.Rows(numColonne).Item("ChampSage").ToString).Trim = "DO_TxEscompt" Then
                            EnteteTauxescompte = Trim(Strings.Mid(Document_Entete, PositionG, Longueur))
                            If Trim(EnteteTauxescompte) = "" Then
                                EnteteTauxescompte = DefaultValeur.Trim
                            End If
                            Continue For
                        End If
                    Else
                        Continue For
                    End If
                Next
                Dim r = 1
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Fonction d'integration")
        End Try
    End Sub
    Private Sub vidage()
        EcheanceConditionPaiement = Nothing
        EcheanceModeleReglement = Nothing
        EcheanceModeReglement = Nothing
        EcheanceDatePied = Nothing
        EnteteFraisExpedition = Nothing
        PieceCommande = Nothing
        PieceArticle = Nothing
        'EnteteStatutdocument = Nothing
        ProvenanceFacture = Nothing
        'LigneTypePrixUnitaire = Nothing
        EnteteUniteColis = Nothing
        LigneCodeArticle = Nothing
        EnteteBLFacture = Nothing
        EnteteCodeAffaire = Nothing
        EnteteCodeTiers = Nothing
        EnteteCodeTiersPayeur = Nothing
        EnteteColisagedeLivraison = Nothing
        EnteteCompteGeneral = Nothing
        EnteteDateDocument = Nothing
        EnteteDateLivraison = Nothing
        EnteteEcartValorisation = Nothing
        Entete1 = Nothing
        Entete2 = Nothing
        Entete3 = Nothing
        Entete4 = Nothing
        IDDepotEntete = Nothing
        IDDepotLigne = Nothing
        EnteteCatégorieComptable = Nothing
        EnteteCatégorietarifaire = Nothing
        EnteteConditiondeLivraison = Nothing
        EnteteIntituleDepot = Nothing
        EnteteIntituleDepotClient = Nothing
        EnteteIntituleDevise = Nothing
        EnteteIntituleExpédition = Nothing
        EntetePieceInterne = Nothing
        EnteteNatureTransaction = Nothing
        EnteteNomReprésentant = Nothing
        EnteteNombredeFacture = Nothing
        EntetePlanAnalytique = Nothing
        EntetePrenomReprésentant = Nothing
        EnteteReference = Nothing
        EnteteRegimeDocument = Nothing
        EnteteSoucheDocument = Nothing
        EnteteTauxescompte = Nothing
        EnteteTyPeDocument = Nothing
        LigneCodeAffaire = Nothing
        LigneDatedeFabrication = Nothing
        LigneDatedeLivraison = Nothing
        LigneDatedePeremption = Nothing
        LigneDesignationArticle = Nothing
        LigneLibelleComplementaire = Nothing
        LigneEnumereConditionnement = Nothing
        LigneFraisApproche = Nothing
        LigneIntituleDepot = Nothing
        LigneNSerieLot = Nothing
        LigneNomRepresentant = Nothing
        LignePlanAnalytique = Nothing
        LignePoidsBrut = Nothing
        LignePoidsNet = Nothing
        LignePrenomRepresentant = Nothing
        LignePrixdeRevientUnitaire = Nothing
        LignePrixUnitaire = Nothing
        LigneQuantite = Nothing
        LigneQuantiteConditionne = Nothing
        LigneReference = Nothing
        LigneArticleCompose = Nothing
        LigneReferenceArticleTiers = Nothing
        LigneTauxRemise1 = Nothing
        LigneTauxRemise2 = Nothing
        LigneTauxRemise3 = Nothing
        LigneTypeRemise1 = Nothing
        LigneTypeRemise2 = Nothing
        LigneTypeRemise3 = Nothing
        EnteteContact = Nothing
        EnteteLangue = Nothing
        EnteteCours = Nothing
        LignePrixUnitaireDevise = Nothing
    End Sub
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
    Public EstTrouverException As Boolean
    Public Sub BtnXFormation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnXFormation.Click
        Try
            EstTrouverException = False
            Button1_Click(sender, e)
            If Directory.Exists(PathsFileCSO) = True Then
                For i As Integer = 0 To Datagridaffiche.RowCount - 1
                    If Datagridaffiche.Rows(i).Cells("C6").Value = True Then
                        Transformation(Datagridaffiche.Rows(i).Cells("C8").Value)
                        If RegardeStatut = True And ExisteLectures = True Then
                            Datagridaffiche.Rows(i).Cells("C7").Value = My.Resources.accepter
                            If ChEncapsuler.Checked = False And EstTrouverException = False Then
                                FichierCSO.Close()
                            Else
                                EstTrouverException = False
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
            'MsgBox("Transformation " & ex.Message)
        End Try
    End Sub
End Class