Imports System.Data.OleDb
Imports System.IO
Public Class FrmExtractionBCClient
    Public ListeExporte As String
    Public NbInfosLibre As Integer
    Public NbInfosLibreVue As Integer
    Public OleAdaptaterschemaSage, OleAdaptaterschemaSageDetails, OleAdaptaterschema, OleAdaptaterschemaLigne, OleAdaptaterschemaFourssAR As OleDbDataAdapter
    Public OleSchemaDatasetSage, OleSchemaDatasetSageDetails, OleSchemaDataset, OleSchemaDatasetLigne, OleSchemaDatasetFourssAR As DataSet
    Public OledatableSchemaSage, _
    OledatableSchemaSageDetails, _
    OledatableSchema, _
    OledatableSchemaLigne, _
    OledatableSchemaFourssAR As DataTable
    Public ListeChampSageEntete As String = ""
    Public ListeChampSageLigne As String = ""
    Public VirguleFlotate As Integer = 0
    Public Item As String = ""
    Public TYPE_CODE As String = ""
    Public Shared ip As Integer
    Public TachePlanifie As String = ""
    Public SocieteCyble As New List(Of String)
    Public Sub AfficheSchemasTachePlanifie()
        Dim i As Integer
        Dim OleAdaptaterschema As OleDbDataAdapter
        Dim OleSchemaDataset As DataSet

        DataListeIntegrer.Rows.Clear()
        OleAdaptaterschema = New OleDbDataAdapter("select * from PARAMETRE WHERE nomtype='COMMERCIAL'", OleConnenectionArticle)
        OleSchemaDataset = New DataSet
        OleAdaptaterschema.Fill(OleSchemaDataset)
        OledatableSchema = OleSchemaDataset.Tables(0)
        DataListeIntegrer.RowCount = OledatableSchema.Rows.Count
        For i = 0 To OledatableSchema.Rows.Count - 1
            DataListeIntegrer.Rows(i).Cells("Societe1").Value = OledatableSchema.Rows(i).Item("Societe")
            DataListeIntegrer.Rows(i).Cells("Chemin1").Value = OledatableSchema.Rows(i).Item("Chemin1")
            DataListeIntegrer.Rows(i).Cells("Type1").Value = OledatableSchema.Rows(i).Item("nomtype")
            DataListeIntegrer.Rows(i).Cells("UserSage1").Value = OledatableSchema.Rows(i).Item("UserSage")
            DataListeIntegrer.Rows(i).Cells("PasseSage1").Value = OledatableSchema.Rows(i).Item("PasseSage")
            DataListeIntegrer.Rows(i).Cells("bdd1").Value = OledatableSchema.Rows(i).Item("BaseDonnee")
            DataListeIntegrer.Rows(i).Cells("Serveur1").Value = OledatableSchema.Rows(i).Item("Serveur")
            DataListeIntegrer.Rows(i).Cells("NomUtil").Value = OledatableSchema.Rows(i).Item("NomUser")
            DataListeIntegrer.Rows(i).Cells("Mot").Value = OledatableSchema.Rows(i).Item("MotPas")
            If TachePlanifie = "Export Commande Client" Then
                For Each Item As String In SocieteCyble
                    If Item = OledatableSchema.Rows(i).Item("Societe") Then
                        DataListeIntegrer.Rows(i).Cells("choix").Value = True
                    End If
                Next
            End If
        Next i
    End Sub
    Public Sub BtnModif_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnModif.Click
        Dim Etat As Boolean = False
        nbreligne = DataListeIntegrer.Rows.Count
        ListBox.Items.Clear()
        ContinuTraitement = 0
        'NbInfosLibre = 0
        'NbInfosLibreVue = 0
        lblsms.Text = "0/0"
        lblLignes.Text = "0/0"
        lblInfos.Visible = False
        ListeChampSageLigne = ""
        ip = 0
        selectindex = ""
        selectid = Nothing
        Try
            For i As Integer = 0 To DataListeIntegrer.Rows.Count - 1
                If DataListeIntegrer.Rows(i).Cells("Choix").Value = True Then
                    selectindex &= DataListeIntegrer.Rows(i).Index & ";"
                    Etat = True
                End If
                DataListeIntegrer.Rows(i).Cells("Status").Value = My.Resources.btFermer22
            Next
            If Etat = True Then
                selectindex = selectindex.Substring(0, selectindex.Length - 1)
                selectid = selectindex.Split(";")
                VerificationChampEnteteObligatoire(True, False)
                If ContinuTraitement = 4 Then
                    VerificateurChampObligatoire(False, True)
                    If ContinuTraitement = 7 Then
                        EstInfosLibre(True, False, True)
                        If NbInfosLibre = NbInfosLibreVue Then 'NbInfosLibreVue
                            EstInfosLibre(False, True, True)
                            If NbInfosLibre = NbInfosLibreVue Then 'NbInfosLibreVue
                                Try
                                    'PictureBox1.Visible = True
                                    ListeChampSageEntete = RecuperationColonneSage(True, False)
                                    lblSne.Text = "Scénario Extraction des BC Clients"
                                    ListeChampSageEntete = ListeChampSageEntete.Substring(0, ListeChampSageEntete.Length - 1)
                                    ListeChampSageLigne = RecuperationColonneSage(False, True)
                                    ListeChampSageLigne = ListeChampSageLigne.Substring(0, ListeChampSageLigne.Length - 1)
                                    BackgroundWorker1.RunWorkerAsync()
                                Catch ex As Exception
                                    MsgBox(ex.Message, MsgBoxStyle.Critical, "Traitement")
                                End Try
                            End If
                        End If
                    End If
                End If
            Else
                lblInfos.Text = "Aucune Société Selectionnée !"
                lblInfos.Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Traitement Fin Erreur")
        End Try
    End Sub
    Public OleAdaptaterschemaVersion As OleDbDataAdapter
    Public OleSchemaDatasetVersion As DataSet
    Public OledatableSchemaVersion As DataTable

    Public Sub FrmExtraction_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            LirefichierConfig()
            Me.WindowState = FormWindowState.Maximized
            If Connected() = True Then
                If TachePlanifie <> "" Then
                    AfficheSchemasTachePlanifie()
                Else
                    BackgroundWorker4.RunWorkerAsync()
                End If
                If DO_StatutClient <> "" Then
                    CmbStatut.SelectedIndex = DO_StatutClient
                    StatutDoc = DO_StatutClient
                End If
                OleAdaptaterschemaVersion = New OleDbDataAdapter("select version from P_TABLECORRESP WHERE CodeTbls='SOR'", OleConnenectionArticle)
                OleSchemaDatasetVersion = New DataSet
                OleAdaptaterschemaVersion.Fill(OleSchemaDatasetVersion)
                OledatableSchemaVersion = OleSchemaDatasetVersion.Tables(0)
            End If
        Catch ex As Exception
        End Try
    End Sub
    Public ContinuTraitement As Integer = 0
    Public Sub VerificationChampEnteteObligatoire(ByVal Entete As Boolean, ByVal Ligne As Boolean)
        OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='SOR' AND Entete=" & Entete & " AND Ligne=" & Ligne & "  ORDER BY ORDRE", OleConnenection)
        OleSchemaDataset = New DataSet
        OleAdaptaterschema.Fill(OleSchemaDataset)
        OledatableSchema = OleSchemaDataset.Tables(0)
        lblSne.Text = "Verification conformite des BC clients"
        ListBox.Items.Add("<-----------------------------------------------------------------DEBUT DU TRAITEMENT DE LA VERIFICATION DES CHAMPS OBLIGATOIRES----------------------------------------------------------------->")
        If OledatableSchema.Rows.Count <> 0 Then
            Try
                For i As Integer = 0 To OledatableSchema.Rows.Count - 1
                    Select Case Trim(OledatableSchema.Rows(i).Item("Cols"))
                        Case "SORDER_CODE"
                            If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False And OledatableSchema.Rows(i).Item("ChampSage").ToString <> "" Then
                                ContinuTraitement += 1
                                ListBox.Items.Add("Un Mapping Sage {" & OledatableSchema.Rows(i).Item("ChampSage") & "} a été appliquer sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}*")
                            Else
                                If IsDBNull(OledatableSchema.Rows(i).Item("DefaultValue")) = False Then
                                    ListBox.Items.Add("Une Valeur par defaut a été paramètrée sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}*")
                                    ContinuTraitement += 1
                                Else
                                    ListBox.Items.Add("Aucun Mapping Sage n'est paramètre sur  le champ {" & OledatableSchema.Rows(i).Item("Cols") & "}* EasyWMS  ni de valeur par defaut n'a ete paramètrée {Ce Champ est Obligatoire}")
                                End If
                            End If
                        Case "WAREHOUSE_CODE"
                            If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False And OledatableSchema.Rows(i).Item("ChampSage").ToString <> "" Then
                                ContinuTraitement += 1
                                ListBox.Items.Add("Un Mapping Sage {" & OledatableSchema.Rows(i).Item("ChampSage") & "} a été appliquer sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}*")
                            Else
                                If IsDBNull(OledatableSchema.Rows(i).Item("DefaultValue")) = False Then
                                    ListBox.Items.Add("Une Valeur par defaut a été paramètrée sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}*")
                                    ContinuTraitement += 1
                                Else
                                    ListBox.Items.Add("Aucun Mapping Sage n'est paramètre sur  le champ {" & OledatableSchema.Rows(i).Item("Cols") & "*} EasyWMS  ni de valeur par defaut n'a ete paramètrée {Ce Champ est Obligatoire}")
                                End If
                            End If
                        Case "TYPE_CODE"
                            If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False And OledatableSchema.Rows(i).Item("ChampSage").ToString <> "" Then
                                ContinuTraitement += 1
                                ListBox.Items.Add("Un Mapping Sage {" & OledatableSchema.Rows(i).Item("ChampSage") & "} a été appliquer sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}*")
                            Else
                                If IsDBNull(OledatableSchema.Rows(i).Item("DefaultValue")) = False Then
                                    ListBox.Items.Add("Une Valeur par defaut a été paramètrée sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}*")
                                    ContinuTraitement += 1
                                Else
                                    ListBox.Items.Add("Aucun Mapping Sage n'est paramètre sur  le champ {" & OledatableSchema.Rows(i).Item("Cols") & "}* EasyWMS  ni de valeur par defaut n'a ete paramètrée {Ce Champ est Obligatoire}")
                                End If
                            End If
                        Case "LINE_COUNT"
                            If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False And OledatableSchema.Rows(i).Item("ChampSage").ToString <> "" Then
                                ContinuTraitement += 1
                                ListBox.Items.Add("Un Mapping Sage {" & OledatableSchema.Rows(i).Item("ChampSage") & "} a été appliquer sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}*")
                            Else
                                If IsDBNull(OledatableSchema.Rows(i).Item("DefaultValue")) = False Then
                                    ListBox.Items.Add("Une Valeur par defaut a été paramètrée sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}*")
                                    ContinuTraitement += 1
                                Else
                                    ListBox.Items.Add("Aucun Mapping Sage n'est paramètre sur  le champ {" & OledatableSchema.Rows(i).Item("Cols") & "}* EasyWMS  ni de valeur par defaut n'a ete paramètrée {Ce Champ est Obligatoire}")
                                End If
                            End If
                    End Select
                Next i
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Verification champs Obligatoire")
            End Try
        Else
            ListBox.Items.Add("Aucun résultat n'est fourni pour ce Scénario de traitement")
        End If
        ListBox.Items.Add("<-----------------------------------------------------------------Fin----------------------------------------------------------------->")
        ListBox.Items.Add("")
    End Sub

    Public Sub VerificateurChampObligatoire(ByVal Entete As Boolean, ByVal Ligne As Boolean)
        ContinuTraitement = 0
        OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='SOR' AND Entete=" & Entete & " AND Ligne=" & Ligne & "  ORDER BY ORDRE", OleConnenection)
        OleSchemaDataset = New DataSet
        OleAdaptaterschema.Fill(OleSchemaDataset)
        OledatableSchema = OleSchemaDataset.Tables(0)
        lblSne.Text = "Verification conformite des BC clients"
        ListBox.Items.Add("<-----------------------------------------------------------------DEBUT DU TRAITEMENT DE LA VERIFICATION DES CHAMPS OBLIGATOIRES----------------------------------------------------------------->")
        If OledatableSchema.Rows.Count <> 0 Then
            For i As Integer = 0 To OledatableSchema.Rows.Count - 1
                Select Case Trim(OledatableSchema.Rows(i).Item("Cols"))
                    Case "LINE_NUMBER"
                        If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False And OledatableSchema.Rows(i).Item("ChampSage").ToString <> "" Then
                            ContinuTraitement += 1
                            ListBox.Items.Add("Un Mapping Sage {" & OledatableSchema.Rows(i).Item("ChampSage") & "} a été appliquer sur  le champ EasyWMS {" & Trim(OledatableSchema.Rows(i).Item("Cols")) & "}*")
                        Else
                            If IsDBNull(OledatableSchema.Rows(i).Item("DefaultValue")) = False Then
                                ListBox.Items.Add("Une Valeur par defaut a été paramètrée sur  le champ EasyWMS {" & Trim(OledatableSchema.Rows(i).Item("Cols")) & "}*")
                                ContinuTraitement += 1
                            Else
                                ListBox.Items.Add("Aucun Mapping Sage n'est paramètre sur  le champ {" & Trim(OledatableSchema.Rows(i).Item("Cols")) & "}* EasyWMS  ni de valeur par defaut n'a ete paramètrée {Ce Champ est Obligatoire}")
                            End If
                        End If
                    Case "PRODUCT_CODE"
                        If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False And OledatableSchema.Rows(i).Item("ChampSage").ToString <> "" Then
                            ContinuTraitement += 1
                            ListBox.Items.Add("Un Mapping Sage {" & OledatableSchema.Rows(i).Item("ChampSage") & "} a été appliquer sur  le champ EasyWMS {" & Trim(OledatableSchema.Rows(i).Item("Cols")) & "}*")
                        Else
                            If IsDBNull(OledatableSchema.Rows(i).Item("DefaultValue")) = False Then
                                ListBox.Items.Add("Une Valeur par defaut a été paramètrée sur  le champ EasyWMS {" & Trim(OledatableSchema.Rows(i).Item("Cols")) & "}*")
                                ContinuTraitement += 1
                            Else
                                ListBox.Items.Add("Aucun Mapping Sage n'est paramètre sur  le champ {" & Trim(OledatableSchema.Rows(i).Item("Cols")) & "*} EasyWMS  ni de valeur par defaut n'a ete paramètrée {Ce Champ est Obligatoire}")
                            End If
                        End If
                    Case "QUANTITY"
                        If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False And OledatableSchema.Rows(i).Item("ChampSage").ToString <> "" Then
                            ContinuTraitement += 1
                            ListBox.Items.Add("Un Mapping Sage {" & OledatableSchema.Rows(i).Item("ChampSage") & "} a été appliquer sur  le champ EasyWMS {" & Trim(OledatableSchema.Rows(i).Item("Cols")) & "}*")
                        Else
                            If IsDBNull(OledatableSchema.Rows(i).Item("DefaultValue")) = False Then
                                ListBox.Items.Add("Une Valeur par defaut a été paramètrée sur  le champ EasyWMS {" & Trim(OledatableSchema.Rows(i).Item("Cols")) & "}*")
                                ContinuTraitement += 1
                            Else
                                ListBox.Items.Add("Aucun Mapping Sage n'est paramètre sur  le champ {" & Trim(OledatableSchema.Rows(i).Item("Cols")) & "}* EasyWMS  ni de valeur par defaut n'a ete paramètrée {Ce Champ est Obligatoire}")
                            End If
                        End If
                    Case "UOM_CODE"
                        If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False And OledatableSchema.Rows(i).Item("ChampSage").ToString <> "" Then
                            ContinuTraitement += 1
                            ListBox.Items.Add("Un Mapping Sage {" & OledatableSchema.Rows(i).Item("ChampSage") & "} a été appliquer sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}*")
                        Else
                            If IsDBNull(OledatableSchema.Rows(i).Item("DefaultValue")) = False Then
                                ListBox.Items.Add("Une Valeur par defaut a été paramètrée sur  le champ EasyWMS {" & Trim(OledatableSchema.Rows(i).Item("Cols")) & "}*")
                                ContinuTraitement += 1
                            Else
                                ListBox.Items.Add("Aucun Mapping Sage n'est paramètre sur  le champ {" & Trim(OledatableSchema.Rows(i).Item("Cols")) & "}* EasyWMS  ni de valeur par defaut n'a ete paramètrée {Ce Champ est Obligatoire}")
                            End If
                        End If
                    Case "CATCH_WEIGHT"
                        If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False And OledatableSchema.Rows(i).Item("ChampSage").ToString <> "" Then
                            ContinuTraitement += 1
                            ListBox.Items.Add("Un Mapping Sage {" & OledatableSchema.Rows(i).Item("ChampSage") & "} a été appliquer sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}*")
                        Else
                            If IsDBNull(OledatableSchema.Rows(i).Item("DefaultValue")) = False Then
                                ListBox.Items.Add("Une Valeur par defaut a été paramètrée sur  le champ EasyWMS {" & Trim(OledatableSchema.Rows(i).Item("Cols")) & "}*")
                                ContinuTraitement += 1
                            Else
                                ListBox.Items.Add("Aucun Mapping Sage n'est paramètre sur  le champ {" & Trim(OledatableSchema.Rows(i).Item("Cols")) & "}* EasyWMS  ni de valeur par defaut n'a ete paramètrée {Ce Champ est Obligatoire}")
                            End If
                        End If
                    Case "SUSTITUTE_QTTY"
                        If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False And OledatableSchema.Rows(i).Item("ChampSage").ToString <> "" Then
                            ContinuTraitement += 1
                            ListBox.Items.Add("Un Mapping Sage {" & OledatableSchema.Rows(i).Item("ChampSage") & "} a été appliquer sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}*")
                        Else
                            If IsDBNull(OledatableSchema.Rows(i).Item("DefaultValue")) = False Then
                                ListBox.Items.Add("Une Valeur par defaut a été paramètrée sur  le champ EasyWMS {" & Trim(OledatableSchema.Rows(i).Item("Cols")) & "}*")
                                ContinuTraitement += 1
                            Else
                                ListBox.Items.Add("Aucun Mapping Sage n'est paramètre sur  le champ {" & Trim(OledatableSchema.Rows(i).Item("Cols")) & "}* EasyWMS  ni de valeur par defaut n'a ete paramètrée {Ce Champ est Obligatoire}")
                            End If
                        End If
                    Case "DISABLE_ALT_PRODUCT"
                        If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False And OledatableSchema.Rows(i).Item("ChampSage").ToString <> "" Then
                            ContinuTraitement += 1
                            ListBox.Items.Add("Un Mapping Sage {" & OledatableSchema.Rows(i).Item("ChampSage") & "} a été appliquer sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}*")
                        Else
                            If IsDBNull(OledatableSchema.Rows(i).Item("DefaultValue")) = False Then
                                ListBox.Items.Add("Une Valeur par defaut a été paramètrée sur  le champ EasyWMS {" & Trim(OledatableSchema.Rows(i).Item("Cols")) & "}*")
                                ContinuTraitement += 1
                            Else
                                ListBox.Items.Add("Aucun Mapping Sage n'est paramètre sur  le champ {" & Trim(OledatableSchema.Rows(i).Item("Cols")) & "}* EasyWMS  ni de valeur par defaut n'a ete paramètrée {Ce Champ est Obligatoire}")
                            End If
                        End If

                End Select
            Next i
        Else
            ListBox.Items.Add("Aucun résultat n'est fourni pour ce Scénario de traitement")
        End If
        ListBox.Items.Add("<-----------------------------------------------------------------Fin----------------------------------------------------------------->")
        ListBox.Items.Add("")
    End Sub
    Public Function RecuperationColonneSage(ByVal Entete As Boolean, ByVal Ligne As Boolean) As String
        'lblentete.Text = ""
        lblSne.Text = "Scenario de recuperation des champs Sage"
        If Entete Then
            ListBox.Items.Add("<-----------------------------------------------------------------DEBUT DU TRAITEMENT DE LA RECUPERATION  DES ENTETE DES CHAMPS  MAPPES----------------------------------------------------------------->")
        Else
            ListBox.Items.Add("<-----------------------------------------------------------------DEBUT DU TRAITEMENT DE LA RECUPERATION  DES LIGNES DES CHAMPS MAPPES----------------------------------------------------------------->")
        End If
        Dim NbEntete As Integer = 0
        Dim NbLigne As Integer = 0
        Dim NbInfosLibres As Integer = 0
        Dim OleAdaptaterschema As OleDbDataAdapter
        Dim OleSchemaDataset As DataSet
        Dim OledatableSchema As DataTable
        Dim ListeChampSage As String = ""
        If Entete Then
            OleAdaptaterschema = New OleDbDataAdapter("SELECT * from P_COLONNEST WHERE CodeTbls='SOR' AND (Entete=" & Entete & " OR InfosLibre=true) AND Ligne=False ORDER BY ORDRE", OleConnenection)
        Else
            OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='SOR' AND (Ligne=" & Ligne & " OR InfosLibre=true) AND Entete=False ORDER BY ORDRE", OleConnenection)
        End If
        OleSchemaDataset = New DataSet
        OleAdaptaterschema.Fill(OleSchemaDataset)
        OledatableSchema = OleSchemaDataset.Tables(0)
        If OledatableSchema.Rows.Count <> 0 Then
            For i As Integer = 0 To OledatableSchema.Rows.Count - 1
                If OledatableSchema.Rows(i).Item("Entete") = "true" And OledatableSchema.Rows(i).Item("InfosLibre") = "true" Then
                    NbInfosLibres += 1
                    lblinfosLibre.Text = NbInfosLibres
                ElseIf OledatableSchema.Rows(i).Item("Entete") = "true" And OledatableSchema.Rows(i).Item("InfosLibre") = "false" Then
                    If OledatableSchema.Rows(i).Item("Entete") Then
                        NbEntete += 1
                        lblentete.Text = NbEntete
                    End If
                ElseIf OledatableSchema.Rows(i).Item("Ligne") = "true" And OledatableSchema.Rows(i).Item("InfosLibre") = "true" Then
                    NbInfosLibres += 1
                    lblinfosLibre.Text = NbInfosLibres
                ElseIf OledatableSchema.Rows(i).Item("Ligne") = "true" And OledatableSchema.Rows(i).Item("InfosLibre") = "false" Then
                    If OledatableSchema.Rows(i).Item("Ligne") Then
                        NbLigne += 1
                        lblligne.Text = NbLigne
                    End If
                Else
                End If
                Dim ExisteColonne As Object = 0
                If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False Then
                    If ExcelConnect(NomServersql, NomBaseCpta, Nom_Utilsql, Mot_Passql) Then
                        Dim LaRequeteVerificationColonne As String = ""
                        If Entete Then
                            LaRequeteVerificationColonne = "SELECT COUNT(*) FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='F_DOCENTETE' AND COLUMN_NAME = '" & OledatableSchema.Rows(i).Item("ChampSage").ToString & "'"

                        Else
                            LaRequeteVerificationColonne = "SELECT COUNT(*) FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='F_DOCLIGNE' AND COLUMN_NAME = '" & OledatableSchema.Rows(i).Item("ChampSage").ToString & "'"

                        End If
                        Dim MaCommande As New OleDbCommand(LaRequeteVerificationColonne, OleExcelConnect)
                        ExisteColonne = MaCommande.ExecuteScalar
                    End If
                End If

                If ExisteColonne <> 0 Then
                    If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False Then
                        If Entete Then
                            If i = OledatableSchema.Rows.Count - 1 Then
                                If OledatableSchema.Rows(i).Item("ChampSage") <> "" Then
                                    ListeChampSage &= OledatableSchema.Rows(i).Item("ChampSage")
                                    ListBox.Items.Add("Recuperation de la colonne en Entete : " & OledatableSchema.Rows(i).Item("ChampSage"))
                                Else
                                    ListBox.Items.Add("Le Champ {" & OledatableSchema.Rows(i).Item("Cols") & "} EasyWMS ne possede pas de Mapping Sage infos sur le Bon de Commande du Client ")
                                End If
                            Else
                                If OledatableSchema.Rows(i).Item("ChampSage") <> "" Then
                                    ListeChampSage &= OledatableSchema.Rows(i).Item("ChampSage") & ","
                                    ListBox.Items.Add("Recuperation de la colonne en Entete : " & OledatableSchema.Rows(i).Item("ChampSage"))
                                Else
                                    ListBox.Items.Add("Le Champ {" & OledatableSchema.Rows(i).Item("Cols") & "} EasyWMS ne possede pas de Mapping Sage infos sur le Bon de Commande du Client")
                                End If
                            End If
                        Else
                            If i = OledatableSchema.Rows.Count - 1 Then
                                If OledatableSchema.Rows(i).Item("ChampSage") <> "" Then
                                    ListeChampSageLigne &= OledatableSchema.Rows(i).Item("ChampSage")
                                    ListBox.Items.Add("Recuperation de la colonne en Ligne : " & OledatableSchema.Rows(i).Item("ChampSage"))
                                Else
                                    ListBox.Items.Add("Le Champ {" & OledatableSchema.Rows(i).Item("Cols") & "} EasyWMS ne possede pas de Mapping Sage infos sur le Bon de Commande du Client")
                                End If
                            Else
                                If OledatableSchema.Rows(i).Item("ChampSage") <> "" Then
                                    ListeChampSageLigne &= OledatableSchema.Rows(i).Item("ChampSage") & ","
                                    ListBox.Items.Add("Recuperation de la colonne en Ligne : " & OledatableSchema.Rows(i).Item("ChampSage"))
                                Else
                                    ListBox.Items.Add("Le Champ {" & OledatableSchema.Rows(i).Item("Cols") & "} EasyWMS ne possede pas de Mapping Sage infos sur le Bon de Commande du Client")
                                End If
                            End If
                        End If
                    Else
                        ListBox.Items.Add("<Le Champ {" & OledatableSchema.Rows(i).Item("Cols") & "} EasyWMS ne possede pas de Mapping Sage>")
                    End If
                End If
            Next i
        End If
        ListBox.Items.Add("<-----------------------------------------------------------------Fin----------------------------------------------------------------->")
        ListBox.Items.Add("")
        If Entete Then
            Return ListeChampSage
        Else
            Return ListeChampSageLigne
        End If
    End Function
    Public Function RecuperationColonneSageLigne(ByVal Entete As Boolean, ByVal Ligne As Boolean) As String
        'lblentete.Text = ""
        lblSne.Text = "Scenario de recuperation des champs Sage"
        ListBox.Items.Add("<-----------------------------------------------------------------DEBUT DU TRAITEMENT DE LA RECUPERATION  DES CHAMPS MAPPES----------------------------------------------------------------->")
        Dim NbEntete As Integer = 0
        Dim NbLigne As Integer = 0
        Dim NbInfosLibres As Integer = 0
        Dim OleAdaptaterschema As OleDbDataAdapter
        Dim OleSchemaDataset As DataSet
        Dim OledatableSchema As DataTable
        Dim ListeChampSage As String = ""
        If Entete Then
            OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='SOR' AND (Entete=" & Entete & " OR InfosLibre=true) AND Ligne=False ORDER BY ORDRE", OleConnenection)
        Else
            OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='SOR' AND (Ligne=" & Ligne & " OR InfosLibre=true) AND Entete=False ORDER BY ORDRE", OleConnenection)
        End If
        OleSchemaDataset = New DataSet
        OleAdaptaterschema.Fill(OleSchemaDataset)
        OledatableSchema = OleSchemaDataset.Tables(0)
        If OledatableSchema.Rows.Count <> 0 Then
            For i As Integer = 0 To OledatableSchema.Rows.Count - 1
                If OledatableSchema.Rows(i).Item("Entete") = "true" And OledatableSchema.Rows(i).Item("InfosLibre") = "true" Then
                    NbInfosLibres += 1
                    lblinfosLibre.Text = NbInfosLibres
                ElseIf OledatableSchema.Rows(i).Item("Entete") = "true" And OledatableSchema.Rows(i).Item("InfosLibre") = "false" Then
                    If OledatableSchema.Rows(i).Item("Entete") Then
                        NbEntete += 1
                        lblentete.Text = NbEntete
                    End If
                ElseIf OledatableSchema.Rows(i).Item("Ligne") = "true" And OledatableSchema.Rows(i).Item("InfosLibre") = "true" Then
                    NbInfosLibres += 1
                    lblinfosLibre.Text = NbInfosLibres
                ElseIf OledatableSchema.Rows(i).Item("Entete") = "true" And OledatableSchema.Rows(i).Item("InfosLibre") = "false" Then
                    If OledatableSchema.Rows(i).Item("Ligne") Then
                        NbLigne += 1
                        lblligne.Text = NbLigne
                    End If
                Else
                End If
                Dim ExisteColonne As Object = 0
                If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False Then
                    If ExcelConnect(NomServersql, NomBaseCpta, Nom_Utilsql, Mot_Passql) Then
                        Dim LaRequeteVerificationColonne As String = "SELECT COUNT(*) FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='F_COMPTET' AND COLUMN_NAME = '" & OledatableSchema.Rows(i).Item("ChampSage").ToString & "'"
                        Dim MaCommande As New OleDbCommand(LaRequeteVerificationColonne, OleExcelConnect)
                        ExisteColonne = MaCommande.ExecuteScalar
                    End If
                End If

                If ExisteColonne <> 0 Then
                    If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False Then
                        If Entete Then
                            If i = OledatableSchema.Rows.Count - 1 Then
                                If OledatableSchema.Rows(i).Item("ChampSage") <> "" Then
                                    ListeChampSage &= OledatableSchema.Rows(i).Item("ChampSage")
                                    ListBox.Items.Add("Recuperation de la colonne en Entete : " & OledatableSchema.Rows(i).Item("ChampSage"))
                                Else
                                    ListBox.Items.Add("Le Champ {" & OledatableSchema.Rows(i).Item("Cols") & "} EasyWMS ne possede pas de Mapping Sage infos sur le Bon de Commande du Client ")
                                End If
                            Else
                                If OledatableSchema.Rows(i).Item("ChampSage") <> "" Then
                                    ListeChampSage &= OledatableSchema.Rows(i).Item("ChampSage") & ","
                                    ListBox.Items.Add("Recuperation de la colonne en Entete : " & OledatableSchema.Rows(i).Item("ChampSage"))
                                Else
                                    ListBox.Items.Add("Le Champ {" & OledatableSchema.Rows(i).Item("Cols") & "} EasyWMS ne possede pas de Mapping Sage infos sur le Bon de Commande du Client")
                                End If
                            End If
                        Else
                            If i = OledatableSchema.Rows.Count - 1 Then
                                If OledatableSchema.Rows(i).Item("ChampSage") <> "" Then
                                    ListeChampSage &= OledatableSchema.Rows(i).Item("ChampSage")
                                    ListBox.Items.Add("Recuperation de la colonne en Ligne : " & OledatableSchema.Rows(i).Item("ChampSage"))
                                Else
                                    ListBox.Items.Add("Le Champ {" & OledatableSchema.Rows(i).Item("Cols") & "} EasyWMS ne possede pas de Mapping Sage infos sur le Bon de Commande du Client")
                                End If
                            Else
                                If OledatableSchema.Rows(i).Item("ChampSage") <> "" Then
                                    ListeChampSage &= OledatableSchema.Rows(i).Item("ChampSage") & ","
                                    ListBox.Items.Add("Recuperation de la colonne en Ligne : " & OledatableSchema.Rows(i).Item("ChampSage"))
                                Else
                                    ListBox.Items.Add("Le Champ {" & OledatableSchema.Rows(i).Item("Cols") & "} EasyWMS ne possede pas de Mapping Sage infos sur le Bon de Commande du Client")
                                End If
                            End If
                        End If
                    Else
                        ListBox.Items.Add("<Le Champ {" & OledatableSchema.Rows(i).Item("Cols") & "} EasyWMS ne possede pas de Mapping Sage>")
                    End If
                End If

            Next i
        End If
        ListBox.Items.Add("<-----------------------------------------------------------------Fin----------------------------------------------------------------->")
        ListBox.Items.Add("")
        Return ListeChampSage
    End Function
    Public Sub EstInfosLibre(ByVal Entete As Boolean, ByVal Ligne As Boolean, ByVal InfosLibre As Boolean)
        Dim OleAdaptaterschema As OleDbDataAdapter
        Dim OleSchemaDataset As DataSet
        Dim OledatableSchema As DataTable
        Dim ListeInfosLibre As String = ""
        NbInfosLibreVue = 0
        lblSne.Text = "Scenario verification infos libre"
        Try
            OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='SOR' AND Entete=" & Entete & " AND InfosLibre=" & InfosLibre & " AND Ligne=" & Ligne, OleConnenection)
            OleSchemaDataset = New DataSet
            OleAdaptaterschema.Fill(OleSchemaDataset)
            OledatableSchema = OleSchemaDataset.Tables(0)
            If OledatableSchema.Rows.Count <> 0 Then
                NbInfosLibre = OledatableSchema.Rows.Count
                ListBox.Items.Add("<-----------------------------------------------------------------DEBUT DU TRAITEMENT DE LA VERIFICATION DES INFOS LIBRES----------------------------------------------------------------->")
                For i As Integer = 0 To OledatableSchema.Rows.Count - 1
                    If OledatableSchema.Rows(i).Item("InfosLibre") = "true" Then
                        If Entete = True Then
                            If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False Then
                                If ExisteInfosLibre(OledatableSchema.Rows(i).Item("ChampSage"), Entete) = False Then
                                    ListBox.Items.Add("la zone infos libre {" & OledatableSchema.Rows(i).Item("ChampSage") & "} en Entete du Bon de Commande Client n'existe pas dans Sage")
                                Else
                                    ListBox.Items.Add("Traitement  de la zone infos libre {" & OledatableSchema.Rows(i).Item("ChampSage") & "} sur le Bon de Commande Client Existe [OK] dans Sage")
                                End If
                            Else
                                ListBox.Items.Add("<--Le Champ indiquant l'infos libre est couché mais ne possede pas de mapping Sage-->")
                            End If
                        ElseIf Ligne = True Then
                            If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False Then
                                If ExisteInfosLibre(OledatableSchema.Rows(i).Item("ChampSage"), Entete) = False Then
                                    ListBox.Items.Add("la zone infos libre {" & OledatableSchema.Rows(i).Item("ChampSage") & "} sur le Bon de Commande Client n'existe pas dans Sage")
                                End If
                            Else
                                ListBox.Items.Add("<--Le Champ indiquant l'infos libre est couché mais ne possede pas de mapping Sage-->")
                            End If
                        Else
                            ListBox.Items.Add("<--Aucune information libre n'est parametrée-->")
                        End If
                    End If
                Next
                ListBox.Items.Add("<-----------------------------------------------------------------Fin----------------------------------------------------------------->")
                ListBox.Items.Add("")
            End If
        Catch ex As Exception
        End Try
    End Sub
    Public Function ExisteInfosLibre(ByVal InfosLibre As String, ByVal Entete As Boolean) As Boolean
        Try
            If ExcelConnect(NomServersql, Societe, Nom_Utilsql, Mot_Passql) Then
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
    Public Erreur As Boolean = True
    Public Societe As String = ""
    Public Sub ExtractionBonCommandeClient(ByVal SelectChamp As String, ByVal indice As Integer)
        Dim Information As String = ""
        Dim ExistColonne As Boolean = False
        Dim ColonnName As String = ""
        Dim PositionLeft As Object = Nothing
        Try
            If ChFlag.Checked Then
                If ExcelConnect(DataListeIntegrer.Rows(indice).Cells("Serveur1").Value, DataListeIntegrer.Rows(indice).Cells("Societe1").Value, DataListeIntegrer.Rows(indice).Cells("NomUtil").Value, DataListeIntegrer.Rows(indice).Cells("Mot").Value) Then
                    Societe = DataListeIntegrer.Rows(indice).Cells("Societe1").Value
                    If ExisteInfosLibre(Flagtampon, True) = True Then
                        If ExcelConnect(DataListeIntegrer.Rows(indice).Cells("Serveur1").Value, DataListeIntegrer.Rows(indice).Cells("Societe1").Value, DataListeIntegrer.Rows(indice).Cells("NomUtil").Value, DataListeIntegrer.Rows(indice).Cells("Mot").Value) Then
                            If ExcelConnect(DataListeIntegrer.Rows(indice).Cells("Serveur1").Value, DataListeIntegrer.Rows(indice).Cells("Societe1").Value, DataListeIntegrer.Rows(indice).Cells("NomUtil").Value, DataListeIntegrer.Rows(indice).Cells("Mot").Value) Then
                                If SelectChamp IsNot Nothing Then
                                    If RbtVente.Checked Then
                                        OleAdaptaterschemaSage = New OleDbDataAdapter("SELECT " & SelectChamp & " FROM F_DOCENTETE WHERE DO_DOMAINE=0 AND DO_TYPE=1   AND DO_Statut=" & StatutDoc & " AND (" & Flagtampon & " IS NULL OR " & Flagtampon & "='') ORDER BY F_DOCENTETE.DO_PIECE ", OleExcelConnect)
                                    ElseIf RbtAchat.Checked Then
                                        OleAdaptaterschemaSage = New OleDbDataAdapter("SELECT " & SelectChamp & " FROM F_DOCENTETE WHERE DO_DOMAINE=1 AND DO_TYPE=14  AND DO_Statut=" & StatutDoc & " AND (" & Flagtampon & " IS NULL OR " & Flagtampon & "='')  ORDER BY F_DOCENTETE.DO_PIECE ", OleExcelConnect)
                                    ElseIf RbtStock.Checked Then
                                        OleAdaptaterschemaSage = New OleDbDataAdapter("SELECT " & SelectChamp & " FROM F_DOCENTETE WHERE DO_DOMAINE=2 AND DO_TYPE=23  AND DO_Statut=" & StatutDoc & " AND (" & Flagtampon & " IS NULL  OR " & Flagtampon & "='')  ORDER BY F_DOCENTETE.DO_PIECE ", OleExcelConnect)
                                    End If
                                Else
                                    If RbtVente.Checked Then
                                        OleAdaptaterschemaSage = New OleDbDataAdapter("SELECT * FROM F_DOCENTETE WHERE DO_DOMAINE=0 AND DO_TYPE=1   AND DO_Statut=" & StatutDoc & "  AND (" & Flagtampon & " IS NULL OR " & Flagtampon & "='') ORDER BY F_DOCENTETE.DO_PIECE ", OleExcelConnect)
                                    ElseIf RbtAchat.Checked Then
                                        OleAdaptaterschemaSage = New OleDbDataAdapter("SELECT * FROM F_DOCENTETE WHERE DO_DOMAINE=1 AND DO_TYPE=14  AND DO_Statut=" & StatutDoc & "  AND (" & Flagtampon & " IS NULL OR " & Flagtampon & "='') ORDER BY F_DOCENTETE.DO_PIECE ", OleExcelConnect)
                                    ElseIf RbtStock.Checked Then
                                        OleAdaptaterschemaSage = New OleDbDataAdapter("SELECT * FROM F_DOCENTETE WHERE DO_DOMAINE=2 AND DO_TYPE=23  AND DO_Statut=" & StatutDoc & "  AND (" & Flagtampon & " IS NULL OR " & Flagtampon & "='') ORDER BY F_DOCENTETE.DO_PIECE ", OleExcelConnect)
                                    End If
                                End If
                                OleSchemaDatasetSage = New DataSet
                                OleAdaptaterschemaSage.Fill(OleSchemaDatasetSage)
                                OledatableSchemaSage = OleSchemaDatasetSage.Tables(0)

                                OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='SOR' AND Entete=True ORDER BY ORDRE ", OleConnenection)
                                OleSchemaDataset = New DataSet
                                OleAdaptaterschema.Fill(OleSchemaDataset)
                                OledatableSchema = OleSchemaDataset.Tables(0)
                                If OledatableSchemaSage.Rows.Count <> 0 Then
                                    EstChoisir = True
                                End If
                            End If
                        End If
                    Else
                        ErreurJrn = File.AppendText(Pathsfilejournal & "ERREURSOR09" & Strings.Right(DateAndTime.Year(Now), 4) & Format(DateAndTime.Month(Now), "00") & Format(DateAndTime.Day(Now), "00") & Format(DateAndTime.Hour(Now), "00") & Format(DateAndTime.Minute(Now), "00") & Format(DateAndTime.Second(Now), "00") & ".txt")
                        ErreurJrn.WriteLine("Connexion à la Société <[" & DataListeIntegrer.Rows(indice).Cells("Societe1").Value & "]> Reussi...")
                        ErreurJrn.WriteLine("Infos Libre <[" & Flagtampon & "]> du Flagage n'existe pas dans la Table cible")
                        ErreurJrn.Close()
                        Erreur = True
                        DataListeIntegrer.Rows(selectid(ip)).Cells("Status").Value = My.Resources.criticalind_status
                        ifRowError = selectid(ip)
                    End If
                Else
                    ErreurJrn = File.AppendText(Pathsfilejournal & "ERREURSOR09" & Strings.Right(DateAndTime.Year(Now), 4) & Format(DateAndTime.Month(Now), "00") & Format(DateAndTime.Day(Now), "00") & Format(DateAndTime.Hour(Now), "00") & Format(DateAndTime.Minute(Now), "00") & Format(DateAndTime.Second(Now), "00") & ".txt")
                    ErreurJrn.WriteLine("Erreur de Connexion à la Société <[" & DataListeIntegrer.Rows(indice).Cells("Societe1").Value & "]>")
                    ErreurJrn.Close()
                    Erreur = True
                End If
            Else
                If ExcelConnect(DataListeIntegrer.Rows(indice).Cells("Serveur1").Value, DataListeIntegrer.Rows(indice).Cells("Societe1").Value, DataListeIntegrer.Rows(indice).Cells("NomUtil").Value, DataListeIntegrer.Rows(indice).Cells("Mot").Value) Then
                    Societe = DataListeIntegrer.Rows(indice).Cells("Societe1").Value
                    If SelectChamp IsNot Nothing Then
                        If RbtVente.Checked Then
                            OleAdaptaterschemaSage = New OleDbDataAdapter("SELECT " & SelectChamp & " FROM F_DOCENTETE WHERE DO_DOMAINE=0 AND DO_TYPE=1 AND DO_Statut=" & StatutDoc & "  ORDER BY F_DOCENTETE.DO_PIECE ", OleExcelConnect) 'AND CAST(DO_Date AS Date) = CAST(CBModification AS Date)
                        ElseIf RbtAchat.Checked Then
                            OleAdaptaterschemaSage = New OleDbDataAdapter("SELECT " & SelectChamp & " FROM F_DOCENTETE WHERE DO_DOMAINE=1 AND DO_TYPE=14 AND DO_Statut=" & StatutDoc & " ORDER BY F_DOCENTETE.DO_PIECE ", OleExcelConnect)
                        ElseIf RbtStock.Checked Then
                            OleAdaptaterschemaSage = New OleDbDataAdapter("SELECT " & SelectChamp & " FROM F_DOCENTETE WHERE DO_DOMAINE=2 AND DO_TYPE=23 AND DO_Statut=" & StatutDoc & " ORDER BY F_DOCENTETE.DO_PIECE ", OleExcelConnect)
                        End If
                    Else
                        If RbtVente.Checked Then
                            OleAdaptaterschemaSage = New OleDbDataAdapter("SELECT * FROM F_DOCENTETE WHERE DO_DOMAINE=0 AND DO_TYPE=1 AND DO_Statut=" & StatutDoc & "  ORDER BY F_DOCENTETE.DO_PIECE ", OleExcelConnect)
                        ElseIf RbtAchat.Checked Then
                            OleAdaptaterschemaSage = New OleDbDataAdapter("SELECT * FROM F_DOCENTETE WHERE DO_DOMAINE=1 AND DO_TYPE=14 AND DO_Statut=" & StatutDoc & " ORDER BY F_DOCENTETE.DO_PIECE ", OleExcelConnect)
                        ElseIf RbtStock.Checked Then
                            OleAdaptaterschemaSage = New OleDbDataAdapter("SELECT * FROM F_DOCENTETE WHERE DO_DOMAINE=2 AND DO_TYPE=23 AND DO_Statut=" & StatutDoc & " ORDER BY F_DOCENTETE.DO_PIECE ", OleExcelConnect)
                        End If
                    End If
                    OleSchemaDatasetSage = New DataSet
                    OleAdaptaterschemaSage.Fill(OleSchemaDatasetSage)
                    OledatableSchemaSage = OleSchemaDatasetSage.Tables(0)

                    OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='SOR' AND Entete=True ORDER BY ORDRE ", OleConnenection)
                    OleSchemaDataset = New DataSet
                    OleAdaptaterschema.Fill(OleSchemaDataset)
                    OledatableSchema = OleSchemaDataset.Tables(0)
                    If OledatableSchemaSage.Rows.Count <> 0 Then
                        EstChoisir = True
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Chargement des Bon de Commande Client")
        End Try
    End Sub
    Dim ifRowError = 0
    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            ExtractionBonCommandeClient(ListeChampSageEntete, selectid(ip))
        Catch ex As Exception
            ifRowError = selectid(ip)
            ErreurJrn = File.AppendText(Pathsfilejournal & "EXTRACTIONACC09" & Strings.Right(DateAndTime.Year(Now), 4) & Format(DateAndTime.Month(Now), "00") & Format(DateAndTime.Day(Now), "00") & Format(DateAndTime.Hour(Now), "00") & Format(DateAndTime.Minute(Now), "00") & Format(DateAndTime.Second(Now), "00") & ".txt")
            ErreurJrn.WriteLine("ERREUR SYSTEME : " & ex.Message)
            ErreurJrn.Close()
        End Try
    End Sub
    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        BackgroundWorker2.RunWorkerAsync()
    End Sub
    Public Function CounteLigne(ByVal code As String) As Integer
        Try
            Dim Query As String = ""
            If RbtVente.Checked Then
                Query = "SELECT count(*) FROM F_DOCLIGNE WHERE F_DOCLIGNE.DO_DOMAINE=0 AND F_DOCLIGNE.DO_TYPE=1 AND F_DOCLIGNE.DO_PIECE='" & code & "'  AND AR_REF IS NOT NULL "
            ElseIf RbtAchat.Checked Then
                Query = "SELECT count(*) FROM F_DOCLIGNE WHERE F_DOCLIGNE.DO_DOMAINE=1 AND F_DOCLIGNE.DO_TYPE=14 AND F_DOCLIGNE.DO_PIECE='" & code & "' AND AR_REF IS NOT NULL "
            ElseIf RbtStock.Checked Then
                Query = "SELECT count(*) FROM F_DOCLIGNE WHERE F_DOCLIGNE.DO_DOMAINE=2 AND F_DOCLIGNE.DO_TYPE=23 AND F_DOCLIGNE.DO_PIECE='" & code & "' AND AR_REF IS NOT NULL "
            End If
            Dim MaCommande As New OleDbCommand(Query, OleExcelConnect)
            Return CInt(MaCommande.ExecuteScalar())
        Catch ex As Exception
        End Try
    End Function
    Public EstVide As Boolean = False
    Delegate Sub Evenement()
    Public Element As String
    Public Element2 As String
    Public Element3 As String
    Public cpteur As Integer
    Public Sub Traitement1()
        ListBox.Items.Add("Extraction " & Item & " Identificateur (Code) : <[" & Element & "]> ")
        lblsms.Text = cpteur + 1 & "/" & OledatableSchemaSage.Rows.Count
    End Sub
    Public Sub Traitement2()
        lblLignes.Text = nbrLigneVue & "/" & OledatableSchemaSageDetails.Rows.Count
        ListBox.Items.Add("     0---> Detail du Document N° Piece : " & Element2 & " Ligne du Document Extrait :[" & Element3 & "]")
        ListBox.SelectedIndex = ListBox.Items.Count - 1
    End Sub
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
    Public Function EcritureEntete() As Boolean
        Dim Information As String = ""
        Dim ExistColonne As Boolean = False
        Dim ColonnName As String = ""
        Dim PositionLeft As Object = Nothing
        Dim PositionLeftLigne As Integer = 0
        Dim ColonnNameLigne As String = ""
        Dim statut As Boolean = False
        Dim version As String = "09"
        cpteur = 0
        Try
            If OledatableSchemaSage.Rows.Count <> 0 Then
                version = OledatableSchemaVersion.Rows(0).Item("version")
                Error_journal = File.AppendText(PathsFileFormatiers & "SOR" & version & Strings.Right(DateAndTime.Year(Now), 4) & Format(DateAndTime.Month(Now), "00") & Format(DateAndTime.Day(Now), "00") & Format(DateAndTime.Hour(Now), "00") & Format(DateAndTime.Minute(Now), "00") & Format(DateAndTime.Second(Now), "00") & ".txt")
                For i As Integer = 0 To OledatableSchemaSage.Rows.Count - 1
                    Element = OledatableSchemaSage.Rows(i).Item("DO_PIECE")
                    cpteur = i
                    If ListBox.InvokeRequired Then
                        Dim MonDelegate As New Evenement(AddressOf Traitement1)
                        ListBox.Invoke(MonDelegate)
                    Else
                        Traitement1()
                    End If
                    For Ligne As Integer = 0 To OledatableSchema.Rows.Count - 1
                        '-------------------------------------------------------------------------------------
                        Dim CARACTERE As String = OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0).ToString
                        If CARACTERE <> "V" Then
                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                        End If
                        Select Case OledatableSchema.Rows(Ligne).Item("Cols") '
                            Case "OPÉRATION"
                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                Else
                                    Information &= Strings.LSet("F", PositionLeft)
                                End If
                                Continue For
                            Case "SORDER_CODE"
                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                        ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans à la variable colonne
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, PositionLeft)
                                        End If
                                        Continue For
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        Else
                                            'un Mapping Sage existe mais ne possede pas de valeur que retourne la requete et ne possede pas de valeur pas defaut
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = ""
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        End If
                                    End If
                                Else 'si le champ Sage ne possede pas de mapping correspondant a la colonne EasyWMS
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut paramètrée
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                            Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else 'sinon le format n'est pas du type variable
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace à la fonction GetPosition qui prend en paramétre le format du champ
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        End If
                                    Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                        'Dim traitementAutres As String = 0
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                            Dim ValeurLue = ""
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet("", PositionLeft)
                                            Else
                                                Information &= Strings.LSet("", PositionLeft)
                                            End If
                                        End If
                                    End If
                                End If
                                Continue For
                            Case "WAREHOUSE_CODE"
                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                        ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans à la variable colonne
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si ce genre de type de format prevu
                                            Dim ValeurLue As String = OledatableSchemaSage.Rows(i).Item(ColonnName).ToString ' recuperation de la valeur
                                            If ValeurLue = "" Then ' si la valriable ne comtien aucune donne alors on ce rabat vers la valeur pas defaut
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then ' s'il ya une valeur pas defaut
                                                    Information &= Strings.LSet("[" & OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString & "]", ValeurLue.Length + 2)
                                                Else
                                                    Information &= Strings.LSet("[]", ValeurLue.Length + 2)
                                                End If
                                            Else
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            End If
                                        Else
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Dim DE_Intitule As String = RenvoieDepotPrincipal(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString)
                                                Information &= Strings.LSet(DE_Intitule, PositionLeft)
                                            End If
                                        End If
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        Else
                                            'un Mapping Sage existe mais ne possede pas de valeur que retourne la requete et ne possede pas de valeur pas defaut
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = ""
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        End If
                                    End If
                                Else 'si le champ Sage ne possede pas de mapping correspondant a la colonne EasyWMS
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut paramètrée
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                            Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else 'sinon le format n'est pas du type variable
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace à la fonction GetPosition qui prend en paramétre le format du champ
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        End If
                                    Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                        'Dim traitementAutres As String = 0
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                            Dim ValeurLue = ""
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(0.0, PositionLeft)
                                            End If
                                        End If
                                    End If
                                End If
                                Continue For
                            Case "DESCRIPTION"
                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                        ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans à la variable colonne
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si ce genre de type de format prevu
                                            Dim ValeurLue As String = OledatableSchemaSage.Rows(i).Item(ColonnName).ToString ' recuperation de la valeur
                                            If ValeurLue = "" Then ' si la valriable ne comtien aucune donne alors on ce rabat vers la valeur pas defaut
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then ' s'il ya une valeur pas defaut
                                                    Information &= Strings.LSet("[" & OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString & "]", ValeurLue.Length + 2)
                                                Else
                                                    Information &= Strings.LSet("[]", ValeurLue.Length + 2)
                                                End If
                                            Else
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            End If
                                        Else
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, PositionLeft)
                                            End If
                                        End If
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        Else
                                            'un Mapping Sage existe mais ne possede pas de valeur que retourne la requete et ne possede pas de valeur pas defaut
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = ""
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        End If
                                    End If
                                Else 'si le champ Sage ne possede pas de mapping correspondant a la colonne EasyWMS
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut paramètrée
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                            Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else 'sinon le format n'est pas du type variable
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace à la fonction GetPosition qui prend en paramétre le format du champ
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        End If
                                    Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                        'Dim traitementAutres As String = 0
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                            Dim ValeurLue = ""
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet("", PositionLeft)
                                            Else
                                                Information &= Strings.LSet("", PositionLeft)
                                            End If
                                        End If
                                    End If
                                End If
                                Continue For
                            Case "COMMENT_"
                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                        ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans à la variable colonne
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si ce genre de type de format prevu
                                            Dim ValeurLue As String = OledatableSchemaSage.Rows(i).Item(ColonnName).ToString ' recuperation de la valeur
                                            If ValeurLue = "" Then ' si la valriable ne comtien aucune donne alors on ce rabat vers la valeur pas defaut
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then ' s'il ya une valeur pas defaut
                                                    Information &= Strings.LSet("[" & OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString & "]", ValeurLue.Length + 2)
                                                Else
                                                    Information &= Strings.LSet("[]", ValeurLue.Length + 2)
                                                End If
                                            Else
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            End If
                                        Else
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, PositionLeft)
                                            End If
                                        End If
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        Else
                                            'un Mapping Sage existe mais ne possede pas de valeur que retourne la requete et ne possede pas de valeur pas defaut
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = ""
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        End If
                                    End If
                                Else 'si le champ Sage ne possede pas de mapping correspondant a la colonne EasyWMS
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut paramètrée
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                            Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else 'sinon le format n'est pas du type variable
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace à la fonction GetPosition qui prend en paramétre le format du champ
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        End If
                                    Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                        'Dim traitementAutres As String = 0
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                            Dim ValeurLue = ""
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet("", PositionLeft)
                                            Else
                                                Information &= Strings.LSet("", PositionLeft)
                                            End If
                                        End If
                                    End If
                                End If
                                Continue For
                            Case "TYPE_CODE"
                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                        ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans à la variable colonne
                                        Select Case Trim(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString)
                                            Case "1"
                                                TYPE_CODE = "1"
                                                Information &= Strings.LSet("CORDER", PositionLeft)
                                            Case "14"
                                                TYPE_CODE = "14"
                                                Information &= Strings.LSet("RETURN", PositionLeft)
                                            Case "23"
                                                TYPE_CODE = "23"
                                                Information &= Strings.LSet("TRANSFER", PositionLeft)
                                        End Select
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                        Else
                                            'un Mapping Sage existe mais ne possede pas de valeur que retourne la requete et ne possede pas de valeur pas defaut
                                            Information &= Strings.LSet("", PositionLeft)
                                        End If
                                    End If
                                Else 'si le champ Sage ne possede pas de mapping correspondant a la colonne EasyWMS
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut paramètrée
                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                    Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                        Information &= Strings.LSet("", PositionLeft)
                                    End If
                                End If
                                Continue For
                            Case "TYPE_DESCRIPTION"
                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                        ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans à la variable colonne
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, PositionLeft)
                                        End If
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        Else
                                            'un Mapping Sage existe mais ne possede pas de valeur que retourne la requete et ne possede pas de valeur pas defaut
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        End If
                                    End If
                                Else 'si le champ Sage ne possede pas de mapping correspondant a la colonne EasyWMS
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut paramètrée
                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace à la fonction GetPosition qui prend en paramétre le format du champ
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                        Else
                                            Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                        End If
                                    Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet("", PositionLeft)
                                        Else
                                            Information &= Strings.LSet("", PositionLeft)
                                        End If
                                    End If
                                End If
                                Continue For
                            Case "PRIORITY"
                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                        ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans à la variable colonne
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, PositionLeft)
                                        End If
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        Else
                                            'un Mapping Sage existe mais ne possede pas de valeur que retourne la requete et ne possede pas de valeur pas defaut
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        End If
                                    End If
                                Else 'si le champ Sage ne possede pas de mapping correspondant a la colonne EasyWMS
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut paramètrée
                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace à la fonction GetPosition qui prend en paramétre le format du champ
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                        Else
                                            Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                        End If
                                    Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet("", PositionLeft)
                                        Else
                                            Information &= Strings.LSet("", PositionLeft)
                                        End If
                                    End If
                                End If
                                Continue For
                            Case "ACCOUNT_CODE"
                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                        ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans à la variable colonne
                                        Select Case Trim(OledatableSchemaSage.Rows(i).Item("DO_TYPE").ToString)
                                            Case "1"
                                                Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, PositionLeft)
                                            Case Else
                                                Information &= Strings.LSet("", PositionLeft)
                                        End Select
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                        Else
                                            'un Mapping Sage existe mais ne possede pas de valeur que retourne la requete et ne possede pas de valeur pas defaut
                                            Information &= Strings.LSet("", PositionLeft)
                                        End If
                                    End If
                                Else ' le champ Sage ne possede pas de mapping correspondant a la colonne EasyWMS
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut paramètrée
                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace à la fonction GetPosition qui prend en paramétre le format du champ
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                        Else
                                            Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                        End If
                                    Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet("", PositionLeft)
                                        Else
                                            Information &= Strings.LSet("", PositionLeft)
                                        End If
                                    End If
                                End If
                                Continue For
                            Case "STAGING_LOC_CODE"
                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                        ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans à la variable colonne
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, PositionLeft)
                                        End If
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        Else
                                            'un Mapping Sage existe mais ne possede pas de valeur que retourne la requete et ne possede pas de valeur pas defaut
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        End If
                                    End If
                                Else 'si le champ Sage ne possede pas de mapping correspondant a la colonne EasyWMS
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut paramètrée
                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace à la fonction GetPosition qui prend en paramétre le format du champ
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                        Else
                                            Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                        End If
                                    Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet("", PositionLeft)
                                        Else
                                            Information &= Strings.LSet("", PositionLeft)
                                        End If
                                    End If
                                End If
                                Continue For
                            Case "DOOR_CODE"
                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                        ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans à la variable colonne
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, PositionLeft)
                                        End If
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        Else
                                            'un Mapping Sage existe mais ne possede pas de valeur que retourne la requete et ne possede pas de valeur pas defaut
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        End If
                                    End If
                                Else 'si le champ Sage ne possede pas de mapping correspondant a la colonne EasyWMS
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut paramètrée
                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace à la fonction GetPosition qui prend en paramétre le format du champ
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                        Else
                                            Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                        End If
                                    Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet("", PositionLeft)
                                        Else
                                            Information &= Strings.LSet("", PositionLeft)
                                        End If
                                    End If
                                End If
                                Continue For
                            Case "ROUTE_CODE"
                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                        ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans à la variable colonne
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si ce genre de type de format prevu
                                            Dim ValeurLue As String = OledatableSchemaSage.Rows(i).Item(ColonnName).ToString ' recuperation de la valeur
                                            If ValeurLue = "" Then ' si la valriable ne comtien aucune donne alors on ce rabat vers la valeur pas defaut
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then ' s'il ya une valeur pas defaut
                                                    Information &= Strings.LSet("[" & OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString & "]", ValeurLue.Length + 2)
                                                Else
                                                    Information &= Strings.LSet("[]", ValeurLue.Length + 2)
                                                End If
                                            Else
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            End If
                                        Else
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, PositionLeft)
                                            End If
                                        End If
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        Else
                                            'un Mapping Sage existe mais ne possede pas de valeur que retourne la requete et ne possede pas de valeur pas defaut
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = ""
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        End If
                                    End If
                                Else 'si le champ Sage ne possede pas de mapping correspondant a la colonne EasyWMS
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut paramètrée
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                            Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else 'sinon le format n'est pas du type variable
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace à la fonction GetPosition qui prend en paramétre le format du champ
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        End If
                                    Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                        'Dim traitementAutres As String = 0
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                            Dim ValeurLue = ""
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet("", PositionLeft)
                                            Else
                                                Information &= Strings.LSet("", PositionLeft)
                                            End If
                                        End If
                                    End If
                                End If
                                Continue For
                            Case "STOP_NUMBER"
                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                        ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans à la variable colonne
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si ce genre de type de format prevu
                                            Dim ValeurLue As String = OledatableSchemaSage.Rows(i).Item(ColonnName).ToString ' recuperation de la valeur
                                            If ValeurLue = "" Then ' si la valriable ne comtien aucune donne alors on ce rabat vers la valeur pas defaut
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then ' s'il ya une valeur pas defaut
                                                    Information &= Strings.LSet("[" & OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString & "]", ValeurLue.Length + 2)
                                                Else
                                                    Information &= Strings.LSet("[]", ValeurLue.Length + 2)
                                                End If
                                            Else
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            End If
                                        Else
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, PositionLeft)
                                            End If
                                        End If
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        Else
                                            'un Mapping Sage existe mais ne possede pas de valeur que retourne la requete et ne possede pas de valeur pas defaut
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = ""
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        End If
                                    End If
                                Else 'si le champ Sage ne possede pas de mapping correspondant a la colonne EasyWMS
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut paramètrée
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                            Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else 'sinon le format n'est pas du type variable
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace à la fonction GetPosition qui prend en paramétre le format du champ
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        End If
                                    Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                        'Dim traitementAutres As String = 0
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                            Dim ValeurLue = ""
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet("", PositionLeft)
                                            Else
                                                Information &= Strings.LSet("", PositionLeft)
                                            End If
                                        End If
                                    End If
                                End If
                                Continue For
                            Case "CARRIER_CODE"
                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                        ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans à la variable colonne
                                        Dim OleAdaptaterschemaCARRIER_CODE As OleDbDataAdapter = Nothing
                                        If ColonnName = "DO_EXPEDIT" Then
                                            OleAdaptaterschemaCARRIER_CODE = New OleDbDataAdapter("SELECT E_Intitule FROM P_EXPEDITION WHERE P_EXPEDITION.CbMarq=" & OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, OleExcelConnect)
                                        ElseIf ColonnName.ToUpper = "DO_CONDITION" Then
                                            OleAdaptaterschemaCARRIER_CODE = New OleDbDataAdapter("SELECT C_Intitule FROM P_CONDLIVR WHERE P_CONDLIVR.CbMarq=" & OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, OleExcelConnect)
                                        End If
                                        Dim OleSchemaDatasetCARRIER_CODE = New DataSet
                                        OleAdaptaterschemaCARRIER_CODE.Fill(OleSchemaDatasetCARRIER_CODE)
                                        Dim OledatableSchemaCARRIER_CODE As DataTable = OleSchemaDatasetCARRIER_CODE.Tables(0)

                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si ce genre de type de format prevu
                                            Dim ValeurLue As String = OledatableSchemaSage.Rows(i).Item(ColonnName).ToString ' recuperation de la valeur
                                            If ValeurLue = "" Then ' si la valriable ne comtien aucune donne alors on ce rabat vers la valeur pas defaut
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then ' s'il ya une valeur pas defaut
                                                    Information &= Strings.LSet("[" & OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString & "]", ValeurLue.Length + 2)
                                                Else
                                                    Information &= Strings.LSet("[]", ValeurLue.Length + 2)
                                                End If
                                            Else
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            End If
                                        Else
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                If OledatableSchemaCARRIER_CODE.Rows.Count <> 0 Then
                                                    Information &= Strings.LSet(OledatableSchemaCARRIER_CODE.Rows(0).Item(0).ToString, PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, PositionLeft)
                                                End If
                                            End If
                                        End If
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        Else
                                            'un Mapping Sage existe mais ne possede pas de valeur que retourne la requete et ne possede pas de valeur pas defaut
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = ""
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        End If
                                    End If
                                Else 'si le champ Sage ne possede pas de mapping correspondant a la colonne EasyWMS
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut paramètrée
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                            Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else 'sinon le format n'est pas du type variable
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace à la fonction GetPosition qui prend en paramétre le format du champ
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        End If
                                    Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                        'Dim traitementAutres As String = 0
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                            Dim ValeurLue = ""
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet("", PositionLeft)
                                            Else
                                                Information &= Strings.LSet("", PositionLeft)
                                            End If
                                        End If
                                    End If
                                End If
                                Continue For
                            Case "CONTAINERTYPE_CODE"
                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                        ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans à la variable colonne
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si ce genre de type de format prevu
                                            Dim ValeurLue As String = OledatableSchemaSage.Rows(i).Item(ColonnName).ToString ' recuperation de la valeur
                                            If ValeurLue = "" Then ' si la valriable ne comtien aucune donne alors on ce rabat vers la valeur pas defaut
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then ' s'il ya une valeur pas defaut
                                                    Information &= Strings.LSet("[" & OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString & "]", ValeurLue.Length + 2)
                                                Else
                                                    Information &= Strings.LSet("[]", ValeurLue.Length + 2)
                                                End If
                                            Else
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            End If
                                        Else
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, PositionLeft)
                                            End If
                                        End If
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        Else
                                            'un Mapping Sage existe mais ne possede pas de valeur que retourne la requete et ne possede pas de valeur pas defaut
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = ""
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        End If
                                    End If
                                Else 'si le champ Sage ne possede pas de mapping correspondant a la colonne EasyWMS
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut paramètrée
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                            Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else 'sinon le format n'est pas du type variable
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace à la fonction GetPosition qui prend en paramétre le format du champ
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        End If
                                    Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                        'Dim traitementAutres As String = 0
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                            Dim ValeurLue = ""
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet("", PositionLeft)
                                            Else
                                                Information &= Strings.LSet("", PositionLeft)
                                            End If
                                        End If
                                    End If
                                End If
                                Continue For
                            Case "DOCUMENT"
                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                        ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans à la variable colonne
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si ce genre de type de format prevu
                                            Dim ValeurLue As String = OledatableSchemaSage.Rows(i).Item(ColonnName).ToString ' recuperation de la valeur
                                            If ValeurLue = "" Then ' si la valriable ne comtien aucune donne alors on ce rabat vers la valeur pas defaut
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then ' s'il ya une valeur pas defaut
                                                    Information &= Strings.LSet("[" & OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString & "]", ValeurLue.Length + 2)
                                                Else
                                                    Information &= Strings.LSet("[]", ValeurLue.Length + 2)
                                                End If
                                            Else
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            End If
                                        Else
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                If Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, PositionLeft) = "" Then
                                                    Information &= Strings.LSet(Societe, PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, PositionLeft)
                                                End If
                                            End If
                                        End If
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        Else
                                            'un Mapping Sage existe mais ne possede pas de valeur que retourne la requete et ne possede pas de valeur pas defaut
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = ""
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(Societe, PositionLeft)
                                                End If
                                            End If
                                        End If
                                    End If
                                Else 'si le champ Sage ne possede pas de mapping correspondant a la colonne EasyWMS
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut paramètrée
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                            Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else 'sinon le format n'est pas du type variable
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace à la fonction GetPosition qui prend en paramétre le format du champ
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        End If
                                    Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                        'Dim traitementAutres As String = 0
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                            Dim ValeurLue = ""
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(Societe, PositionLeft)
                                            Else
                                                Information &= Strings.LSet(Societe, PositionLeft)
                                            End If
                                        End If
                                    End If
                                End If
                                Continue For
                            Case "TRANSPORT_TYPE"
                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                        ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans à la variable colonne
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si ce genre de type de format prevu
                                            Dim ValeurLue As String = OledatableSchemaSage.Rows(i).Item(ColonnName).ToString ' recuperation de la valeur
                                            If ValeurLue = "" Then ' si la valriable ne comtien aucune donne alors on ce rabat vers la valeur pas defaut
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then ' s'il ya une valeur pas defaut
                                                    Information &= Strings.LSet("[" & OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString & "]", ValeurLue.Length + 2)
                                                Else
                                                    Information &= Strings.LSet("[]", ValeurLue.Length + 2)
                                                End If
                                            Else
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            End If
                                        Else
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, PositionLeft)
                                            End If
                                        End If
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        Else
                                            'un Mapping Sage existe mais ne possede pas de valeur que retourne la requete et ne possede pas de valeur pas defaut
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = ""
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        End If
                                    End If
                                Else 'si le champ Sage ne possede pas de mapping correspondant a la colonne EasyWMS
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut paramètrée
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                            Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else 'sinon le format n'est pas du type variable
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace à la fonction GetPosition qui prend en paramétre le format du champ
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        End If
                                    Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                        'Dim traitementAutres As String = 0
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                            Dim ValeurLue = ""
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet("", PositionLeft)
                                            Else
                                                Information &= Strings.LSet("", PositionLeft)
                                            End If
                                        End If
                                    End If
                                End If
                                Continue For
                            Case "SOURCE"
                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                        ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans à la variable colonne
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si ce genre de type de format prevu
                                            Dim ValeurLue As String = OledatableSchemaSage.Rows(i).Item(ColonnName).ToString ' recuperation de la valeur
                                            If ValeurLue = "" Then ' si la valriable ne comtien aucune donne alors on ce rabat vers la valeur pas defaut
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then ' s'il ya une valeur pas defaut
                                                    Information &= Strings.LSet("[" & OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString & "]", ValeurLue.Length + 2)
                                                Else
                                                    Information &= Strings.LSet("[]", ValeurLue.Length + 2)
                                                End If
                                            Else
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            End If
                                        Else
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, PositionLeft)
                                            End If
                                        End If
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        Else
                                            'un Mapping Sage existe mais ne possede pas de valeur que retourne la requete et ne possede pas de valeur pas defaut
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = ""
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        End If
                                    End If
                                Else 'si le champ Sage ne possede pas de mapping correspondant a la colonne EasyWMS
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut paramètrée
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                            Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else 'sinon le format n'est pas du type variable
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace à la fonction GetPosition qui prend en paramétre le format du champ
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        End If
                                    Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                        'Dim traitementAutres As String = 0
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                            Dim ValeurLue = ""
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet("", PositionLeft)
                                            Else
                                                Information &= Strings.LSet("", PositionLeft)
                                            End If
                                        End If
                                    End If
                                End If
                                Continue For

                            Case "DELIVERY_INST"
                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                        ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans à la variable colonne
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si ce genre de type de format prevu
                                            Dim ValeurLue As String = OledatableSchemaSage.Rows(i).Item(ColonnName).ToString ' recuperation de la valeur
                                            If ValeurLue = "" Then ' si la valriable ne comtien aucune donne alors on ce rabat vers la valeur pas defaut
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then ' s'il ya une valeur pas defaut
                                                    Information &= Strings.LSet("[" & OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString & "]", ValeurLue.Length + 2)
                                                Else
                                                    Information &= Strings.LSet("[]", ValeurLue.Length + 2)
                                                End If
                                            Else
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            End If
                                        Else
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, PositionLeft)
                                            End If
                                        End If
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        Else
                                            'un Mapping Sage existe mais ne possede pas de valeur que retourne la requete et ne possede pas de valeur pas defaut
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = ""
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        End If
                                    End If
                                Else 'si le champ Sage ne possede pas de mapping correspondant a la colonne EasyWMS
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut paramètrée
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                            Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else 'sinon le format n'est pas du type variable
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace à la fonction GetPosition qui prend en paramétre le format du champ
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        End If
                                    Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                        'Dim traitementAutres As String = 0
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                            Dim ValeurLue = ""
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet("", PositionLeft)
                                            Else
                                                Information &= Strings.LSet("", PositionLeft)
                                            End If
                                        End If
                                    End If
                                End If
                                Continue For
                            Case "ALLOCATE_DATE"
                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                        ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans à la variable colonne
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si ce genre de type de format prevu
                                            Dim ValeurLue As String = OledatableSchemaSage.Rows(i).Item(ColonnName).ToString ' recuperation de la valeur
                                            If ValeurLue = "" Then ' si la valriable ne comtien aucune donne alors on ce rabat vers la valeur pas defaut
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then ' s'il ya une valeur pas defaut
                                                    Information &= Strings.LSet("[" & OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString & "]", ValeurLue.Length + 2)
                                                Else
                                                    Information &= Strings.LSet("[]", ValeurLue.Length + 2)
                                                End If
                                            Else
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            End If
                                        Else
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, PositionLeft)
                                            End If
                                        End If
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        Else
                                            'un Mapping Sage existe mais ne possede pas de valeur que retourne la requete et ne possede pas de valeur pas defaut
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = ""
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        End If
                                    End If
                                Else 'si le champ Sage ne possede pas de mapping correspondant a la colonne EasyWMS
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut paramètrée
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                            Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else 'sinon le format n'est pas du type variable
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace à la fonction GetPosition qui prend en paramétre le format du champ
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        End If
                                    Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                        'Dim traitementAutres As String = 0
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                            Dim ValeurLue = ""
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet("", PositionLeft)
                                            Else
                                                Information &= Strings.LSet("", PositionLeft)
                                            End If
                                        End If
                                    End If
                                End If
                                Continue For
                            Case "PROCESS_DATE"
                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                        ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans à la variable colonne
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si ce genre de type de format prevu
                                            Dim ValeurLue As String = OledatableSchemaSage.Rows(i).Item(ColonnName).ToString ' recuperation de la valeur
                                            If ValeurLue = "" Then ' si la valriable ne comtien aucune donne alors on ce rabat vers la valeur pas defaut
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then ' s'il ya une valeur pas defaut
                                                    Information &= Strings.LSet("[" & OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString & "]", ValeurLue.Length + 2)
                                                Else
                                                    Information &= Strings.LSet("[]", ValeurLue.Length + 2)
                                                End If
                                            Else
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            End If
                                        Else
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, PositionLeft)
                                            End If
                                        End If
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        Else
                                            'un Mapping Sage existe mais ne possede pas de valeur que retourne la requete et ne possede pas de valeur pas defaut
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = ""
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        End If
                                    End If
                                Else 'si le champ Sage ne possede pas de mapping correspondant a la colonne EasyWMS
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut paramètrée
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                            Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else 'sinon le format n'est pas du type variable
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace à la fonction GetPosition qui prend en paramétre le format du champ
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        End If
                                    Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                        'Dim traitementAutres As String = 0
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                            Dim ValeurLue = ""
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet("", PositionLeft)
                                            Else
                                                Information &= Strings.LSet("", PositionLeft)
                                            End If
                                        End If
                                    End If
                                End If
                                Continue For
                            Case "SHIPPING_DATE"
                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                        ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans à la variable colonne
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si ce genre de type de format prevu
                                            Dim ValeurLue As String = OledatableSchemaSage.Rows(i).Item(ColonnName).ToString ' recuperation de la valeur
                                            If ValeurLue = "" Then ' si la valriable ne comtien aucune donne alors on ce rabat vers la valeur pas defaut
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then ' s'il ya une valeur pas defaut
                                                    Information &= Strings.LSet("[" & OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString & "]", ValeurLue.Length + 2)
                                                Else
                                                    Information &= Strings.LSet("[]", ValeurLue.Length + 2)
                                                End If
                                            Else
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            End If
                                        Else
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, PositionLeft)
                                            End If
                                        End If
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        Else
                                            'un Mapping Sage existe mais ne possede pas de valeur que retourne la requete et ne possede pas de valeur pas defaut
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = ""
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        End If
                                    End If
                                Else 'si le champ Sage ne possede pas de mapping correspondant a la colonne EasyWMS
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut paramètrée
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                            Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else 'sinon le format n'est pas du type variable
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace à la fonction GetPosition qui prend en paramétre le format du champ
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        End If
                                    Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                        'Dim traitementAutres As String = 0
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                            Dim ValeurLue = ""
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet("", PositionLeft)
                                            Else
                                                Information &= Strings.LSet("", PositionLeft)
                                            End If
                                        End If
                                    End If
                                End If
                                Continue For
                            Case "PLANNED_LOAD_DATE"
                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                        ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans à la variable colonne
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si ce genre de type de format prevu
                                            Dim ValeurLue As String = OledatableSchemaSage.Rows(i).Item(ColonnName).ToString ' recuperation de la valeur
                                            If ValeurLue = "" Then ' si la valriable ne comtien aucune donne alors on ce rabat vers la valeur pas defaut
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then ' s'il ya une valeur pas defaut
                                                    Information &= Strings.LSet("[" & OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString & "]", ValeurLue.Length + 2)
                                                Else
                                                    Information &= Strings.LSet("[]", ValeurLue.Length + 2)
                                                End If
                                            Else
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            End If
                                        Else
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, PositionLeft)
                                            End If
                                        End If
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        Else
                                            'un Mapping Sage existe mais ne possede pas de valeur que retourne la requete et ne possede pas de valeur pas defaut
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = ""
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        End If
                                    End If
                                Else 'si le champ Sage ne possede pas de mapping correspondant a la colonne EasyWMS
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut paramètrée
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                            Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else 'sinon le format n'est pas du type variable
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace à la fonction GetPosition qui prend en paramétre le format du champ
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        End If
                                    Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                        'Dim traitementAutres As String = 0
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                            Dim ValeurLue = ""
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet("", PositionLeft)
                                            Else
                                                Information &= Strings.LSet("", PositionLeft)
                                            End If
                                        End If
                                    End If
                                End If
                                Continue For
                            Case "CONTAINERS"
                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                        ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans à la variable colonne
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si ce genre de type de format prevu
                                            Dim ValeurLue As String = OledatableSchemaSage.Rows(i).Item(ColonnName).ToString ' recuperation de la valeur
                                            If ValeurLue = "" Then ' si la valriable ne comtien aucune donne alors on ce rabat vers la valeur pas defaut
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then ' s'il ya une valeur pas defaut
                                                    Information &= Strings.LSet("[" & OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString & "]", ValeurLue.Length + 2)
                                                Else
                                                    Information &= Strings.LSet("[]", ValeurLue.Length + 2)
                                                End If
                                            Else
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            End If
                                        Else
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, PositionLeft)
                                            End If
                                        End If
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        Else
                                            'un Mapping Sage existe mais ne possede pas de valeur que retourne la requete et ne possede pas de valeur pas defaut
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = ""
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        End If
                                    End If
                                Else 'si le champ Sage ne possede pas de mapping correspondant a la colonne EasyWMS
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut paramètrée
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                            Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else 'sinon le format n'est pas du type variable
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace à la fonction GetPosition qui prend en paramétre le format du champ
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        End If
                                    Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                        'Dim traitementAutres As String = 0
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                            Dim ValeurLue = ""
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet("", PositionLeft)
                                            Else
                                                Information &= Strings.LSet("", PositionLeft)
                                            End If
                                        End If
                                    End If
                                End If
                                Continue For
                            Case "VERIFY_STOCK"
                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                        ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans à la variable colonne
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si ce genre de type de format prevu
                                            Dim ValeurLue As String = OledatableSchemaSage.Rows(i).Item(ColonnName).ToString ' recuperation de la valeur
                                            If ValeurLue = "" Then ' si la valriable ne comtien aucune donne alors on ce rabat vers la valeur pas defaut
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then ' s'il ya une valeur pas defaut
                                                    Information &= Strings.LSet("[" & OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString & "]", ValeurLue.Length + 2)
                                                Else
                                                    Information &= Strings.LSet("[]", ValeurLue.Length + 2)
                                                End If
                                            Else
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            End If
                                        Else
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, PositionLeft)
                                            End If
                                        End If
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        Else
                                            'un Mapping Sage existe mais ne possede pas de valeur que retourne la requete et ne possede pas de valeur pas defaut
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = ""
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        End If
                                    End If
                                Else 'si le champ Sage ne possede pas de mapping correspondant a la colonne EasyWMS
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut paramètrée
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                            Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else 'sinon le format n'est pas du type variable
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace à la fonction GetPosition qui prend en paramétre le format du champ
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        End If
                                    Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                        'Dim traitementAutres As String = 0
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                            Dim ValeurLue = ""
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet("", PositionLeft)
                                            Else
                                                Information &= Strings.LSet("", PositionLeft)
                                            End If
                                        End If
                                    End If
                                End If
                                Continue For

                            Case "FOLLOW_SEQUENCE"
                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                        ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans à la variable colonne
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si ce genre de type de format prevu
                                            Dim ValeurLue As String = OledatableSchemaSage.Rows(i).Item(ColonnName).ToString ' recuperation de la valeur
                                            If ValeurLue = "" Then ' si la valriable ne comtien aucune donne alors on ce rabat vers la valeur pas defaut
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then ' s'il ya une valeur pas defaut
                                                    Information &= Strings.LSet("[" & OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString & "]", ValeurLue.Length + 2)
                                                Else
                                                    Information &= Strings.LSet("[]", ValeurLue.Length + 2)
                                                End If
                                            Else
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            End If
                                        Else
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, PositionLeft)
                                            End If
                                        End If
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        Else
                                            'un Mapping Sage existe mais ne possede pas de valeur que retourne la requete et ne possede pas de valeur pas defaut
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = ""
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        End If
                                    End If
                                Else 'si le champ Sage ne possede pas de mapping correspondant a la colonne EasyWMS
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut paramètrée
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                            Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else 'sinon le format n'est pas du type variable
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace à la fonction GetPosition qui prend en paramétre le format du champ
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        End If
                                    Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                        'Dim traitementAutres As String = 0
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                            Dim ValeurLue = ""
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet("", PositionLeft)
                                            Else
                                                Information &= Strings.LSet("", PositionLeft)
                                            End If
                                        End If
                                    End If
                                End If
                                Continue For
                            Case "RELEASE_DATE"
                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                        ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans à la variable colonne
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si ce genre de type de format prevu
                                            Dim ValeurLue As String = OledatableSchemaSage.Rows(i).Item(ColonnName).ToString ' recuperation de la valeur
                                            If ValeurLue = "" Then ' si la valriable ne comtien aucune donne alors on ce rabat vers la valeur pas defaut
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then ' s'il ya une valeur pas defaut
                                                    Information &= Strings.LSet("[" & OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString & "]", ValeurLue.Length + 2)
                                                Else
                                                    Information &= Strings.LSet("[]", ValeurLue.Length + 2)
                                                End If
                                            Else
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            End If
                                        Else
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, PositionLeft)
                                            End If
                                        End If
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        Else
                                            'un Mapping Sage existe mais ne possede pas de valeur que retourne la requete et ne possede pas de valeur pas defaut
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = ""
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        End If
                                    End If
                                Else 'si le champ Sage ne possede pas de mapping correspondant a la colonne EasyWMS
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut paramètrée
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                            Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else 'sinon le format n'est pas du type variable
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace à la fonction GetPosition qui prend en paramétre le format du champ
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        End If
                                    Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                        'Dim traitementAutres As String = 0
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                            Dim ValeurLue = ""
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet("", PositionLeft)
                                            Else
                                                Information &= Strings.LSet("", PositionLeft)
                                            End If
                                        End If
                                    End If
                                End If
                                Continue For
                            Case "SUPPLIER_CODE"
                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                        ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans à la variable colonne
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si ce genre de type de format prevu
                                            Dim ValeurLue As String = OledatableSchemaSage.Rows(i).Item(ColonnName).ToString ' recuperation de la valeur
                                            If ValeurLue = "" Then ' si la valriable ne comtien aucune donne alors on ce rabat vers la valeur pas defaut
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then ' s'il ya une valeur pas defaut
                                                    Information &= Strings.LSet("[" & OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString & "]", ValeurLue.Length + 2)
                                                Else
                                                    Information &= Strings.LSet("[]", ValeurLue.Length + 2)
                                                End If
                                            Else
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            End If
                                        Else
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                If TYPE_CODE = "14" Then
                                                    Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, PositionLeft)
                                                Else
                                                    Information &= Strings.LSet("", PositionLeft)
                                                End If
                                            End If
                                        End If
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        Else
                                            'un Mapping Sage existe mais ne possede pas de valeur que retourne la requete et ne possede pas de valeur pas defaut
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = ""
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        End If
                                    End If
                                Else 'si le champ Sage ne possede pas de mapping correspondant a la colonne EasyWMS
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut paramètrée
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                            Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else 'sinon le format n'est pas du type variable
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace à la fonction GetPosition qui prend en paramétre le format du champ
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        End If
                                    Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                        'Dim traitementAutres As String = 0
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                            Dim ValeurLue = ""
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet("", PositionLeft)
                                            Else
                                                Information &= Strings.LSet("", PositionLeft)
                                            End If
                                        End If
                                    End If
                                End If
                                Continue For
                            Case "AUTO_RESERVE_DATE"

                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                        ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans à la variable colonne
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si ce genre de type de format prevu
                                            Dim ValeurLue As String = OledatableSchemaSage.Rows(i).Item(ColonnName).ToString ' recuperation de la valeur
                                            If ValeurLue = "" Then ' si la valriable ne comtien aucune donne alors on ce rabat vers la valeur pas defaut
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then ' s'il ya une valeur pas defaut
                                                    Information &= Strings.LSet("[" & OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString & "]", ValeurLue.Length + 2)
                                                Else
                                                    Information &= Strings.LSet("[]", ValeurLue.Length + 2)
                                                End If
                                            Else
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            End If
                                        Else
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, PositionLeft)
                                            End If
                                        End If
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        Else
                                            'un Mapping Sage existe mais ne possede pas de valeur que retourne la requete et ne possede pas de valeur pas defaut
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = ""
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        End If
                                    End If
                                Else 'si le champ Sage ne possede pas de mapping correspondant a la colonne EasyWMS
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut paramètrée
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                            Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else 'sinon le format n'est pas du type variable
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace à la fonction GetPosition qui prend en paramétre le format du champ
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        End If
                                    Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                        'Dim traitementAutres As String = 0
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                            Dim ValeurLue = ""
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet("", PositionLeft)
                                            Else
                                                Information &= Strings.LSet("", PositionLeft)
                                            End If
                                        End If
                                    End If
                                End If
                                Continue For

                            Case "VALID_DATE"

                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                        ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans à la variable colonne
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si ce genre de type de format prevu
                                            Dim ValeurLue As String = OledatableSchemaSage.Rows(i).Item(ColonnName).ToString ' recuperation de la valeur
                                            If ValeurLue = "" Then ' si la valriable ne comtien aucune donne alors on ce rabat vers la valeur pas defaut
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then ' s'il ya une valeur pas defaut
                                                    Information &= Strings.LSet("[" & OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString & "]", ValeurLue.Length + 2)
                                                Else
                                                    Information &= Strings.LSet("[]", ValeurLue.Length + 2)
                                                End If
                                            Else
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            End If
                                        Else
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, PositionLeft)
                                            End If
                                        End If
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        Else
                                            'un Mapping Sage existe mais ne possede pas de valeur que retourne la requete et ne possede pas de valeur pas defaut
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = ""
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        End If
                                    End If
                                Else 'si le champ Sage ne possede pas de mapping correspondant a la colonne EasyWMS
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut paramètrée
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                            Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else 'sinon le format n'est pas du type variable
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace à la fonction GetPosition qui prend en paramétre le format du champ
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        End If
                                    Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                        'Dim traitementAutres As String = 0
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                            Dim ValeurLue = ""
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet("", PositionLeft)
                                            Else
                                                Information &= Strings.LSet("", PositionLeft)
                                            End If
                                        End If
                                    End If
                                End If
                                Continue For
                            Case "WAREHOUSE_CODE_DEST"
                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                        ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans à la variable colonne
                                        If OledatableSchemaSage.Rows(i).Item(ColonnName).ToString = "23" Then
                                            Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, PositionLeft)
                                        Else
                                            Information &= Strings.LSet("", PositionLeft)
                                        End If
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        Else
                                            'un Mapping Sage existe mais ne possede pas de valeur que retourne la requete et ne possede pas de valeur pas defaut
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = ""
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        End If
                                    End If
                                Else 'si le champ Sage ne possede pas de mapping correspondant a la colonne EasyWMS
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut paramètrée
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                            Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else 'sinon le format n'est pas du type variable
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace à la fonction GetPosition qui prend en paramétre le format du champ
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        End If
                                    Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                        'Dim traitementAutres As String = 0
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                            Dim ValeurLue = ""
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet("", PositionLeft)
                                            Else
                                                Information &= Strings.LSet("", PositionLeft)
                                            End If
                                        End If
                                    End If
                                End If
                                Continue For
                            Case "LINE_COUNT"
                                Dim NbreLigne As Integer = CounteLigne(OledatableSchemaSage.Rows(i).Item("DO_PIECE").ToString)
                                Information &= Strings.LSet(RenvoiValeurLue(NbreLigne, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                Continue For
                            Case Else
                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                        ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans à la variable colonne
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si ce genre de type de format prevu
                                            Dim ValeurLue As String = OledatableSchemaSage.Rows(i).Item(ColonnName).ToString ' recuperation de la valeur
                                            If ValeurLue = "" Then ' si la valriable ne comtien aucune donne alors on ce rabat vers la valeur pas defaut
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then ' s'il ya une valeur pas defaut
                                                    Information &= Strings.LSet("[" & OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString & "]", ValeurLue.Length + 2)
                                                Else
                                                    Information &= Strings.LSet("[]", ValeurLue.Length + 2)
                                                End If
                                            Else
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            End If
                                        Else
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, PositionLeft)
                                            End If
                                        End If
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        Else
                                            'un Mapping Sage existe mais ne possede pas de valeur que retourne la requete et ne possede pas de valeur pas defaut
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue = ""
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        End If
                                    End If
                                Else 'si le champ Sage ne possede pas de mapping correspondant a la colonne EasyWMS
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut paramètrée
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                            Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else 'sinon le format n'est pas du type variable
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace à la fonction GetPosition qui prend en paramétre le format du champ
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        End If
                                    Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                        'Dim traitementAutres As String = 0
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                            Dim ValeurLue = ""
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet("", PositionLeft)
                                            Else
                                                Information &= Strings.LSet("", PositionLeft)
                                            End If
                                        End If
                                    End If
                                End If
                                Continue For
                        End Select
                        '-------------------------------------**********FIN DONNEES ENTETE********************------------------------------------------
                    Next
                    Error_journal.WriteLine(Information)
                    Information = ""
                    '--------------------------debut traitement de la ligne---------------------------------
                    CreationLigne(OledatableSchemaSage.Rows(i).Item("DO_PIECE").ToString)
                    '-------------------------------***********Fin Traitement de la ligne**********------------
                    statut = True
                    Try
                        If ChFlag.Checked Then
                            Dim OleCommande As New OleDbCommand("UPDATE F_DOCENTETE SET " & Flagtampon & "='" & Format(DateTime.Now, "yyyyMMddhhmm") & "'  WHERE DO_PIECE='" & OledatableSchemaSage.Rows(i).Item("DO_PIECE").ToString.Trim & "'", OleExcelConnect)
                            OleCommande.ExecuteNonQuery()
                        End If
                    Catch ex As Exception
                    End Try
                Next
                If statut = True Then
                    Error_journal.Close()
                End If
            Else
                'cette requette ne possede pas ligne 
                EstVide = False
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Extraction données d'entete ")
            Return False
        End Try
    End Function
    Dim nbrLigneVue As Integer = 0
    Public Sub CreationLigne(ByVal valeurLue As Object)
        Dim PositionLeftLigne As Integer = 0
        Dim Information As String = ""
        Dim ColonnNameLigne As String = ""
        If ListeChampSageLigne IsNot Nothing Then
            If RbtVente.Checked Then
                OleAdaptaterschemaSageDetails = New OleDbDataAdapter("SELECT " & ListeChampSageLigne & " FROM F_DOCLIGNE WHERE F_DOCLIGNE.DO_DOMAINE=0 AND F_DOCLIGNE.DO_TYPE=1 AND F_DOCLIGNE.DO_PIECE='" & valeurLue & "'  AND AR_REF IS NOT NULL ORDER BY F_DOCLIGNE.DL_LIGNE", OleExcelConnect)
            ElseIf RbtAchat.Checked Then
                OleAdaptaterschemaSageDetails = New OleDbDataAdapter("SELECT " & ListeChampSageLigne & " FROM F_DOCLIGNE WHERE F_DOCLIGNE.DO_DOMAINE=1 AND F_DOCLIGNE.DO_TYPE=14 AND F_DOCLIGNE.DO_PIECE='" & valeurLue & "' AND AR_REF IS NOT NULL ORDER BY F_DOCLIGNE.DL_LIGNE", OleExcelConnect)
            ElseIf RbtStock.Checked Then
                OleAdaptaterschemaSageDetails = New OleDbDataAdapter("SELECT " & ListeChampSageLigne & " FROM F_DOCLIGNE WHERE F_DOCLIGNE.DO_DOMAINE=2 AND F_DOCLIGNE.DO_TYPE=23 AND F_DOCLIGNE.DO_PIECE='" & valeurLue & "' AND AR_REF IS NOT NULL ORDER BY F_DOCLIGNE.DL_LIGNE", OleExcelConnect)
            End If
        Else
            If RbtVente.Checked Then
                OleAdaptaterschemaSageDetails = New OleDbDataAdapter("SELECT * FROM F_DOCLIGNE WHERE F_DOCLIGNE.DO_DOMAINE=0 AND F_DOCLIGNE.DO_TYPE=1 AND F_DOCLIGNE.DO_PIECE='" & valeurLue & "'  AND AR_REF IS NOT NULL ORDER BY F_DOCLIGNE.DL_LIGNE", OleExcelConnect)
            ElseIf RbtAchat.Checked Then
                OleAdaptaterschemaSageDetails = New OleDbDataAdapter("SELECT * FROM F_DOCLIGNE WHERE F_DOCLIGNE.DO_DOMAINE=1 AND F_DOCLIGNE.DO_TYPE=14 AND F_DOCLIGNE.DO_PIECE='" & valeurLue & "' AND AR_REF IS NOT NULL ORDER BY F_DOCLIGNE.DL_LIGNE", OleExcelConnect)
            ElseIf RbtStock.Checked Then
                OleAdaptaterschemaSageDetails = New OleDbDataAdapter("SELECT * FROM F_DOCLIGNE WHERE F_DOCLIGNE.DO_DOMAINE=2 AND F_DOCLIGNE.DO_TYPE=23 AND F_DOCLIGNE.DO_PIECE='" & valeurLue & "' AND AR_REF IS NOT NULL ORDER BY F_DOCLIGNE.DL_LIGNE", OleExcelConnect)
            End If
        End If
        OleSchemaDatasetSageDetails = New DataSet
        OleAdaptaterschemaSageDetails.Fill(OleSchemaDatasetSageDetails)
        OledatableSchemaSageDetails = OleSchemaDatasetSageDetails.Tables(0)

        OleAdaptaterschemaLigne = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='SOR' AND Ligne=True ORDER BY ORDRE", OleConnenection)
        OleSchemaDatasetLigne = New DataSet
        OleAdaptaterschemaLigne.Fill(OleSchemaDatasetLigne)
        OledatableSchemaLigne = OleSchemaDatasetLigne.Tables(0)
        Try
            If OledatableSchemaSageDetails.Rows.Count <> 0 Then
                For i As Integer = 0 To OledatableSchemaSageDetails.Rows.Count - 1
                    '
                    nbrLigneVue = i + 1
                    Element2 = valeurLue
                    Element3 = OledatableSchemaSageDetails.Rows(i).Item("DL_LIGNE")
                    If ListBox.InvokeRequired Then
                        Dim MonDelegate As New Evenement(AddressOf Traitement2)
                        ListBox.Invoke(MonDelegate)
                    Else
                        Traitement2()
                    End If
                    For Ligne As Integer = 0 To OledatableSchemaLigne.Rows.Count - 1
                        If Ligne = 0 Then
                            Information &= Strings.LSet("/", 1)
                            Continue For
                        End If
                        Dim CARACTERE As String = OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0).ToString
                        If CARACTERE <> "V" Then
                            PositionLeftLigne = GetPosition(OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString)
                        End If
                        Dim oo = OledatableSchemaLigne.Rows(Ligne).Item("Cols").ToString
                        Select Case Trim(OledatableSchemaLigne.Rows(Ligne).Item("Cols").ToString)
                            Case "LINE_NUMBER"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "PRODUCT_CODE"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "LOT_CODE"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For

                            Case "QUANTITY"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "EAN_CODE"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "SERIAL_NO"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "EXPIRATION_DATE"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "SHELF_DAYS"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "UOM_CODE"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then

                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    Dim OleAdaptaterschemaEC_EMUNERE = New OleDbDataAdapter("SELECT P_CONDITIONNEMENT FROM F_ARTICLE INNER JOIN P_CONDITIONNEMENT ON F_ARTICLE.AR_CONDITION = P_CONDITIONNEMENT.cbindice AND F_ARTICLE.AR_Ref='" & OledatableSchemaSageDetails.Rows(i).Item("AR_REF").ToString & "'", OleExcelConnect)
                                    Dim OleSchemaDatasetEC_EMUNERE = New DataSet
                                    OleAdaptaterschemaEC_EMUNERE.Fill(OleSchemaDatasetEC_EMUNERE)
                                    Dim OledatableSchemaEC_EMUNERE As DataTable = OleSchemaDatasetEC_EMUNERE.Tables(0)

                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            If OledatableSchemaEC_EMUNERE.Rows.Count <> 0 Then
                                                Information &= Strings.LSet(OledatableSchemaEC_EMUNERE.Rows(0).Item("P_CONDITIONNEMENT").ToString, PositionLeftLigne)
                                            Else
                                                If Trim(Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)) <> "" Then
                                                    Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                                Else
                                                    Information &= OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString
                                                End If
                                            End If
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For

                            Case "REQUIRED_TO_SHIP"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "COMMENT"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "IS_CRITICAL"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "CUSTSP_LINE_NO"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "EXPECTED_DATE"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "CUSTOMER _CODE"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "REQUIRE_STATUS_CODE"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "REJECT_STATUS_CODE"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "OWNER_CODE"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If ColonnNameLigne <> Nothing Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            If Trim(Societe) <> Nothing Then
                                                Information &= Strings.LSet(Societe, PositionLeftLigne)
                                            Else
                                                Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                            End If
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            If Trim(Societe) <> Nothing Then
                                                Information &= Strings.LSet(Societe, PositionLeftLigne)
                                            Else
                                                Information &= Strings.LSet("", PositionLeftLigne)
                                            End If
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        If Trim(Societe) <> Nothing Then
                                            Information &= Strings.LSet(Societe, PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                End If
                                Continue For
                            Case "PREF_REQ_STATUS_CODE"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "CONTAINER_NO"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "QUANTITY_TO_RESERVE"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "SERIE_CONTROL"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "SALE_PRICE"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "LABELS_NO"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "CONTAINERTYPE_CODE"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For

                            Case "SUBSTITUTE_PROD_CODE"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "PREFERABLE_UOM_CODE"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "MIN_SHELF_LIFE"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "ROUND_CONTROL"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "MIXED_LOGISTIC"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "QUALITY"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                            Case "BEST_BEFORE_DATE"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For

                            Case "LOCATION_LABEL"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "CATCH_WEIGHT"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "SUSTITUTE_QTTY"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "DISABLE_ALT_PRODUCT"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "ALT_UOM_CODE"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "FROM_FACTOR"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "TO_FACTOR"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "SIZE_CODE"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "COLOR_CODE"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "SOURCE"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "VERSION_CODE"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case "RESERVE_MARK"
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                        End If
                                    ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                        End If
                                    Else
                                        'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet("", PositionLeftLigne)
                                        End If
                                    End If
                                ElseIf OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                    End If
                                Else
                                    'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                    If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    Else
                                        Information &= Strings.LSet("", PositionLeftLigne)
                                    End If
                                End If
                                Continue For
                            Case Else
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                    If OledatableSchemaSageDetails.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                        ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans à la variable colonne
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString = "V" Then 'si ce genre de type de format prevu
                                            Dim ValeurLues As String = OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString ' recuperation de la valeur
                                            If ValeurLues = "" Then ' si la valriable ne comtien aucune donne alors on ce rabat vers la valeur pas defaut
                                                If OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then ' s'il ya une valeur pas defaut
                                                    Information &= Strings.LSet("[" & OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString & "]", ValeurLues.Length + 2)
                                                Else
                                                    Information &= Strings.LSet("[]", ValeurLues.Length + 2)
                                                End If
                                            Else
                                                Information &= Strings.LSet("[" & ValeurLues & "]", ValeurLues.Length + 2)
                                            End If
                                        Else
                                            If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Else
                                                Information &= Strings.LSet(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, PositionLeftLigne)
                                            End If
                                        End If
                                    Else
                                        If OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLues = OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLues & "]", ValeurLues.Length + 2)
                                            Else
                                                If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                                End If
                                            End If
                                        Else
                                            'un Mapping Sage existe mais ne possede pas de valeur que retourne la requete et ne possede pas de valeur pas defaut
                                            If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLues = ""
                                                Information &= Strings.LSet("[" & ValeurLues & "]", ValeurLues.Length + 2)
                                            Else
                                                If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                                End If
                                            End If
                                        End If
                                    End If
                                Else 'si le champ Sage ne possede pas de mapping correspondant a la colonne EasyWMS
                                    If OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut paramètrée
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                            Dim ValeurLues = OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString
                                            Information &= Strings.LSet("[" & valeurLue & "]", valeurLue.Length + 2)
                                        Else 'sinon le format n'est pas du type variable
                                            PositionLeftLigne = GetPosition(OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace à la fonction GetPosition qui prend en paramétre le format du champ
                                            If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                            Else
                                                Information &= Strings.LSet(OledatableSchemaLigne.Rows(Ligne).Item("DefaultValue").ToString, PositionLeftLigne)
                                            End If
                                        End If
                                    Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                        'Dim traitementAutres As String = 0
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString = "V" Then
                                            Dim ValeurLues = ""
                                            Information &= Strings.LSet("[" & ValeurLues & "]", ValeurLues.Length + 2)
                                        Else
                                            PositionLeftLigne = GetPosition(OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString)
                                            If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet("", PositionLeftLigne)
                                            Else
                                                Information &= Strings.LSet("", PositionLeftLigne)
                                            End If
                                        End If
                                    End If
                                End If
                                Continue For
                        End Select
                    Next
                    Error_journal.WriteLine(Information)
                    Information = ""
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Creation Ligne")
        End Try
    End Sub
    Public Function GetPosition(ByVal Format As Object) As Integer
        Try

            Dim Position() As Object = Format.ToString.Split("(")(1).ToString.Split(")")(0).Split(".")
            If Position.Length = 2 Then
                Return Position(0)
            Else
                Return Position(0)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Lecture Position")
        End Try
    End Function


    Public Function RenvoiValeurLue(ByVal Valeur As Object, ByVal Format As Object) As Object
        Try
            Dim FormatFloat As String = ""
            Select Case Format.ToString.Split("(")(0)
                Case "N"
                    Dim Tableau() As String = Format.ToString.Split("(")(1).Split(".")
                    If Tableau.Length = 2 Then
                        For Cpte As Integer = 0 To Tableau(0) - (Tableau(1).Split(")")(0) + 1)
                            If Cpte = (Tableau(0) - (Tableau(1).Split(")")(0) + 1)) Then
                                FormatFloat &= "0."
                            Else
                                FormatFloat &= "0"
                            End If
                        Next
                        For Cpte As Integer = 0 To Tableau(1).Split(")")(0) - 1
                            FormatFloat &= "0"
                        Next

                        If IsNumeric(CDbl(Valeur.ToString.Replace(".", ","))) = True Then
                            Return CDbl(Valeur.ToString.Replace(".", ",")).ToString(FormatFloat).Replace(",", "")
                        Else
                            Return CDbl(Valeur.ToString.Replace(".", ",")).ToString(FormatFloat).Replace(",", "")
                        End If
                    Else
                        For Cpte As Integer = 0 To Tableau(0).Split(")")(0) - 1
                            FormatFloat &= "0"
                        Next
                        Return CDbl(Valeur.ToString.Replace(".", ",")).ToString(FormatFloat).Replace(",", "")
                        'Return Valeur
                    End If
                Case "D"
                    Dim dates As Object = Valeur.ToString().Split("/")
                    Dim Jour = dates(0)
                    Dim Mois = dates(1)
                    Dim ListeYearTime = dates(2)
                    Dim Year = ListeYearTime.ToString.Split(" ")(0)
                    Dim Timer = ListeYearTime.ToString.Split(" ")(1).Replace(":", "")
                    Valeur = Year & Mois & Jour & Timer
                    Return Valeur
            End Select
            Return Valeur
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Retourne la Valeur Lue")
            Return ""
        End Try
    End Function
    Dim EstChoisir As Boolean = False
    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        If EstChoisir Then
            EcritureEntete()
        End If
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            'PictureBox1.Visible = False
            If EstVide Or EstChoisir = False Then
                If Erreur = True Then
                    If OledatableSchemaSage.Rows.Count = 0 Then
                        lblInfos.Visible = True
                        lblInfos.Text = "Aucune donnée retournée par ce Filtre"
                    Else
                        lblInfos.Text = "Erreur trouvée consulter le fichier journal Ctrl+J" '"L'infos Libre<[" & Flagtampon & "]> n'existe pas dans Sage"
                        lblInfos.Visible = True
                    End If
                    DataListeIntegrer.Rows(selectid(ip)).Cells("Status").Value = My.Resources.criticalind_status
                Else
                    lblInfos.Text = "Aucune donnée retournée par ce Filtre"
                    lblInfos.Visible = True
                    DataListeIntegrer.Rows(selectid(ip)).Cells("Status").Value = My.Resources.accepter
                End If
            Else
                If ifRowError > 0 Then
                    DataListeIntegrer.Rows(selectid(ip)).Cells("Status").Value = My.Resources.criticalind_status
                    ifRowError = 0
                Else
                    DataListeIntegrer.Rows(selectid(ip)).Cells("Status").Value = My.Resources.accepter
                End If
            End If
            If ip < selectid.Length - 1 Then
                ip = ip + 1
                BackgroundWorker1.RunWorkerAsync()
            Else
                ip = 0
                nbreligne = 0
                selectindex = ""
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Hide()
    End Sub

    Private Sub RbtAchat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RbtAchat.Click
        Me.Text = "Extraction Bon de retour Fournisseur "
    End Sub

    Private Sub RbtStock_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RbtStock.Click
        Me.Text = "Extraction Transfért de Dépôt "
    End Sub

    Private Sub BackgroundWorker4_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker4.DoWork
        Try
            Dim OleAdaptaterschema As OleDbDataAdapter
            Dim OleSchemaDataset As DataSet
            DataListeIntegrer.Rows.Clear()
            OleAdaptaterschema = New OleDbDataAdapter("select * from PARAMETRE WHERE nomtype='COMMERCIAL'", OleConnenection)
            OleSchemaDataset = New DataSet
            OleAdaptaterschema.Fill(OleSchemaDataset)
            OledatableSchema = OleSchemaDataset.Tables(0)
        Catch ex As Exception
        End Try
    End Sub
    Private Sub BackgroundWorker4_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker4.RunWorkerCompleted
        Try
            AfficheSchemasConso()
        Catch ex As Exception
        End Try
    End Sub
    Private Sub AfficheSchemasConso()
        Dim i As Integer
        DataListeIntegrer.RowCount = OledatableSchema.Rows.Count
        For i = 0 To OledatableSchema.Rows.Count - 1
            DataListeIntegrer.Rows(i).Cells("Societe1").Value = OledatableSchema.Rows(i).Item("Societe")
            DataListeIntegrer.Rows(i).Cells("Chemin1").Value = OledatableSchema.Rows(i).Item("Chemin1")
            DataListeIntegrer.Rows(i).Cells("Type1").Value = OledatableSchema.Rows(i).Item("nomtype")
            DataListeIntegrer.Rows(i).Cells("UserSage1").Value = OledatableSchema.Rows(i).Item("UserSage")
            DataListeIntegrer.Rows(i).Cells("PasseSage1").Value = OledatableSchema.Rows(i).Item("PasseSage")
            DataListeIntegrer.Rows(i).Cells("bdd1").Value = OledatableSchema.Rows(i).Item("BaseDonnee")
            DataListeIntegrer.Rows(i).Cells("Serveur1").Value = OledatableSchema.Rows(i).Item("Serveur")
            DataListeIntegrer.Rows(i).Cells("NomUtil").Value = OledatableSchema.Rows(i).Item("NomUser")
            DataListeIntegrer.Rows(i).Cells("Mot").Value = OledatableSchema.Rows(i).Item("MotPas")
            DataListeIntegrer.Rows(i).Cells("choix").Value = True
        Next i
    End Sub
    Public Function IDernierDate(ByRef IDTraitement As Object) As Date
        Try
            Dim Insertion As String
            Dim OleCmdIns As OleDbCommand
            Dim OleAdaptater As OleDbDataAdapter
            Dim OleAfficheDataset As DataSet
            Dim Oledatable As DataTable
            OleAdaptater = New OleDbDataAdapter("select * from WEAT_ADATE WHERE IDDossier=" & IDTraitement & " And TypeExport='Export Commande Client'", OleConnenection)
            OleAfficheDataset = New DataSet
            OleAdaptater.Fill(OleAfficheDataset)
            Oledatable = OleAfficheDataset.Tables(0)
            If Oledatable.Rows.Count <> 0 Then
                Insertion = " UPDATE WEAT_ADATE SET DateDern='" & Strings.FormatDateTime(Now, Microsoft.VisualBasic.DateFormat.ShortDate) & "' WHERE IDDossier=" & IDTraitement & " And TypeExport='Export Commande Client'"
                OleCmdIns = New OleDbCommand(Insertion)
                OleCmdIns.Connection = OleConnenection
                OleCmdIns.ExecuteNonQuery()
            Else
                Insertion = " Insert Into WEAT_ADATE (TypeExport,IDDossier,DateDern) Values ('Export Commande Client'," & IDTraitement & ",'" & Strings.FormatDateTime(Now, Microsoft.VisualBasic.DateFormat.ShortDate) & "')"
                OleCmdIns = New OleDbCommand(Insertion)
                OleCmdIns.Connection = OleConnenection
                OleCmdIns.ExecuteNonQuery()
            End If
        Catch ex As Exception
        End Try
    End Function
    Private Function ArticleRecemmentModifier(ByRef DatDerntraitement As Date, ByRef OleArtModifierDt As DataTable, ByRef m As Integer) As Boolean
        Dim MustModified As Boolean = False
        Try
            If DatDerntraitement <= OleArtModifierDt.Rows(m).Item("CBModification") Then
                MustModified = True
            End If
        Catch ex As Exception
        End Try
        Return MustModified
    End Function
    Public Function RDernierDate(ByRef IDTraitement As Object) As Date
        Try
            Dim OleAdaptater As OleDbDataAdapter
            Dim OleAfficheDataset As DataSet
            Dim Oledatable As DataTable
            OleAdaptater = New OleDbDataAdapter("select * from WEAT_ADATE WHERE IDDossier=" & IDTraitement & " And TypeExport='Export Commande Client'", OleConnenection)
            OleAfficheDataset = New DataSet
            OleAdaptater.Fill(OleAfficheDataset)
            Oledatable = OleAfficheDataset.Tables(0)
            If Oledatable.Rows.Count <> 0 Then
                RDernierDate = CDate(Strings.FormatDateTime(Oledatable.Rows(0).Item("DateDern"), Microsoft.VisualBasic.DateFormat.ShortDate))
            Else
                RDernierDate = CDate(Strings.FormatDateTime(Now, Microsoft.VisualBasic.DateFormat.ShortDate))
            End If
        Catch ex As Exception
        End Try
    End Function
    Public StatutDoc As Integer
    Private Sub CmbStatut_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbStatut.SelectedIndexChanged
        Select Case CmbStatut.Text
            Case "Saisi"
                StatutDoc = 0
            Case "Confirmé"
                StatutDoc = 1
            Case "Réceptionné"
                StatutDoc = 2
        End Select
    End Sub
End Class