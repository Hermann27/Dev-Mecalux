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
    Private Sub BtnModif_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnModif.Click
        ListBox.Items.Clear()
        ContinuTraitement = 0
        NbInfosLibre = 0
        NbInfosLibreVue = 0
        ListeChampSageLigne = ""
        Try
            VerificationChampEnteteObligatoire(True, False)

            If ContinuTraitement = 4 Then
                VerificateurChampObligatoire(False, True)
                If ContinuTraitement = 7 Then
                    EstInfosLibre(True, False, True)
                    If NbInfosLibre = NbInfosLibreVue Then 'NbInfosLibreVue
                        EstInfosLibre(False, True, True)
                        If NbInfosLibre = NbInfosLibreVue Then 'NbInfosLibreVue
                            Try
                                PictureBox1.Visible = True
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
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Traitement Fin Erreur")
        End Try

    End Sub

    Private Sub FrmExtraction_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LirefichierConfig()
        Connected()
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

    Public Sub ExtractionBonCommandeClient(ByVal SelectChamp As String)
        Dim Information As String = ""
        Dim ExistColonne As Boolean = False
        Dim ColonnName As String = ""
        Dim PositionLeft As Object = Nothing
        Try
            If ExcelConnect(NomServersql, NomBaseCpta, Nom_Utilsql, Mot_Passql) Then
                If RbARcreation.IsChecked Then
                    If SelectChamp IsNot Nothing Then
                        If RbtVente.IsChecked Then
                            OleAdaptaterschemaSage = New OleDbDataAdapter("SELECT " & SelectChamp & " FROM F_DOCENTETE WHERE DO_DOMAINE=0 AND DO_TYPE=1 AND CAST(DO_Date AS Date) = CAST(CBModification AS Date) ORDER BY F_DOCENTETE.DO_PIECE ", OleExcelConnect)
                        ElseIf RbtAchat.IsChecked Then
                            OleAdaptaterschemaSage = New OleDbDataAdapter("SELECT " & SelectChamp & " FROM F_DOCENTETE WHERE DO_DOMAINE=1 AND DO_TYPE=14 AND CAST(DO_Date AS Date) = CAST(CBModification AS Date) ORDER BY F_DOCENTETE.DO_PIECE ", OleExcelConnect)
                        ElseIf RbtStock.IsChecked Then
                            OleAdaptaterschemaSage = New OleDbDataAdapter("SELECT " & SelectChamp & " FROM F_DOCENTETE WHERE DO_DOMAINE=2 AND DO_TYPE=23 AND CAST(DO_Date AS Date) = CAST(CBModification AS Date) ORDER BY F_DOCENTETE.DO_PIECE ", OleExcelConnect)
                        End If
                    Else
                        If RbtVente.IsChecked Then
                            OleAdaptaterschemaSage = New OleDbDataAdapter("SELECT * FROM F_DOCENTETE WHERE DO_DOMAINE=0 AND DO_TYPE=1 AND CAST(DO_Date AS Date) = CAST(CBModification AS Date) ORDER BY F_DOCENTETE.DO_PIECE ", OleExcelConnect)
                        ElseIf RbtAchat.IsChecked Then
                            OleAdaptaterschemaSage = New OleDbDataAdapter("SELECT * FROM F_DOCENTETE WHERE DO_DOMAINE=1 AND DO_TYPE=14 AND CAST(DO_Date AS Date) = CAST(CBModification AS Date) ORDER BY F_DOCENTETE.DO_PIECE ", OleExcelConnect)
                        ElseIf RbtStock.IsChecked Then
                            OleAdaptaterschemaSage = New OleDbDataAdapter("SELECT * FROM F_DOCENTETE WHERE DO_DOMAINE=2 AND DO_TYPE=23 AND CAST(DO_Date AS Date) = CAST(CBModification AS Date) ORDER BY F_DOCENTETE.DO_PIECE ", OleExcelConnect)
                        End If
                    End If
                ElseIf RbARmodifier.IsChecked Then
                    If SelectChamp IsNot Nothing Then
                        If RbtVente.IsChecked Then
                            OleAdaptaterschemaSage = New OleDbDataAdapter("SELECT " & SelectChamp & " FROM F_DOCENTETE WHERE DO_DOMAINE=0 AND DO_TYPE=1 AND DO_Date <> CBModification ORDER BY F_DOCENTETE.DO_PIECE ", OleExcelConnect)
                        ElseIf RbtAchat.IsChecked Then
                            OleAdaptaterschemaSage = New OleDbDataAdapter("SELECT " & SelectChamp & " FROM F_DOCENTETE WHERE DO_DOMAINE=1 AND DO_TYPE=14 AND DO_Date <> CBModification ORDER BY F_DOCENTETE.DO_PIECE ", OleExcelConnect)
                        ElseIf RbtStock.IsChecked Then
                            OleAdaptaterschemaSage = New OleDbDataAdapter("SELECT " & SelectChamp & " FROM F_DOCENTETE WHERE DO_DOMAINE=2 AND DO_TYPE=23 AND DO_Date <> CBModification ORDER BY F_DOCENTETE.DO_PIECE ", OleExcelConnect)
                        End If
                    Else
                        If RbtVente.IsChecked Then
                            OleAdaptaterschemaSage = New OleDbDataAdapter("SELECT * FROM F_DOCENTETE WHERE DO_DOMAINE=0 AND DO_TYPE=1 AND DO_Date <> CBModification ORDER BY F_DOCENTETE.DO_PIECE ", OleExcelConnect)
                        ElseIf RbtAchat.IsChecked Then
                            OleAdaptaterschemaSage = New OleDbDataAdapter("SELECT * FROM F_DOCENTETE WHERE DO_DOMAINE=1 AND DO_TYPE=14 AND DO_Date <> CBModification ORDER BY F_DOCENTETE.DO_PIECE ", OleExcelConnect)
                        ElseIf RbtStock.IsChecked Then
                            OleAdaptaterschemaSage = New OleDbDataAdapter("SELECT * FROM F_DOCENTETE WHERE DO_DOMAINE=2 AND DO_TYPE=23 AND DO_Date <> CBModification ORDER BY F_DOCENTETE.DO_PIECE ", OleExcelConnect)
                        End If
                    End If
                ElseIf (RbAR2action.IsChecked) Then
                    If SelectChamp IsNot Nothing Then
                        If RbtVente.IsChecked Then
                            OleAdaptaterschemaSage = New OleDbDataAdapter("SELECT " & SelectChamp & " FROM F_DOCENTETE WHERE DO_DOMAINE=0 AND DO_TYPE=1  ORDER BY F_DOCENTETE.DO_PIECE ", OleExcelConnect)
                        ElseIf RbtAchat.IsChecked Then
                            OleAdaptaterschemaSage = New OleDbDataAdapter("SELECT " & SelectChamp & " FROM F_DOCENTETE WHERE DO_DOMAINE=1 AND DO_TYPE=14 ORDER BY F_DOCENTETE.DO_PIECE ", OleExcelConnect)
                        ElseIf RbtStock.IsChecked Then
                            OleAdaptaterschemaSage = New OleDbDataAdapter("SELECT " & SelectChamp & " FROM F_DOCENTETE WHERE DO_DOMAINE=2 AND DO_TYPE=23 ORDER BY F_DOCENTETE.DO_PIECE ", OleExcelConnect)
                        End If
                    Else
                        If RbtVente.IsChecked Then
                            OleAdaptaterschemaSage = New OleDbDataAdapter("SELECT * FROM F_DOCENTETE WHERE DO_DOMAINE=0 AND DO_TYPE=1  ORDER BY F_DOCENTETE.DO_PIECE ", OleExcelConnect)
                        ElseIf RbtAchat.IsChecked Then
                            OleAdaptaterschemaSage = New OleDbDataAdapter("SELECT * FROM F_DOCENTETE WHERE DO_DOMAINE=1 AND DO_TYPE=14 ORDER BY F_DOCENTETE.DO_PIECE ", OleExcelConnect)
                        ElseIf RbtStock.IsChecked Then
                            OleAdaptaterschemaSage = New OleDbDataAdapter("SELECT * FROM F_DOCENTETE WHERE DO_DOMAINE=2 AND DO_TYPE=23 ORDER BY F_DOCENTETE.DO_PIECE ", OleExcelConnect)
                        End If
                    End If
                End If

                OleSchemaDatasetSage = New DataSet
                OleAdaptaterschemaSage.Fill(OleSchemaDatasetSage)
                OledatableSchemaSage = OleSchemaDatasetSage.Tables(0)

                OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='SOR' AND Entete=True ORDER BY ORDRE ", OleConnenection)
                OleSchemaDataset = New DataSet
                OleAdaptaterschema.Fill(OleSchemaDataset)
                OledatableSchema = OleSchemaDataset.Tables(0)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Chargement des Bon de Commande Client")
        End Try
    End Sub
    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExtractionBonCommandeClient(ListeChampSageEntete)
    End Sub
    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        BackgroundWorker2.RunWorkerAsync()
    End Sub

    Public Function EcritureEntete() As Boolean
        Dim Information As String = ""
        Dim ExistColonne As Boolean = False
        Dim ColonnName As String = ""
        Dim PositionLeft As Object = Nothing
        Dim PositionLeftLigne As Integer = 0
        Dim ColonnNameLigne As String = ""
        Try
            If OledatableSchemaSage.Rows.Count <> 0 Then
                Error_journal = File.AppendText(PathsFileFormatiers & "SOR09" & Strings.Right(DateAndTime.Year(Now), 4) & Format(DateAndTime.Month(Now), "00") & Format(DateAndTime.Day(Now), "00") & Format(DateAndTime.Hour(Now), "00") & Format(DateAndTime.Minute(Now), "00") & Format(DateAndTime.Second(Now), "00") & ".txt")
                For i As Integer = 0 To OledatableSchemaSage.Rows.Count - 1
                    ListBox.Items.Add("Extraction " & Item & " Identificateur (Code) : <[" & OledatableSchemaSage.Rows(i).Item("DO_PIECE") & "]> ")
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
                                                Information &= Strings.LSet("CORDER", PositionLeft)
                                            Case "14"
                                                Information &= Strings.LSet("RETURN", PositionLeft)
                                            Case "23"
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
                            Case "CONTACT_NAME"
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
                            Case "CARRIER_CODE"
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
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
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
                Next
                Error_journal.Close()
            Else
                'cette requette ne possede pas ligne 
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Extraction données d'entete ")
            Return False
        End Try
    End Function

    Public Sub CreationLigne(ByVal valeurLue As Object)
        Dim PositionLeftLigne As Integer = 0
        Dim Information As String = ""
        Dim ColonnNameLigne As String = ""
        If RbARcreation.IsChecked Then
            If ListeChampSageLigne IsNot Nothing Then
                If RbtVente.IsChecked Then
                    OleAdaptaterschemaSageDetails = New OleDbDataAdapter("SELECT " & ListeChampSageLigne & " FROM F_DOCLIGNE WHERE F_DOCLIGNE.DO_DOMAINE=0 AND F_DOCLIGNE.DO_TYPE=1 AND F_DOCLIGNE.DO_PIECE='" & valeurLue & "' ORDER BY F_DOCLIGNE.DO_PIECE", OleExcelConnect)
                ElseIf RbtAchat.IsChecked Then
                    OleAdaptaterschemaSageDetails = New OleDbDataAdapter("SELECT " & ListeChampSageLigne & " FROM F_DOCLIGNE WHERE F_DOCLIGNE.DO_DOMAINE=1 AND F_DOCLIGNE.DO_TYPE=14 AND F_DOCLIGNE.DO_PIECE='" & valeurLue & "' ORDER BY F_DOCLIGNE.DO_PIECE", OleExcelConnect)
                ElseIf RbtStock.IsChecked Then
                    OleAdaptaterschemaSageDetails = New OleDbDataAdapter("SELECT " & ListeChampSageLigne & " FROM F_DOCLIGNE WHERE F_DOCLIGNE.DO_DOMAINE=2 AND F_DOCLIGNE.DO_TYPE=23 AND F_DOCLIGNE.DO_PIECE='" & valeurLue & "' ORDER BY F_DOCLIGNE.DO_PIECE", OleExcelConnect)
                End If
            Else
                If RbtVente.IsChecked Then
                    OleAdaptaterschemaSageDetails = New OleDbDataAdapter("SELECT * FROM F_DOCLIGNE WHERE F_DOCLIGNE.DO_DOMAINE=0 AND F_DOCLIGNE.DO_TYPE=1 AND F_DOCLIGNE.DO_PIECE='" & valeurLue & "' ORDER BY F_DOCLIGNE.DO_PIECE", OleExcelConnect)
                ElseIf RbtAchat.IsChecked Then
                    OleAdaptaterschemaSageDetails = New OleDbDataAdapter("SELECT * FROM F_DOCLIGNE WHERE F_DOCLIGNE.DO_DOMAINE=1 AND F_DOCLIGNE.DO_TYPE=14 AND F_DOCLIGNE.DO_PIECE='" & valeurLue & "' ORDER BY F_DOCLIGNE.DO_PIECE", OleExcelConnect)
                ElseIf RbtStock.IsChecked Then
                    OleAdaptaterschemaSageDetails = New OleDbDataAdapter("SELECT * FROM F_DOCLIGNE WHERE F_DOCLIGNE.DO_DOMAINE=2 AND F_DOCLIGNE.DO_TYPE=23 AND F_DOCLIGNE.DO_PIECE='" & valeurLue & "' ORDER BY F_DOCLIGNE.DO_PIECE", OleExcelConnect)
                End If
            End If
        ElseIf RbARmodifier.IsChecked Then
            If ListeChampSageLigne IsNot Nothing Then
                If RbtVente.IsChecked Then
                    OleAdaptaterschemaSageDetails = New OleDbDataAdapter("SELECT " & ListeChampSageLigne & " FROM F_DOCLIGNE WHERE F_DOCLIGNE.DO_DOMAINE=0 AND F_DOCLIGNE.DO_TYPE=1 AND F_DOCLIGNE.DO_PIECE='" & valeurLue & "' ORDER BY F_DOCLIGNE.DO_PIECE", OleExcelConnect)
                ElseIf RbtAchat.IsChecked Then
                    OleAdaptaterschemaSageDetails = New OleDbDataAdapter("SELECT " & ListeChampSageLigne & " FROM F_DOCLIGNE WHERE F_DOCLIGNE.DO_DOMAINE=1 AND F_DOCLIGNE.DO_TYPE=14 AND F_DOCLIGNE.DO_PIECE='" & valeurLue & "' ORDER BY F_DOCLIGNE.DO_PIECE", OleExcelConnect)
                ElseIf RbtStock.IsChecked Then
                    OleAdaptaterschemaSageDetails = New OleDbDataAdapter("SELECT " & ListeChampSageLigne & " FROM F_DOCLIGNE WHERE F_DOCLIGNE.DO_DOMAINE=2 AND F_DOCLIGNE.DO_TYPE=23 AND F_DOCLIGNE.DO_PIECE='" & valeurLue & "' ORDER BY F_DOCLIGNE.DO_PIECE", OleExcelConnect)
                End If
            Else
                If RbtVente.IsChecked Then
                    OleAdaptaterschemaSageDetails = New OleDbDataAdapter("SELECT * FROM F_DOCLIGNE WHERE F_DOCLIGNE.DO_DOMAINE=0 AND F_DOCLIGNE.DO_TYPE=1 AND F_DOCLIGNE.DO_PIECE='" & valeurLue & "' ORDER BY F_DOCLIGNE.DO_PIECE", OleExcelConnect)
                ElseIf RbtAchat.IsChecked Then
                    OleAdaptaterschemaSageDetails = New OleDbDataAdapter("SELECT * FROM F_DOCLIGNE WHERE F_DOCLIGNE.DO_DOMAINE=1 AND F_DOCLIGNE.DO_TYPE=14 AND F_DOCLIGNE.DO_PIECE='" & valeurLue & "' ORDER BY F_DOCLIGNE.DO_PIECE", OleExcelConnect)
                ElseIf RbtStock.IsChecked Then
                    OleAdaptaterschemaSageDetails = New OleDbDataAdapter("SELECT * FROM F_DOCLIGNE WHERE F_DOCLIGNE.DO_DOMAINE=2 AND F_DOCLIGNE.DO_TYPE=23 AND F_DOCLIGNE.DO_PIECE='" & valeurLue & "' ORDER BY F_DOCLIGNE.DO_PIECE", OleExcelConnect)
                End If
            End If
        ElseIf (RbAR2action.IsChecked) Then
            If ListeChampSageLigne IsNot Nothing Then
                If RbtVente.IsChecked Then
                    OleAdaptaterschemaSageDetails = New OleDbDataAdapter("SELECT " & ListeChampSageLigne & " FROM F_DOCLIGNE WHERE F_DOCLIGNE.DO_DOMAINE=0 AND F_DOCLIGNE.DO_TYPE=1 AND F_DOCLIGNE.DO_PIECE='" & valeurLue & "' ORDER BY F_DOCLIGNE.DO_PIECE", OleExcelConnect)
                ElseIf RbtAchat.IsChecked Then
                    OleAdaptaterschemaSageDetails = New OleDbDataAdapter("SELECT " & ListeChampSageLigne & " FROM F_DOCLIGNE WHERE F_DOCLIGNE.DO_DOMAINE=1 AND F_DOCLIGNE.DO_TYPE=14 AND F_DOCLIGNE.DO_PIECE='" & valeurLue & "' ORDER BY F_DOCLIGNE.DO_PIECE", OleExcelConnect)
                ElseIf RbtStock.IsChecked Then
                    OleAdaptaterschemaSageDetails = New OleDbDataAdapter("SELECT " & ListeChampSageLigne & " FROM F_DOCLIGNE WHERE F_DOCLIGNE.DO_DOMAINE=2 AND F_DOCLIGNE.DO_TYPE=23 AND F_DOCLIGNE.DO_PIECE='" & valeurLue & "' ORDER BY F_DOCLIGNE.DO_PIECE", OleExcelConnect)
                End If
            Else
                If RbtVente.IsChecked Then
                    OleAdaptaterschemaSageDetails = New OleDbDataAdapter("SELECT * FROM F_DOCLIGNE WHERE F_DOCLIGNE.DO_DOMAINE=0 AND F_DOCLIGNE.DO_TYPE=1 AND F_DOCLIGNE.DO_PIECE='" & valeurLue & "' ORDER BY F_DOCLIGNE.DO_PIECE", OleExcelConnect)
                ElseIf RbtAchat.IsChecked Then
                    OleAdaptaterschemaSageDetails = New OleDbDataAdapter("SELECT * FROM F_DOCLIGNE WHERE F_DOCLIGNE.DO_DOMAINE=1 AND F_DOCLIGNE.DO_TYPE=14 AND F_DOCLIGNE.DO_PIECE='" & valeurLue & "' ORDER BY F_DOCLIGNE.DO_PIECE", OleExcelConnect)
                ElseIf RbtStock.IsChecked Then
                    OleAdaptaterschemaSageDetails = New OleDbDataAdapter("SELECT * FROM F_DOCLIGNE WHERE F_DOCLIGNE.DO_DOMAINE=2 AND F_DOCLIGNE.DO_TYPE=23 AND F_DOCLIGNE.DO_PIECE='" & valeurLue & "' ORDER BY F_DOCLIGNE.DO_PIECE", OleExcelConnect)
                End If
            End If
        End If
        OleSchemaDatasetSageDetails = New DataSet
        OleAdaptaterschemaSageDetails.Fill(OleSchemaDatasetSageDetails)
        OledatableSchemaSageDetails = OleSchemaDatasetSageDetails.Tables(0)

        OleAdaptaterschemaLigne = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='SOR' AND Ligne=True", OleConnenection)
        OleSchemaDatasetLigne = New DataSet
        OleAdaptaterschemaLigne.Fill(OleSchemaDatasetLigne)
        OledatableSchemaLigne = OleSchemaDatasetLigne.Tables(0)
        Try
            If OledatableSchemaSageDetails.Rows.Count <> 0 Then
                For i As Integer = 0 To OledatableSchemaSageDetails.Rows.Count - 1
                    ListBox.Items.Add("     0---> Detail du Document N° Piece : " & valeurLue & " Ligne du Document Extrait :[" & OledatableSchemaSageDetails.Rows(i).Item("DL_LIGNE") & "]")
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
                        For Cpte As Integer = 0 To Tableau(0).Split(")")(0)
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
    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        'TraitementExtraction()
        EcritureEntete()

    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        PictureBox1.Visible = False
        If RbtAchat.IsChecked Then
            Me.Alerte.ContentText = "Fin d'extraction des bons de retours fournisseurs"
            Me.Alerte.Show()
        ElseIf RbtStock.IsChecked Then
            Me.Alerte.ContentText = "Fin d'extraction des Transferts de dépôts"
            Me.Alerte.Show()
        ElseIf RbtVente.IsChecked Then
            Me.Alerte.ContentText = "Fin d'extraction des bons de commandes clients"
            Me.Alerte.Show()
        End If

    End Sub

    Private Sub RadButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadButton2.Click
        Me.Close()
    End Sub

    Private Sub RadButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadButton1.Click
        Me.Hide()
    End Sub
End Class
