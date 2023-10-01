Imports System.Data.OleDb
Imports System.IO
Public Class FrmExtraction
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

    Private Sub BtnModif_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnModif.Click
        ListBox.Items.Clear()
        ContinuTraitement = 0
        NbInfosLibre = 0
        NbInfosLibreVue = 0
        Try
            VerificateurChampObligatoire(True, False)
            If ContinuTraitement = 7 Then
                EstInfosLibre(True, False, True)
                If NbInfosLibre = NbInfosLibreVue Then 'NbInfosLibreVue
                    Try

                        PictureBox1.Visible = True
                        ListeChampSageEntete = RecuperationColonneSage(True, False)
                        lblSne.Text = "Scénario Extraction des article"
                        ListeChampSageEntete = ListeChampSageEntete.Substring(0, ListeChampSageEntete.Length - 1)
                        BackgroundWorker1.RunWorkerAsync()
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        
    End Sub

    Private Sub FrmExtraction_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LirefichierConfig()
        Connected()
    End Sub
    Public ContinuTraitement As Integer = 0
    Public Sub VerificateurChampObligatoire(ByVal Entete As Boolean, ByVal Ligne As Boolean)
        OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='PRO' AND Entete=" & Entete & " AND Ligne=" & Ligne, OleConnenection)
        OleSchemaDataset = New DataSet
        OleAdaptaterschema.Fill(OleSchemaDataset)
        OledatableSchema = OleSchemaDataset.Tables(0)
        lblSne.Text = "Verification conformite article"
        ListBox.Items.Add("<-----------------------------------------------------------------DEBUT DU TRAITEMENT DE LA VERIFICATION DES CHAMPS----------------------------------------------------------------->")
        If OledatableSchema.Rows.Count <> 0 Then
            For i As Integer = 0 To OledatableSchema.Rows.Count - 1
                Select Case OledatableSchema.Rows(i).Item("Cols")
                    Case "Operation"
                        If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False And OledatableSchema.Rows(i).Item("ChampSage") <> "" Then
                            ContinuTraitement += 1
                            ListBox.Items.Add("Un Mapping Sage {" & OledatableSchema.Rows(i).Item("ChampSage") & "} a été appliquer sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}")
                        Else
                            If IsDBNull(OledatableSchema.Rows(i).Item("DefaultValue")) = False Then
                                ListBox.Items.Add("Une Valeur par defaut a été paramètrée sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}")
                                ContinuTraitement += 1
                            Else
                                ListBox.Items.Add("Aucun Mapping Sage n'est paramètre sur  le champ {" & OledatableSchema.Rows(i).Item("Cols") & "} EasyWMS  ni de valeur par defaut n'a ete paramètrée {Ce Champ est Obligatoire}")
                            End If
                        End If
                    Case "PRODUCT_CODE"
                        If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False And OledatableSchema.Rows(i).Item("ChampSage") <> "" Then
                            ContinuTraitement += 1
                            ListBox.Items.Add("Un Mapping Sage {" & OledatableSchema.Rows(i).Item("ChampSage") & "} a été appliquer sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}")
                        Else
                            If IsDBNull(OledatableSchema.Rows(i).Item("DefaultValue")) = False Then
                                ListBox.Items.Add("Une Valeur par defaut a été paramètrée sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}")
                                ContinuTraitement += 1
                            Else
                                ListBox.Items.Add("Aucun Mapping Sage n'est paramètre sur  le champ {" & OledatableSchema.Rows(i).Item("Cols") & "} EasyWMS  ni de valeur par defaut n'a ete paramètrée {Ce Champ est Obligatoire}")
                            End If
                        End If
                    Case "OWNER_CODE"
                        If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False And OledatableSchema.Rows(i).Item("ChampSage") <> "" Then
                            ContinuTraitement += 1
                            ListBox.Items.Add("Un Mapping Sage {" & OledatableSchema.Rows(i).Item("ChampSage") & "} a été appliquer sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}")
                        Else
                            If IsDBNull(OledatableSchema.Rows(i).Item("DefaultValue")) = False Then
                                ListBox.Items.Add("Une Valeur par defaut a été paramètrée sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}")
                                ContinuTraitement += 1
                            Else
                                ListBox.Items.Add("Aucun Mapping Sage n'est paramètre sur  le champ {" & OledatableSchema.Rows(i).Item("Cols") & "} EasyWMS  ni de valeur par defaut n'a ete paramètrée {Ce Champ est Obligatoire}")
                            End If
                        End If
                    Case "MATERIAL_ABC_CODE"
                        If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False And OledatableSchema.Rows(i).Item("ChampSage") <> "" Then
                            ContinuTraitement += 1
                            ListBox.Items.Add("Un Mapping Sage {" & OledatableSchema.Rows(i).Item("ChampSage") & "} a été appliquer sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}")
                        Else
                            If IsDBNull(OledatableSchema.Rows(i).Item("DefaultValue")) = False Then
                                ListBox.Items.Add("Une Valeur par defaut a été paramètrée sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}")
                                ContinuTraitement += 1
                            Else
                                ListBox.Items.Add("Aucun Mapping Sage n'est paramètre sur  le champ {" & OledatableSchema.Rows(i).Item("Cols") & "} EasyWMS  ni de valeur par defaut n'a ete paramètrée {Ce Champ est Obligatoire}")
                            End If
                        End If
                    Case "UOM_BASE_CODE"
                        If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False And OledatableSchema.Rows(i).Item("ChampSage") <> "" Then
                            ContinuTraitement += 1
                            ListBox.Items.Add("Un Mapping Sage {" & OledatableSchema.Rows(i).Item("ChampSage") & "} a été appliquer sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}")
                        Else
                            If IsDBNull(OledatableSchema.Rows(i).Item("DefaultValue")) = False Then
                                ListBox.Items.Add("Une Valeur par defaut a été paramètrée sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}")
                                ContinuTraitement += 1
                            Else
                                ListBox.Items.Add("Aucun Mapping Sage n'est paramètre sur  le champ {" & OledatableSchema.Rows(i).Item("Cols") & "} EasyWMS  ni de valeur par defaut n'a ete paramètrée {Ce Champ est Obligatoire}")
                            End If
                        End If
                    Case "DESC_UOM_BASE_CODE"
                        If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False And OledatableSchema.Rows(i).Item("ChampSage") <> "" Then
                            ContinuTraitement += 1
                            ListBox.Items.Add("Un Mapping Sage {" & OledatableSchema.Rows(i).Item("ChampSage") & "} a été appliquer sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}")
                        Else
                            If IsDBNull(OledatableSchema.Rows(i).Item("DefaultValue")) = False Then
                                ListBox.Items.Add("Une Valeur par defaut a été paramètrée sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}")
                                ContinuTraitement += 1
                            Else
                                ListBox.Items.Add("Aucun Mapping Sage n'est paramètre sur  le champ {" & OledatableSchema.Rows(i).Item("Cols") & "} EasyWMS  ni de valeur par defaut n'a ete paramètrée {Ce Champ est Obligatoire}")
                            End If
                        End If
                    Case "LINE_COUNT"
                        If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False And OledatableSchema.Rows(i).Item("ChampSage") <> "" Then
                            ContinuTraitement += 1
                            ListBox.Items.Add("Un Mapping Sage {" & OledatableSchema.Rows(i).Item("ChampSage") & "} a été appliquer sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}")
                        Else
                            If IsDBNull(OledatableSchema.Rows(i).Item("DefaultValue")) = False Then
                                ListBox.Items.Add("Une Valeur par defaut a été paramètrée sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}")
                                ContinuTraitement += 1
                            Else
                                ListBox.Items.Add("Aucun Mapping Sage n'est paramètre sur  le champ {" & OledatableSchema.Rows(i).Item("Cols") & "} EasyWMS  ni de valeur par defaut n'a ete paramètrée {Ce Champ est Obligatoire}")
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
        ListBox.Items.Add("<-----------------------------------------------------------------DEBUT DU TRAITEMENT DE LA RECUPERATION  DES CHAMPS MAPPES----------------------------------------------------------------->")
        Dim NbEntete As Integer = 0
        Dim NbLigne As Integer = 0
        Dim NbInfosLibres As Integer = 0
        Dim OleAdaptaterschema As OleDbDataAdapter
        Dim OleSchemaDataset As DataSet
        Dim OledatableSchema As DataTable
        Dim ListeChampSage As String = ""
        If Entete Then
            OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='PRO' AND (Entete=" & Entete & " OR InfosLibre=true) AND Ligne=False", OleConnenection)
        Else
            OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='PRO' AND (Ligne=" & Ligne & " OR InfosLibre=true) AND Entete=False", OleConnenection)
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
                If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False Then
                    If Entete Then
                        If i = OledatableSchema.Rows.Count - 1 Then
                            If OledatableSchema.Rows(i).Item("ChampSage") <> "" Then
                                ListeChampSage &= OledatableSchema.Rows(i).Item("ChampSage")
                                ListBox.Items.Add("Recuperation de la colonne en Entete : " & OledatableSchema.Rows(i).Item("ChampSage"))
                            Else
                                ListBox.Items.Add("Le Champ {" & OledatableSchema.Rows(i).Item("Cols") & "} EasyWMS ne possede pas de Mapping Sage infos en Entete de l'article ")
                            End If
                        Else
                            If OledatableSchema.Rows(i).Item("ChampSage") <> "" Then
                                ListeChampSage &= OledatableSchema.Rows(i).Item("ChampSage") & ","
                                ListBox.Items.Add("Recuperation de la colonne en Entete : " & OledatableSchema.Rows(i).Item("ChampSage"))
                            Else
                                ListBox.Items.Add("Le Champ {" & OledatableSchema.Rows(i).Item("Cols") & "} EasyWMS ne possede pas de Mapping Sage infos en Entete de l'article ")
                            End If
                        End If
                    Else
                        If i = OledatableSchema.Rows.Count - 1 Then
                            If OledatableSchema.Rows(i).Item("ChampSage") <> "" Then
                                ListeChampSage &= OledatableSchema.Rows(i).Item("ChampSage")
                                ListBox.Items.Add("Recuperation de la colonne en Ligne : " & OledatableSchema.Rows(i).Item("ChampSage"))
                            Else
                                ListBox.Items.Add("Le Champ {" & OledatableSchema.Rows(i).Item("Cols") & "} EasyWMS ne possede pas de Mapping Sage infos en Ligne d'article ")
                            End If
                        Else
                            If OledatableSchema.Rows(i).Item("ChampSage") <> "" Then
                                ListeChampSage &= OledatableSchema.Rows(i).Item("ChampSage") & ","
                                ListBox.Items.Add("Recuperation de la colonne en Ligne : " & OledatableSchema.Rows(i).Item("ChampSage"))
                            Else
                                ListBox.Items.Add("Le Champ {" & OledatableSchema.Rows(i).Item("Cols") & "} EasyWMS ne possede pas de Mapping Sage infos en Ligne d'article ")
                            End If
                        End If
                    End If
                Else
                    ListBox.Items.Add("<Le Champ {" & OledatableSchema.Rows(i).Item("Cols") & "} EasyWMS ne possede pas de Mapping Sage>")
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
        lblSne.Text = "Scenario verification infos libre"
        Try
            OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='PRO' AND Entete=" & Entete & " AND InfosLibre=" & InfosLibre & " AND Ligne=" & Ligne, OleConnenection)
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
                                If ExisteInfosLibre(OledatableSchema.Rows(i).Item("ChampSage")) = False Then
                                    ListBox.Items.Add("la zone infos libre {" & OledatableSchema.Rows(i).Item("ChampSage") & "} en Entete de l'article n'existe pas dans Sage")
                                Else
                                    ListBox.Items.Add("Traitement  de la zone infos libre {" & OledatableSchema.Rows(i).Item("ChampSage") & "} en Entete de l'article Existe [OK] dans Sage")
                                End If
                            Else
                                ListBox.Items.Add("<--Le Champ indiquant l'infos libre est couché mais ne possede pas de mapping Sage-->")
                            End If
                        ElseIf Ligne = True Then
                            If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False Then
                                If ExisteInfosLibre(OledatableSchema.Rows(i).Item("ChampSage")) = False Then
                                    ListBox.Items.Add("la zone infos libre {" & OledatableSchema.Rows(i).Item("ChampSage") & "} en Ligne de l'article n'existe pas dans Sage")
                                Else
                                    ListBox.Items.Add("Traitement  de la zone infos libre {" & OledatableSchema.Rows(i).Item("ChampSage") & "} en Entete de l'article Existe [OK] dans Sage")
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
    Public Function ExisteInfosLibre(ByVal InfosLibre As String) As Boolean
        Try
            If ExcelConnect(NomServersql, NomBaseCpta, Nom_Utilsql, Mot_Passql) Then
                OleAdaptaterschemaSage = New OleDbDataAdapter("select * from cbSysLibre WHERE CB_File='F_ARTICLE' And CB_Name='" & Join(Split(Trim(InfosLibre), "'"), "''") & "'", OleExcelConnect)
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

    Public Sub ExtractionArticle(ByVal SelectChamp As String)
        Dim Information As String = ""
        Dim ExistColonne As Boolean = False
        Dim ColonnName As String = ""
        Dim PositionLeft As Object = Nothing

        Try
            If ExcelConnect(NomServersql, NomBaseCpta, Nom_Utilsql, Mot_Passql) Then
                If RbARcreation.IsChecked Then
                    If SelectChamp IsNot Nothing Then
                        OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & " FROM F_ARTICLE WHERE AR_SuiviStock<>0 AND AR_DateCreation = AR_DateModif ORDER BY F_ARTICLE.AR_REF ", OleExcelConnect)
                    Else
                        OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock<>0 AND AR_DateCreation = AR_DateModif ORDER BY F_ARTICLE.AR_REF ", OleExcelConnect)
                    End If
                ElseIf RbARmodifier.IsChecked Then
                    If SelectChamp IsNot Nothing Then
                        OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & " FROM F_ARTICLE WHERE AR_SuiviStock<>0 AND AR_DateCreation <> AR_DateModif  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnect)
                    Else
                        OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock<>0 AND AR_DateCreation <> AR_DateModif ORDER BY F_ARTICLE.AR_REF ", OleExcelConnect)
                    End If
                ElseIf (RbAR2action.IsChecked) Then
                    If SelectChamp IsNot Nothing Then
                        OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & " FROM F_ARTICLE WHERE AR_SuiviStock<>0 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnect)
                    Else
                        OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock<>0 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnect)
                    End If
                End If

                OleSchemaDatasetSage = New DataSet
                OleAdaptaterschemaSage.Fill(OleSchemaDatasetSage)
                OledatableSchemaSage = OleSchemaDatasetSage.Tables(0)

                OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='PRO' AND Entete=True ", OleConnenection)
                OleSchemaDataset = New DataSet
                OleAdaptaterschema.Fill(OleSchemaDataset)
                OledatableSchema = OleSchemaDataset.Tables(0)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Sub CreationLigne(ByVal valeurLue As Object)
        Dim PositionLeftLigne As Integer = 0
        Dim Information As String = ""
        Dim ColonnNameLigne As String = ""
        OleAdaptaterschemaSageDetails = New OleDbDataAdapter("select * FROM F_CONDITION,F_ARTICLE WHERE F_CONDITION.AR_REF=F_ARTICLE.AR_REF AND F_CONDITION.AR_REF='" & valeurLue & "'", OleExcelConnect)
        OleSchemaDatasetSageDetails = New DataSet
        OleAdaptaterschemaSageDetails.Fill(OleSchemaDatasetSageDetails)
        OledatableSchemaSageDetails = OleSchemaDatasetSageDetails.Tables(0)

        OleAdaptaterschemaLigne = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='PRO' AND Ligne=True", OleConnenection)
        OleSchemaDatasetLigne = New DataSet
        OleAdaptaterschemaLigne.Fill(OleSchemaDatasetLigne)
        OledatableSchemaLigne = OleSchemaDatasetLigne.Tables(0)
        Try
            If OledatableSchemaSageDetails.Rows.Count <> 0 Then
                For i As Integer = 0 To OledatableSchemaSageDetails.Rows.Count - 1
                    ListBox.Items.Add("     0---> Extraction detail de l'article  Code de l'article : " & valeurLue & " Code de la Condition :[" & OledatableSchemaSageDetails.Rows(i).Item("CO_NO") & "]")
                    For Ligne As Integer = 0 To OledatableSchemaLigne.Rows.Count - 1
                        If Ligne = 0 Then
                            Information &= Strings.LSet("/", 1)
                        ElseIf Ligne = 1 Then
                            Information &= Strings.LSet("M", 1)
                        End If
                        PositionLeftLigne = GetPosition(OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString)
                        Dim oo = OledatableSchemaLigne.Rows(Ligne).Item("Cols").ToString
                        Select Case OledatableSchemaLigne.Rows(Ligne).Item("Cols").ToString
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
                            Case "UOM_ DESC"
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

                            Case "PACK_WEIGHT"
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
                            Case "LENGTH"
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
                            Case "WIDTH"
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

                            Case "HEIGHT"
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
                            Case "PREP_TIME"
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

                            Case "DESC_CONTAINERTYPR_CODE"
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
                            Case "QUANTITY_CONTTYPE"
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
                            Case "TIE"
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
                            Case "HIGH"
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
                            Case "PICK_NEG_PERCENT"
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
                            Case "PICK_COMP_PERCENT"
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
                            Case "COMPLETE_PERCENT"
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
                            Case "PREFERABLE"
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
                            Case "REQUIRED"
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
                            Case "LOCK_FROM"
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
                            Case "LOCK_TO"
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
                            Case "MIN_QUANTITY"
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
                            Case "CONTAINER_HEIGHT"
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
                            Case "MIN_QUANTITY_ROUND_PERCENT"
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
                            Case "MIN_QUANTITY_MAX_PERCENT"
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
                        End Select
                    Next
                    Error_journal.WriteLine(Information)
                    Information = ""
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExtractionArticle(ListeChampSageEntete)
    End Sub
    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        BackgroundWorker2.RunWorkerAsync()
    End Sub
    
    Public Function EcritureEntete() As Boolean
        Dim Information As String = ""
        Dim ExistColonne As Boolean = False
        Dim ColonnName As String = ""
        Dim PositionLeft As Object = Nothing
        Error_journal = File.AppendText(PathsFileFormatiers & "PRO09" & Strings.Right(DateAndTime.Year(Now), 4) & Format(DateAndTime.Month(Now), "00") & Format(DateAndTime.Day(Now), "00") & Format(DateAndTime.Hour(Now), "00") & Format(DateAndTime.Minute(Now), "00") & Format(DateAndTime.Second(Now), "00") & ".txt")
        Try
            If OledatableSchemaSage.Rows.Count <> 0 Then
                For i As Integer = 0 To OledatableSchemaSage.Rows.Count - 1
                    ListBox.Items.Add("Extraction de l'article : " & OledatableSchemaSage.Rows(i).Item(0))
                    For Ligne As Integer = 0 To OledatableSchema.Rows.Count - 1
                        '-------------------------------------------------------------------------------------
                        Select Case OledatableSchema.Rows(Ligne).Item("Cols") '
                            Case "LOT_CONTROL"
                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString = "AR_SuiviStock" Then
                                    PositionLeft = OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(1).ToString.Split(")")(0).Split(".")(0)
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString) = "5" Then
                                        Information &= Strings.LSet("1", PositionLeft)
                                    Else
                                        Information &= Strings.LSet("0", PositionLeft)
                                    End If
                                    Continue For '
                                ElseIf OledatableSchema.Rows(Ligne).Item("InfosLibre").ToString = "True" Then
                                    PositionLeft = OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(1).ToString.Split(")")(0).Split(".")(0)
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString) = "5" Then
                                        Information &= Strings.LSet("1", PositionLeft)
                                    Else
                                        Information &= Strings.LSet("0", PositionLeft)
                                    End If
                                    Continue For '
                                End If
                            Case "SERIAL_NO_CONTROL"
                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString = "AR_SuiviStock" Then
                                    PositionLeft = OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(1).ToString.Split(")")(0).Split(".")(0)
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString) = "1" Then
                                        Information &= Strings.LSet("1", PositionLeft)
                                    Else
                                        Information &= Strings.LSet("0", PositionLeft)
                                    End If
                                    Continue For
                                Else
                                    'si infos libre correspondant
                                    PositionLeft = OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(1).ToString.Split(")")(0).Split(".")(0)
                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString) = "1" Then
                                        Information &= Strings.LSet("1", PositionLeft)
                                    Else
                                        Information &= Strings.LSet("0", PositionLeft)
                                    End If
                                    Continue For '
                                End If
                            Case "SUPPLIER_CODE"
                                PositionLeft = OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(1).ToString.Split(")")(0).Split(".")(0)
                                OleAdaptaterschemaFourssAR = New OleDbDataAdapter("select CT_NUM FROM F_ARTFOURNISS  WHERE F_ARTFOURNISS.AR_Ref='" & OledatableSchemaSage.Rows(i).Item("AR_REF") & "' AND F_ARTFOURNISS.AF_PRINCIPAL =1", OleExcelConnect)
                                OleSchemaDatasetFourssAR = New DataSet
                                OleAdaptaterschemaFourssAR.Fill(OleSchemaDatasetFourssAR)
                                OledatableSchemaFourssAR = OleSchemaDatasetFourssAR.Tables(0)
                                If OledatableSchemaFourssAR.Rows.Count > 0 Then
                                    If OledatableSchemaFourssAR.Rows(0).Item("CT_NUM").ToString = "" Then
                                        Information &= Strings.LSet("", PositionLeft)
                                    Else
                                        Information &= Strings.LSet(OledatableSchemaFourssAR.Rows(0).Item("CT_NUM").ToString, PositionLeft)
                                    End If
                                Else
                                    Information &= Strings.LSet("", PositionLeft)
                                End If
                                Continue For
                        End Select
                        '-------------------------------------******************************------------------------------------------
                        If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                            If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans à la variable colonne
                                If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si ce genre de type de format prevu
                                    Dim ValeurLue As String = OledatableSchemaSage.Rows(i).Item(ColonnName).ToString ' recuperation de la valeur
                                    If ValeurLue = "" Then ' si la valriable ne comtien aucune donne alors on ce rabat vers la valeur pas defaut
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then ' s'il ya une valeur pas defaut
                                            Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                        Else
                                            'si rien n'est prevue pour lavaleur pas defaut
                                        End If
                                    Else
                                        Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                    End If
                                Else
                                    PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                    If OledatableSchema.Rows(Ligne).Item("Cols").ToString = "LINE_COUNT" Then '
                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
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
                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                        If OledatableSchema.Rows(Ligne).Item("Cols").ToString = "LINE_COUNT" Then '
                                            Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                        Else
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        End If
                                    End If
                                Else
                                    'un Mapping Sage existe mais ne possede pas de valeur que retourne la requete et ne possede pas de valeur pas defaut
                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                        Dim ValeurLue = ""
                                        Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                    Else
                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                        If OledatableSchema.Rows(Ligne).Item("Cols").ToString = "LINE_COUNT" Then '
                                            Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                        Else
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Else
                            If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                    Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                    Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                Else
                                    PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                    If OledatableSchema.Rows(Ligne).Item("Cols").ToString = "LINE_COUNT" Then '
                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                        Else
                                            Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                        End If
                                    End If

                                End If
                            Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                'Dim traitementAutres As String = 0
                                If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                    Dim ValeurLue = ""
                                    Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                Else
                                    Dim r = OledatableSchema.Rows(Ligne).Item("Cols").ToString
                                    Dim f = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString
                                    Dim g = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString

                                    PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                    If OledatableSchema.Rows(Ligne).Item("Cols").ToString = "LINE_COUNT" Then '
                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet("", PositionLeft) 'RenvoiValeurLue(0.0, OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                        Else
                                            Information &= Strings.LSet("", PositionLeft) '
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    Error_journal.WriteLine(Information)
                    Information = ""
                    CreationLigne(OledatableSchemaSage.Rows(i).Item("AR_REF"))
                Next
            Else
                'cette requette ne possede pas ligne 
            End If
            Error_journal.Close()
            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try
    End Function
    Public Function GetPosition(ByVal Format As Object) As Integer
        Dim Position() As Object = Format.ToString.Split("(")(1).ToString.Split(")")(0).Split(".")
        If Position.Length = 2 Then
            Return Position(0)
        Else
            Return Position(0)
        End If
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
            'MsgBox(ex.Message)
            Return ""
        End Try
    End Function
    Private Sub BackgroundWorker2_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        'TraitementExtraction()
        EcritureEntete()

    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        PictureBox1.Visible = False
        Me.Alerte.Show()
    End Sub

    Private Sub RadButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadButton2.Click
        Me.Close()
    End Sub

    Private Sub RadButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadButton1.Click
        Me.Hide()
    End Sub
End Class
