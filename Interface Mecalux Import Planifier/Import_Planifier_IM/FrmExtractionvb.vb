Imports System.Data.OleDb
Imports System.IO
Public Class FrmExtraction
    Friend WithEvents DataListeSchema As System.Windows.Forms.DataGridView
    Friend WithEvents Type As System.Windows.Forms.DataGridViewComboBoxColumn
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
    Public CodeArticle As String = ""
    Public ContinuTraitement As Integer = 0
    Public TachePlanifie As String = ""
    Public SocieteCyble As New List(Of String)
    'Dim OledatableSchema As DataTable
    Public Sub AfficheSchemasConso()
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
            If TachePlanifie = "Export Article" Then
                For Each Item As String In SocieteCyble
                    If Item = OledatableSchema.Rows(i).Item("Societe") Then
                        DataListeIntegrer.Rows(i).Cells("choix").Value = True
                    End If
                Next
            Else
                DataListeIntegrer.Rows(i).Cells("choix").Value = True
            End If
        Next i
    End Sub
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
            If TachePlanifie = "Export Article" Then
                For Each Item As String In SocieteCyble
                    If Item = OledatableSchema.Rows(i).Item("Societe") Then
                        DataListeIntegrer.Rows(i).Cells("choix").Value = True
                    End If
                Next
            End If
        Next i
    End Sub
    Public OleAdaptaterschemaVersion As OleDbDataAdapter
    Public OleSchemaDatasetVersion As DataSet
    Public OledatableSchemaVersion As DataTable
    Public Sub FrmExtraction_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
      
        Try
            LirefichierConfig()
            ComboSuivi.SelectedIndex = 2
            Me.WindowState = FormWindowState.Maximized
            TypeSuivi = ComboSuivi.Text
            ChargeInfosLibre()
            If Connected() = True Then
                If TachePlanifie <> "" Then
                    AfficheSchemasTachePlanifie()
                Else
                    BackgroundWorker4.RunWorkerAsync()
                End If
            End If
            OleAdaptaterschemaVersion = New OleDbDataAdapter("select version from P_TABLECORRESP WHERE CodeTbls='PRO'", OleConnenectionArticle)
            OleSchemaDatasetVersion = New DataSet
            OleAdaptaterschemaVersion.Fill(OleSchemaDatasetVersion)
            OledatableSchemaVersion = OleSchemaDatasetVersion.Tables(0)
        Catch ex As Exception

        End Try
    End Sub
    Dim nbrelignes As Integer = 0
    Public Shared ips As Integer
    Public Etat As Boolean = False
    Public Sub BtnModif_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnModif.Click
        nbrelignes = DataListeIntegrer.Rows.Count
        ListBox.Items.Clear()
        ContinuTraitement = 0
        lblsmss.Text = "0/0"
        'Etat = False
        selectindexe = Nothing
        selectids = Nothing
        Try
            For i As Integer = 0 To DataListeIntegrer.Rows.Count - 1
                If DataListeIntegrer.Rows(i).Cells("Choix").Value = True Then
                    selectindexe &= DataListeIntegrer.Rows(i).Index & ";"
                    Etat = True
                End If
                DataListeIntegrer.Rows(i).Cells("Status").Value = My.Resources.btFermer22
            Next
            If Etat = True Then
                lblsms.Visible = False
                selectindexe = selectindexe.Substring(0, selectindexe.Length - 1)
                selectids = selectindexe.Split(";")
                VerificateurChampObligatoire(True, False)
                If ContinuTraitement = 7 Then
                    EstInfosLibre(True, False, True)
                    If NbInfosLibre = NbInfosLibreVue Then 'NbInfosLibreVue
                        Try
                            lblsms.Visible = False
                            EstVide = False
                            'PictureBox1.Visible = True
                            ListeChampSageEntete = RecuperationColonneSage(True, False)
                            ListeChampSageLigne = RecuperationColonneSage(False, True)
                            lblSne.Text = "Scénario Extraction des article"
                            ListeChampSageEntete = ListeChampSageEntete.Substring(0, ListeChampSageEntete.Length - 1)
                            ListeChampSageLigne = ListeChampSageLigne.Substring(0, ListeChampSageLigne.Length - 1)
                              BackgroundWorker1.RunWorkerAsync()
                        Catch ex As Exception
                            'MsgBox(ex.Message)
                        End Try
                    End If
                End If
            Else
                lblsms.Text = "Aucune Société Selectionnée !"
                lblsms.Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Sub VerificateurChampObligatoire(ByVal Entete As Boolean, ByVal Ligne As Boolean)
        OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='PRO' AND Entete=" & Entete & " AND Ligne=" & Ligne & "  ORDER BY ORDRE", OleConnenectionArticle)
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
            OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='PRO' AND (Entete=" & Entete & " OR InfosLibre=true) AND Ligne=False  ORDER BY ORDRE", OleConnenectionArticle)
        Else
            OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='PRO' AND (Ligne=" & Ligne & " OR InfosLibre=true) AND Entete=False   ORDER BY ORDRE", OleConnenectionArticle)
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
            OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='PRO' AND Entete=" & Entete & " AND InfosLibre=" & InfosLibre & " AND Ligne=" & Ligne & "  ORDER BY ORDRE", OleConnenectionArticle)
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
    Public Sub ChargeInfosLibre()
        Try
            ComboInfosLibre.Items.Clear()
            If ExcelConnectS(NomServersql, NomBaseCpta, Nom_Utilsql, Mot_Passql) Then
                OleAdaptaterschemaSage = New OleDbDataAdapter("select CB_Name from cbSysLibre WHERE CB_File='F_ARTICLE'", OleExcelConnectedArticle)
                OleSchemaDatasetSage = New DataSet
                OleAdaptaterschemaSage.Fill(OleSchemaDatasetSage)
                OledatableSchemaSage = OleSchemaDatasetSage.Tables(0)
                If OledatableSchemaSage.Rows.Count <> 0 Then
                    For i As Integer = 0 To OledatableSchemaSage.Rows.Count - 1
                        ComboInfosLibre.Items.Add(OledatableSchemaSage.Rows(i).Item("CB_Name"))
                    Next
                End If
            End If
        Catch ex As Exception
        End Try
    End Sub
    Public Function ExisteInfosLibre(ByVal InfosLibre As String) As Boolean
        Try
            Dim OleAdaptaterschemaSageInfosLibre As OleDbDataAdapter = Nothing
            Dim OleSchemaDatasetSageInfosLibre As DataSet
            Dim OledatableSchemaSageInfosLibre As DataTable
            'For i As Integer = 0 To DataListeIntegrer.Rows.Count - 1
            '    If DataListeIntegrer.Rows(i).Cells("Choix").Value = True Then

            '    End If
            'Next
            If ExcelConnectS(NomServersql, NomBaseCpta, Nom_Utilsql, Mot_Passql) Then
                OleAdaptaterschemaSageInfosLibre = New OleDbDataAdapter("select * from cbSysLibre WHERE CB_File='F_ARTICLE' And CB_Name='" & Join(Split(Trim(InfosLibre), "'"), "''") & "'", OleExcelConnectedArticle)
                OleSchemaDatasetSageInfosLibre = New DataSet
                OleAdaptaterschemaSageInfosLibre.Fill(OleSchemaDatasetSageInfosLibre)
                OledatableSchemaSageInfosLibre = OleSchemaDatasetSageInfosLibre.Tables(0)
                If OledatableSchemaSageInfosLibre.Rows.Count <> 0 Then
                    NbInfosLibreVue += 1
                    Return True
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
        End Try
    End Function
    Public Sub videElement()
        ListeExporte = Nothing
        NbInfosLibre = Nothing
        NbInfosLibreVue = Nothing
        OleAdaptaterschemaSage = Nothing
        OleAdaptaterschemaSageDetails = Nothing
        OleAdaptaterschema = Nothing
        OleAdaptaterschemaLigne = Nothing
        OleAdaptaterschemaFourssAR = Nothing

        OleSchemaDatasetSage = Nothing
        OleSchemaDatasetSageDetails = Nothing
        OleSchemaDataset = Nothing
        OleSchemaDatasetLigne = Nothing
        OleSchemaDatasetFourssAR = Nothing
        OledatableSchemaSage = Nothing
        OledatableSchemaSageDetails = Nothing
        OledatableSchema = Nothing
        OledatableSchemaLigne = Nothing
        OledatableSchemaFourssAR = Nothing
    End Sub
    Public Societe As String = ""
    Public Sub ExtractionArticle(ByVal SelectChamp As String, ByVal indice As Integer)
        Dim Information As String = ""
        Dim ExistColonne As Boolean = False
        Dim ColonnName As String = ""
        Dim PositionLeft As Object = Nothing
        Try
            videElement()
            If DataListeIntegrer.Rows(indice).Cells("Choix").Value = True Then
                If ExcelConnectS(DataListeIntegrer.Rows(indice).Cells("Serveur1").Value, DataListeIntegrer.Rows(indice).Cells("Societe1").Value, DataListeIntegrer.Rows(indice).Cells("NomUtil").Value, DataListeIntegrer.Rows(indice).Cells("Mot").Value) Then
                    Societe = DataListeIntegrer.Rows(indice).Cells("Societe1").Value
                    If CheckSuivi.Checked And CheckSommeil.Checked Then
                        Select Case TypeSuivi
                            Case "Aucun"
                                If RbtSommeilON.Checked Then
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=1  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=1 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=1  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=1 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                                ElseIf RbtSommeilOFF.Checked Then
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=0 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=0 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                                Else
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock<>0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock<>0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock<>0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock<>0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                                End If
                            Case "Sérialisé"
                                If RbtSommeilON.Checked Then
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=1 AND AR_Sommeil=1  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=1 AND AR_Sommeil=1 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=1 AND AR_Sommeil=1  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=1 AND AR_Sommeil=1 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                                ElseIf RbtSommeilOFF.Checked Then
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=1 AND AR_Sommeil=0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=1 AND AR_Sommeil=0 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=1 AND AR_Sommeil=0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=1 AND AR_Sommeil=0 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                                Else
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock<>0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock<>0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock<>0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock<>0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                                End If
                            Case "CMUP"
                                If RbtSommeilON.Checked Then
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=2 AND AR_Sommeil=1  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=2 AND AR_Sommeil=1 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=2 AND AR_Sommeil=1  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=2 AND AR_Sommeil=1 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                                ElseIf RbtSommeilOFF.Checked Then
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=2 AND AR_Sommeil=0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=2 AND AR_Sommeil=0 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=2 AND AR_Sommeil=0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=2 AND AR_Sommeil=0 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                                Else
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock<>0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock<>0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock<>0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock<>0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                                End If
                            Case "FIFO"
                                If RbtSommeilON.Checked Then
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=3 AND AR_Sommeil=1  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=3 AND AR_Sommeil=1 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=3 AND AR_Sommeil=1  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=3 AND AR_Sommeil=1 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                                ElseIf RbtSommeilOFF.Checked Then
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=3 AND AR_Sommeil=0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=3 AND AR_Sommeil=0 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=3 AND AR_Sommeil=0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=3 AND AR_Sommeil=0 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                                Else
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock<>0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock<>0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock<>0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock<>0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                                End If
                            Case "LIFO"
                                If RbtSommeilON.Checked Then
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=4 AND AR_Sommeil=1  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=4 AND AR_Sommeil=1 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=4 AND AR_Sommeil=1  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=4 AND AR_Sommeil=1 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                                ElseIf RbtSommeilOFF.Checked Then
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=4 AND AR_Sommeil=0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=4 AND AR_Sommeil=0 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=4 AND AR_Sommeil=0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=4 AND AR_Sommeil=0 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                                Else
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock<>0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock<>0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock<>0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock<>0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                                End If
                            Case "Par Lot"
                                If RbtSommeilON.Checked Then
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=5 AND AR_Sommeil=1  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=5 AND AR_Sommeil=1 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=5 AND AR_Sommeil=1  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=5 AND AR_Sommeil=1 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                                ElseIf RbtSommeilOFF.Checked Then
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=5 AND AR_Sommeil=0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=5 AND AR_Sommeil=0 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=5 AND AR_Sommeil=0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=5 AND AR_Sommeil=0 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                                Else
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock<>0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock<>0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle)
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock<>0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock<>0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                                End If
                        End Select
                    ElseIf CheckSuivi.Checked And CheckInfosLibre.Checked Then
                        Select Case TypeSuivi
                            Case "Aucun"
                                If InfosLibre <> "" Then
                                    If txtinfosLibre.Text <> "" Then
                                        If IsNumeric(txtinfosLibre.Text) Then
                                            If Ckmodifier.Checked Then
                                                If SelectChamp IsNot Nothing Then
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND " & InfosLibre & "=" & txtinfosLibre.Text & "  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                Else
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND " & InfosLibre & "=" & txtinfosLibre.Text & "  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                End If
                                            Else
                                                If SelectChamp IsNot Nothing Then
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND " & InfosLibre & "=" & txtinfosLibre.Text & "  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                Else
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND " & InfosLibre & "=" & txtinfosLibre.Text & "  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                End If
                                            End If
                                        Else
                                            If Ckmodifier.Checked Then
                                                If SelectChamp IsNot Nothing Then
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                Else
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                End If
                                            Else
                                                If SelectChamp IsNot Nothing Then
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                Else
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                End If
                                            End If
                                        End If
                                    Else
                                        If Ckmodifier.Checked Then
                                            If SelectChamp IsNot Nothing Then
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            Else
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            End If
                                        Else
                                            If SelectChamp IsNot Nothing Then
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            Else
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            End If
                                        End If
                                    End If
                                Else
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                                End If
                            Case "Sérialisé"
                                If InfosLibre <> "" Then
                                    If txtinfosLibre.Text <> "" Then
                                        If IsNumeric(txtinfosLibre.Text) Then
                                            If Ckmodifier.Checked Then
                                                If SelectChamp IsNot Nothing Then
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=1 AND " & InfosLibre & "=" & txtinfosLibre.Text & "  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                Else
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=1 AND " & InfosLibre & "=" & txtinfosLibre.Text & "  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                End If
                                            Else
                                                If SelectChamp IsNot Nothing Then
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=1 AND " & InfosLibre & "=" & txtinfosLibre.Text & "  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                Else
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=1 AND " & InfosLibre & "=" & txtinfosLibre.Text & "  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                End If
                                            End If
                                        Else
                                            If Ckmodifier.Checked Then
                                                If SelectChamp IsNot Nothing Then
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=1 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                Else
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=1 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                End If
                                            Else
                                                If SelectChamp IsNot Nothing Then
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=1 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                Else
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=1 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                End If
                                            End If
                                        End If
                                    Else
                                        If Ckmodifier.Checked Then
                                            If SelectChamp IsNot Nothing Then
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=1 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            Else
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=1 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            End If
                                        Else
                                            If SelectChamp IsNot Nothing Then
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=1 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            Else
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=1 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            End If
                                        End If
                                    End If
                                Else
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=1 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=1 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=1 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=1 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                                End If
                            Case "CMUP"
                                If InfosLibre <> "" Then
                                    If txtinfosLibre.Text <> "" Then
                                        If IsNumeric(txtinfosLibre.Text) Then
                                            If Ckmodifier.Checked Then
                                                If SelectChamp IsNot Nothing Then
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=2 AND " & InfosLibre & "=" & txtinfosLibre.Text & "  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                Else
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=2 AND " & InfosLibre & "=" & txtinfosLibre.Text & "  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                End If
                                            Else
                                                If SelectChamp IsNot Nothing Then
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=2 AND " & InfosLibre & "=" & txtinfosLibre.Text & "  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                Else
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=2 AND " & InfosLibre & "=" & txtinfosLibre.Text & "  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                End If
                                            End If
                                        Else
                                            If Ckmodifier.Checked Then
                                                If SelectChamp IsNot Nothing Then
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=2 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                Else
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=2 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                End If
                                            Else
                                                If SelectChamp IsNot Nothing Then
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=2 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                Else
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=2 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                End If
                                            End If
                                        End If
                                    Else
                                        If Ckmodifier.Checked Then
                                            If SelectChamp IsNot Nothing Then
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=2 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            Else
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=2 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            End If
                                        Else
                                            If SelectChamp IsNot Nothing Then
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=2 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            Else
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=2 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            End If
                                        End If
                                    End If
                                Else
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=2 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=2 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=2 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=2 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                                End If
                            Case "FIFO"
                                If InfosLibre <> "" Then
                                    If txtinfosLibre.Text <> "" Then
                                        If IsNumeric(txtinfosLibre.Text) Then
                                            If Ckmodifier.Checked Then
                                                If SelectChamp IsNot Nothing Then
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=3 AND " & InfosLibre & "=" & txtinfosLibre.Text & "  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                Else
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=3 AND " & InfosLibre & "=" & txtinfosLibre.Text & "  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                End If
                                            Else
                                                If SelectChamp IsNot Nothing Then
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=3 AND " & InfosLibre & "=" & txtinfosLibre.Text & "  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                Else
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=3 AND " & InfosLibre & "=" & txtinfosLibre.Text & "  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                End If
                                            End If
                                        Else
                                            If Ckmodifier.Checked Then
                                                If SelectChamp IsNot Nothing Then
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=3 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                Else
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=3 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                End If
                                            Else
                                                If SelectChamp IsNot Nothing Then
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=3 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                Else
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=3 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                End If
                                            End If
                                        End If
                                    Else
                                        If Ckmodifier.Checked Then
                                            If SelectChamp IsNot Nothing Then
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=3 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            Else
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=3 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            End If
                                        Else
                                            If SelectChamp IsNot Nothing Then
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=3 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            Else
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=3 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            End If
                                        End If
                                    End If
                                Else
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=3 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=3 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=3 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=3 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                                End If
                            Case "LIFO"
                                If InfosLibre <> "" Then
                                    If txtinfosLibre.Text <> "" Then
                                        If IsNumeric(txtinfosLibre.Text) Then
                                            If Ckmodifier.Checked Then
                                                If SelectChamp IsNot Nothing Then
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=4 AND " & InfosLibre & "=" & txtinfosLibre.Text & "  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                Else
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=4 AND " & InfosLibre & "=" & txtinfosLibre.Text & "  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                End If
                                            Else
                                                If SelectChamp IsNot Nothing Then
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=4 AND " & InfosLibre & "=" & txtinfosLibre.Text & "  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                Else
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=4 AND " & InfosLibre & "=" & txtinfosLibre.Text & "  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                End If
                                            End If
                                        Else
                                            If Ckmodifier.Checked Then
                                                If SelectChamp IsNot Nothing Then
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=4 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                Else
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=4 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                End If
                                            Else
                                                If SelectChamp IsNot Nothing Then
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=4 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                Else
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=4 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                End If
                                            End If
                                        End If
                                    Else
                                        If Ckmodifier.Checked Then
                                            If SelectChamp IsNot Nothing Then
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=4 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            Else
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=4 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            End If
                                        Else
                                            If SelectChamp IsNot Nothing Then
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=4 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            Else
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=4 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            End If
                                        End If
                                    End If
                                Else
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=4 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=4 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=4 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=4 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                                End If
                            Case "Par Lot"
                                If InfosLibre <> "" Then
                                    If txtinfosLibre.Text <> "" Then
                                        If IsNumeric(txtinfosLibre.Text) Then
                                            If Ckmodifier.Checked Then
                                                If SelectChamp IsNot Nothing Then
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=5 AND " & InfosLibre & "=" & txtinfosLibre.Text & "  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                Else
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=5 AND " & InfosLibre & "=" & txtinfosLibre.Text & "  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                End If
                                            Else
                                                If SelectChamp IsNot Nothing Then
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=5 AND " & InfosLibre & "=" & txtinfosLibre.Text & "  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                Else
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=5 AND " & InfosLibre & "=" & txtinfosLibre.Text & "  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                End If
                                            End If
                                        Else
                                            If Ckmodifier.Checked Then
                                                If SelectChamp IsNot Nothing Then
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=5 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                Else
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=5 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                End If
                                            Else
                                                If SelectChamp IsNot Nothing Then
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=5 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                Else
                                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=5 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                                End If
                                            End If
                                        End If
                                    Else
                                        If Ckmodifier.Checked Then
                                            If SelectChamp IsNot Nothing Then
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=5 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            Else
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=5 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            End If
                                        Else
                                            If SelectChamp IsNot Nothing Then
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=5 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            Else
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=5 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            End If
                                        End If
                                    End If
                                Else
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=5 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=5 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=5 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=5 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                                End If
                        End Select
                    ElseIf CheckSommeil.Checked And CheckInfosLibre.Checked Then
                        If RbtSommeilON.Checked Then
                            If InfosLibre <> "" Then
                                If txtinfosLibre.Text <> "" Then
                                    If IsNumeric(txtinfosLibre.Text) Then
                                        If Ckmodifier.Checked Then
                                            If SelectChamp IsNot Nothing Then
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=1 AND " & InfosLibre & "=" & txtinfosLibre.Text & "  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            Else
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=1 AND " & InfosLibre & "=" & txtinfosLibre.Text & "   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            End If
                                        Else
                                            If SelectChamp IsNot Nothing Then
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=1 AND " & InfosLibre & "=" & txtinfosLibre.Text & "  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            Else
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=1 AND " & InfosLibre & "=" & txtinfosLibre.Text & "   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            End If
                                        End If
                                    Else
                                        If Ckmodifier.Checked Then
                                            If SelectChamp IsNot Nothing Then
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=1 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            Else
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=1 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            End If
                                        Else
                                            If SelectChamp IsNot Nothing Then
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=1 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            Else
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=1 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            End If
                                        End If
                                    End If
                                Else
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=1 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=1 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=1 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=1 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                                End If
                            Else
                                If Ckmodifier.Checked Then
                                    If SelectChamp IsNot Nothing Then
                                        OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=1 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                    Else
                                        OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=1  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                    End If
                                Else
                                    If SelectChamp IsNot Nothing Then
                                        OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=1 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                    Else
                                        OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=1  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                    End If
                                End If
                            End If
                        ElseIf RbtSommeilOFF.Checked Then
                            If InfosLibre <> "" Then
                                If txtinfosLibre.Text <> "" Then
                                    If IsNumeric(txtinfosLibre.Text) Then
                                        If Ckmodifier.Checked Then
                                            If SelectChamp IsNot Nothing Then
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=0 AND " & InfosLibre & "=" & txtinfosLibre.Text & "  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            Else
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=0 AND " & InfosLibre & "=" & txtinfosLibre.Text & "   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            End If
                                        Else
                                            If SelectChamp IsNot Nothing Then
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=0 AND " & InfosLibre & "=" & txtinfosLibre.Text & "  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            Else
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=0 AND " & InfosLibre & "=" & txtinfosLibre.Text & "   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            End If
                                        End If
                                    Else
                                        If Ckmodifier.Checked Then
                                            If SelectChamp IsNot Nothing Then
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=0 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            Else
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=0 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            End If
                                        Else
                                            If SelectChamp IsNot Nothing Then
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=0 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            Else
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=0 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            End If
                                        End If
                                    End If
                                Else
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=0 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=0 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=0 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=0 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                                End If
                            Else
                                If Ckmodifier.Checked Then
                                    If SelectChamp IsNot Nothing Then
                                        OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=0 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                    Else
                                        OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                    End If
                                Else
                                    If SelectChamp IsNot Nothing Then
                                        OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=0 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                    Else
                                        OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                    End If
                                End If
                            End If
                        Else
                            If InfosLibre <> "" Then
                                If txtinfosLibre.Text <> "" Then
                                    If IsNumeric(txtinfosLibre.Text) Then
                                        If Ckmodifier.Checked Then
                                            If SelectChamp IsNot Nothing Then
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND " & InfosLibre & "=" & txtinfosLibre.Text & "  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            Else
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0  AND " & InfosLibre & "=" & txtinfosLibre.Text & "   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            End If
                                        Else
                                            If SelectChamp IsNot Nothing Then
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND " & InfosLibre & "=" & txtinfosLibre.Text & "  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            Else
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0  AND " & InfosLibre & "=" & txtinfosLibre.Text & "   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            End If
                                        End If
                                    Else
                                        If Ckmodifier.Checked Then
                                            If SelectChamp IsNot Nothing Then
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0  AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            Else
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            End If
                                        Else
                                            If SelectChamp IsNot Nothing Then
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0  AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            Else
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            End If
                                        End If
                                    End If
                                Else
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0  AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0  AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                                End If
                            Else
                                If Ckmodifier.Checked Then
                                    If SelectChamp IsNot Nothing Then
                                        OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                    Else
                                        OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                    End If
                                Else
                                    If SelectChamp IsNot Nothing Then
                                        OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                    Else
                                        OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                    End If
                                End If

                            End If
                        End If
                    ElseIf CheckSuivi.Checked And CheckSommeil.Checked And CheckInfosLibre.Checked Then
                        'trois cas reunie
                    Else
                        If CheckSuivi.Checked Then
                            Select Case TypeSuivi
                                Case "Aucun"
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                                Case "Sérialisé"
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=1 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=1 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=1 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=1 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                                Case "CMUP"
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=2 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=2 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=2 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=2 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                                Case "FIFO"
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=3 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=3 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=3 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=3 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                                Case "LIFO"
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=4 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=4 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=4 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=4 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                                Case "Par Lot"
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=5 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=5 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=5 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=5 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                            End Select
                        ElseIf CheckSommeil.Checked Then
                            If RbtSommeilON.Checked Then
                                If Ckmodifier.Checked Then
                                    If SelectChamp IsNot Nothing Then
                                        OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE  AR_Sommeil=1 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                    Else
                                        OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_Sommeil=1 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                    End If
                                Else
                                    If SelectChamp IsNot Nothing Then
                                        OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE  AR_Sommeil=1 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                    Else
                                        OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_Sommeil=1 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                    End If
                                End If
                            ElseIf RbtSommeilOFF.Checked Then
                                If Ckmodifier.Checked Then
                                    If SelectChamp IsNot Nothing Then
                                        OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE  AR_Sommeil=0 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                    Else
                                        OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_Sommeil=0 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                    End If
                                Else
                                    If SelectChamp IsNot Nothing Then
                                        OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE  AR_Sommeil=0 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                    Else
                                        OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_Sommeil=0 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                    End If
                                End If

                            Else
                                Exit Sub
                                'If SelectChamp IsNot Nothing Then
                                '    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                'Else
                                '    OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                'End If
                            End If
                        ElseIf CheckInfosLibre.Checked Then
                            If InfosLibre <> "" Then
                                If txtinfosLibre.Text <> "" Then
                                    If IsNumeric(txtinfosLibre.Text) Then
                                        If Ckmodifier.Checked Then
                                            If SelectChamp IsNot Nothing Then
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=0 AND " & InfosLibre & "=" & txtinfosLibre.Text & "  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            Else
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=0 AND " & InfosLibre & "=" & txtinfosLibre.Text & "   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            End If
                                        Else
                                            If SelectChamp IsNot Nothing Then
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=0 AND " & InfosLibre & "=" & txtinfosLibre.Text & "  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            Else
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=0 AND " & InfosLibre & "=" & txtinfosLibre.Text & "   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            End If
                                        End If
                                    Else
                                        If Ckmodifier.Checked Then
                                            If SelectChamp IsNot Nothing Then
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=0 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            Else
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=0 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            End If
                                        Else
                                            If SelectChamp IsNot Nothing Then
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=0 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            Else
                                                OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock=0 AND AR_Sommeil=0 AND " & InfosLibre & "='" & txtinfosLibre.Text & "'   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                            End If
                                        End If
                                    End If
                                Else
                                    If Ckmodifier.Checked Then
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE  AR_Sommeil=0 AND " & InfosLibre & "='" & txtinfosLibre.Text & "' ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE  AR_Sommeil=0 AND " & InfosLibre & "='" & txtinfosLibre.Text & "' ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    Else
                                        If SelectChamp IsNot Nothing Then
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE  AR_Sommeil=0 AND " & InfosLibre & "='" & txtinfosLibre.Text & "' ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        Else
                                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE  AR_Sommeil=0 AND " & InfosLibre & "='" & txtinfosLibre.Text & "' ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                        End If
                                    End If
                                End If
                            Else
                                If Ckmodifier.Checked Then
                                    If SelectChamp IsNot Nothing Then
                                        OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE  AR_Sommeil=0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                    Else
                                        OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_Sommeil=0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                    End If
                                Else
                                    If SelectChamp IsNot Nothing Then
                                        OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE  AR_Sommeil=0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                    Else
                                        OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_Sommeil=0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif
                                    End If
                                End If
                            End If
                        Else
                            If Ckmodifier.Checked Then
                                If SelectChamp IsNot Nothing Then
                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock<>0   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif AND AR_DateCreation <> AR_DateModif
                                Else
                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock<>0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif AND AR_DateCreation <> AR_DateModif
                                End If
                            Else
                                If SelectChamp IsNot Nothing Then
                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif FROM F_ARTICLE WHERE AR_SuiviStock<>0   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif AND AR_DateCreation <> AR_DateModif
                                Else
                                    OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock<>0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedArticle) 'AND AR_DateCreation = AR_DateModif AND AR_DateCreation <> AR_DateModif
                                End If
                            End If
                        End If
                    End If
                    OleSchemaDatasetSage = New DataSet
                    OleAdaptaterschemaSage.Fill(OleSchemaDatasetSage)
                    OledatableSchemaSage = OleSchemaDatasetSage.Tables(0)

                    OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='PRO' AND Entete=True   ORDER BY ORDRE", OleConnenectionArticle)
                    OleSchemaDataset = New DataSet
                    OleAdaptaterschema.Fill(OleSchemaDataset)
                    OledatableSchema = OleSchemaDataset.Tables(0)
                End If
            End If
        Catch ex As Exception
            If Not TachePlanifie <> "" Then
                Throw New Exception(ex.Message)
            End If
        End Try
    End Sub
    Public TypedeTraitement As String = ""
    Public Function RDernierDate(ByRef IDTraitement As Object) As Date
        Try
            Dim OleAdaptater As OleDbDataAdapter
            Dim OleAfficheDataset As DataSet
            Dim Oledatable As DataTable
            OleAdaptater = New OleDbDataAdapter("select * from WEAT_ADATE WHERE IDDossier=" & IDTraitement & " And TypeExport='Export Article'", OleConnenectionArticle)
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
    Private Function ArticleRecemmentModifier(ByRef DatDerntraitement As Date, ByRef OleArtModifierDt As DataTable, ByRef m As Integer) As Boolean
        Dim MustModified As Boolean = False
        Try
            If DatDerntraitement <= OleArtModifierDt.Rows(m).Item("AR_DateModif") Then
                MustModified = True
            End If
        Catch ex As Exception
        End Try
        Return MustModified
    End Function
    Public Function IDernierDate(ByRef IDTraitement As Object) As Date
        Try
            Dim Insertion As String
            Dim OleCmdIns As OleDbCommand
            Dim OleAdaptater As OleDbDataAdapter
            Dim OleAfficheDataset As DataSet
            Dim Oledatable As DataTable
            OleAdaptater = New OleDbDataAdapter("select * from WEAT_ADATE WHERE IDDossier=" & IDTraitement & " And TypeExport='Export Article'", OleConnenectionArticle)
            OleAfficheDataset = New DataSet
            OleAdaptater.Fill(OleAfficheDataset)
            Oledatable = OleAfficheDataset.Tables(0)
            If Oledatable.Rows.Count <> 0 Then
                Insertion = " UPDATE WEAT_ADATE SET DateDern='" & Strings.FormatDateTime(Now, Microsoft.VisualBasic.DateFormat.ShortDate) & "' WHERE IDDossier=" & IDTraitement & " And TypeExport='Export Article'"
                OleCmdIns = New OleDbCommand(Insertion)
                OleCmdIns.Connection = OleConnenectionArticle
                OleCmdIns.ExecuteNonQuery()
            Else
                Insertion = " Insert Into WEAT_ADATE (TypeExport,IDDossier,DateDern) Values ('Export Article'," & IDTraitement & ",'" & Strings.FormatDateTime(Now, Microsoft.VisualBasic.DateFormat.ShortDate) & "')"
                OleCmdIns = New OleDbCommand(Insertion)
                OleCmdIns.Connection = OleConnenectionArticle
                OleCmdIns.ExecuteNonQuery()
            End If
        Catch ex As Exception
        End Try
    End Function
    Public Sub CreationLigne(ByVal valeurLue As Object)
        Dim PositionLeftLigne As Integer = 0
        Dim Information As String = ""
        Dim ColonnNameLigne As String = ""
        Dim P_CONDITIONNEMENT As String = ""

        Dim oleAdatarSchemaP_CON As New OleDbDataAdapter("SELECT P_CONDITIONNEMENT FROM F_ARTICLE INNER JOIN P_CONDITIONNEMENT ON F_ARTICLE.AR_CONDITION = P_CONDITIONNEMENT.cbindice AND F_ARTICLE.AR_Ref='" & valeurLue & "'", OleExcelConnectedArticle)
        Dim oleShemanDataSetP_CON As New DataSet
        Dim oleSchemaTable As DataTable
        oleAdatarSchemaP_CON.Fill(oleShemanDataSetP_CON)
        oleSchemaTable = oleShemanDataSetP_CON.Tables(0)

        If ChBoxEC_Qté.Checked Then
            OleAdaptaterschemaSageDetails = New OleDbDataAdapter("select * FROM F_CONDITION,F_ARTICLE WHERE F_CONDITION.AR_REF=F_ARTICLE.AR_REF AND F_CONDITION.EC_Quantite<>1 AND F_CONDITION.CO_Principal =1 AND F_CONDITION.AR_REF='" & valeurLue & "'", OleExcelConnectedArticle)
        Else
            OleAdaptaterschemaSageDetails = New OleDbDataAdapter("select * FROM F_CONDITION,F_ARTICLE WHERE F_CONDITION.AR_REF=F_ARTICLE.AR_REF AND F_CONDITION.CO_Principal =1 AND F_CONDITION.AR_REF='" & valeurLue & "'", OleExcelConnectedArticle)
        End If
        OleSchemaDatasetSageDetails = New DataSet
        OleAdaptaterschemaSageDetails.Fill(OleSchemaDatasetSageDetails)
        OledatableSchemaSageDetails = OleSchemaDatasetSageDetails.Tables(0)

        OleAdaptaterschemaLigne = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='PRO' AND Ligne=True   ORDER BY ORDRE", OleConnenectionArticle)
        OleSchemaDatasetLigne = New DataSet
        OleAdaptaterschemaLigne.Fill(OleSchemaDatasetLigne)
        OledatableSchemaLigne = OleSchemaDatasetLigne.Tables(0)
        Try
            If OledatableSchemaSageDetails.Rows.Count <> 0 Then
                For i As Integer = 0 To OledatableSchemaSageDetails.Rows.Count - 1
                    Element2 = OledatableSchemaSageDetails.Rows(i).Item("CO_NO")
                    If ListBox.InvokeRequired Then
                        Dim MonDelegate As New Evenement(AddressOf Traitement2)
                        ListBox.Invoke(MonDelegate)
                    Else
                        Traitement2()
                    End If
                    For Ligne As Integer = 0 To OledatableSchemaLigne.Rows.Count - 1
                        If Ligne = 0 Then
                            Information &= Strings.LSet("/", 1)
                        ElseIf Ligne = 1 Then
                            Dim TypeDef = EstCreerOuMoudiffier(valeurLue)
                            Information &= Strings.LSet(TypeDef, 1)
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
                                            P_CONDITIONNEMENT = oleSchemaTable.Rows(0).Item("P_CONDITIONNEMENT").ToString
                                            Information &= Strings.LSet(P_CONDITIONNEMENT, PositionLeftLigne)
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
                                P_CONDITIONNEMENT = oleSchemaTable.Rows(0).Item("P_CONDITIONNEMENT").ToString
                                If OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                    ColonnNameLigne = OledatableSchemaLigne.Rows(Ligne).Item("ChampSage").ToString
                                    If OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString <> "" Then
                                        If OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSageDetails.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                        Else
                                            Information &= Strings.LSet(P_CONDITIONNEMENT, PositionLeftLigne)
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
                            Case "DESC_CONTAINERTYPE_CODE"
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
                    Error_journalArt.WriteLine(Information)
                    Information = ""
                Next
                Try
                    If FlagtamponArticle <> "" Then
                        Dim OleCommande As New OleDbCommand("UPDATE F_ARTICLE SET " & FlagtamponArticle & "='" & Format(DateTime.Now, "yyyyMMddhhmm") & "' WHERE AR_REF='" & valeurLue & "'", OleExcelConnectedArticle)
                        OleCommande.ExecuteNonQuery()
                    End If
                Catch ex As Exception
                End Try
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Dim ifRowError = 0
    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            ExtractionArticle(ListeChampSageEntete, selectids(ips))
        Catch ex As Exception
            If Not TachePlanifie <> "" Then
                ifRowError = selectids(ips)
                DataListeIntegrer.Rows(ifRowError).Cells("Status").Value = My.Resources.btSupprimer221
                ErreurJrnArt = File.AppendText(Pathsfilejournal & "ERREURPRO09" & Strings.Right(DateAndTime.Year(Now), 4) & Format(DateAndTime.Month(Now), "00") & Format(DateAndTime.Day(Now), "00") & Format(DateAndTime.Hour(Now), "00") & Format(DateAndTime.Minute(Now), "00") & Format(DateAndTime.Second(Now), "00") & ".txt")
                ErreurJrnArt.WriteLine("ERREUR SYSTEME : " & ex.Message)
                ErreurJrnArt.Close()
            End If
        End Try
    End Sub
    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Try
            BackgroundWorker2.RunWorkerAsync()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub
    Delegate Sub Evenement()
    Public Element As String
    Public cpteurs As Integer
    Public Element2 As String
    Public Sub Traitement1()
        ListBox.Items.Add("Extraction de l'article : " & Element)
        lblsmss.Text = cpteurs + 1 & "/" & OledatableSchemaSage.Rows.Count
    End Sub
    Public Sub Traitement2()
        ListBox.Items.Add("     0---> Extraction detail de l'article  Code de l'article : " & Element & " Code de la Condition :[" & Element2 & "]")
        ListBox.SelectedIndex = ListBox.Items.Count - 1
    End Sub
    Public Function EcritureEntete() As Boolean
        Dim Information As String = ""
        Dim ExistColonne As Boolean = False
        Dim ColonnName As String = ""
        Dim PositionLeft As Object = Nothing
        Dim statut As Boolean = False
        Dim version As String = "04"
        Try
            If OledatableSchemaSage.Rows.Count <> 0 Then
                Dim DateTraitement As Date = RDernierDate(1)
                version = OledatableSchemaVersion.Rows(0).Item("version")
                For i As Integer = 0 To OledatableSchemaSage.Rows.Count - 1
                    If Ckmodifier.Checked Then
                        If ArticleRecemmentModifier(DateTraitement, OledatableSchemaSage, i) = True Then
                            If Ckmodifier.Checked And statut = False Then
                                Error_journalArt = File.AppendText(PathsFileFormatiers & "PRO" & version & Strings.Right(DateAndTime.Year(Now), 4) & Format(DateAndTime.Month(Now), "00") & Format(DateAndTime.Day(Now), "00") & Format(DateAndTime.Hour(Now), "00") & Format(DateAndTime.Minute(Now), "00") & Format(DateAndTime.Second(Now), "00") & ".txt")
                                statut = True
                            End If
                            Element = OledatableSchemaSage.Rows(i).Item(0)
                            If ListBox.InvokeRequired Then
                                Dim MonDelegate As New Evenement(AddressOf Traitement1)
                                ListBox.Invoke(MonDelegate)
                            Else
                                Traitement1()
                            End If
                            cpteurs = i
                            Try
                                For Ligne As Integer = 0 To OledatableSchema.Rows.Count - 1
                                    '-------------------------------------------------------------------------------------
                                    If OledatableSchema.Rows(Ligne).Item("Cols").ToString = "OWNER_CODE" Then
                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                        If Societe <> "" Then
                                            Information &= Strings.LSet(Societe, PositionLeft)
                                        Else
                                            Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                        End If
                                        'If ExisteInfosLibre(FlagtamponArticle) = False Then
                                        '    ListBox.Items.Add("la zone infos libre {" & FlagtamponArticle & "}  N'existe pas sur les article")
                                        '    Exit Function
                                        'End If
                                        Continue For
                                    End If
                                    Select Case OledatableSchema.Rows(Ligne).Item("Cols") '
                                        Case "LOT_CONTROL"
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString = "AR_SuiviStock" Then
                                                PositionLeft = OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(1).ToString.Split(")")(0).Split(".")(0)
                                                If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString) = "5" Then
                                                    Information &= Strings.LSet("1", PositionLeft)
                                                Else
                                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue") <> Nothing Then
                                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue"), PositionLeft)
                                                    Else
                                                        Information &= Strings.LSet("0", PositionLeft)
                                                    End If
                                                End If
                                                Continue For '
                                            ElseIf OledatableSchema.Rows(Ligne).Item("InfosLibre").ToString = "True" Then
                                                PositionLeft = OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(1).ToString.Split(")")(0).Split(".")(0)
                                                If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString) = "5" Then
                                                    Information &= Strings.LSet("1", PositionLeft)
                                                Else
                                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue") <> Nothing Then
                                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue"), PositionLeft)
                                                    Else
                                                        Information &= Strings.LSet("0", PositionLeft)
                                                    End If
                                                End If
                                                Continue For '
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue"), PositionLeft)
                                                Continue For
                                            End If
                                        Case "SERIAL_NO_CONTROL"
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString = "AR_SuiviStock" Then
                                                PositionLeft = OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(1).ToString.Split(")")(0).Split(".")(0)
                                                If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString) = "1" Then
                                                    Information &= Strings.LSet("1", PositionLeft)
                                                Else
                                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue") <> Nothing Then
                                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue"), PositionLeft)
                                                    Else
                                                        Information &= Strings.LSet("0", PositionLeft)
                                                    End If
                                                End If
                                                Continue For
                                            ElseIf OledatableSchema.Rows(Ligne).Item("InfosLibre").ToString = "True" Then
                                                'si infos libre correspondant
                                                PositionLeft = OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(1).ToString.Split(")")(0).Split(".")(0)
                                                If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString) = "1" Then
                                                    Information &= Strings.LSet("1", PositionLeft)
                                                Else
                                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue") <> Nothing Then
                                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue"), PositionLeft)
                                                    Else
                                                        Information &= Strings.LSet("0", PositionLeft)
                                                    End If
                                                End If
                                                Continue For '
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue"), PositionLeft)
                                                Continue For
                                            End If
                                        Case "SUPPLIER_CODE"
                                            PositionLeft = OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(1).ToString.Split(")")(0).Split(".")(0)
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage") <> Nothing Then
                                                If OledatableSchema.Rows(Ligne).Item("ChampSage") <> "" Then
                                                    OleAdaptaterschemaFourssAR = New OleDbDataAdapter("select CT_NUM FROM F_ARTFOURNISS  WHERE F_ARTFOURNISS.AR_Ref='" & OledatableSchemaSage.Rows(i).Item("AR_REF") & "' AND F_ARTFOURNISS.AF_PRINCIPAL =1", OleExcelConnectedArticle)
                                                    OleSchemaDatasetFourssAR = New DataSet
                                                    OleAdaptaterschemaFourssAR.Fill(OleSchemaDatasetFourssAR)
                                                    OledatableSchemaFourssAR = OleSchemaDatasetFourssAR.Tables(0)
                                                    If OledatableSchemaFourssAR.Rows.Count > 0 Then
                                                        If OledatableSchemaFourssAR.Rows(0).Item("CT_NUM").ToString = "" Then
                                                            Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue"), PositionLeft)
                                                        Else
                                                            Information &= Strings.LSet(OledatableSchemaFourssAR.Rows(0).Item("CT_NUM").ToString, PositionLeft)
                                                        End If
                                                    Else
                                                        Information &= Strings.LSet("", PositionLeft)
                                                    End If
                                                    Continue For
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue"), PositionLeft)
                                                    Continue For
                                                End If
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue"), PositionLeft)
                                                Continue For
                                            End If
                                        Case "UOM_BASE_CODE"
                                            PositionLeft = OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(1).ToString.Split(")")(0).Split(".")(0)
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString = "AR_UNITEVEN" Then
                                                    Dim oleDataAdapeterLibelle As OleDbDataAdapter
                                                    Dim oleDataSetLibelle As New DataSet
                                                    Dim oleDataTable As DataTable

                                                    oleDataAdapeterLibelle = New OleDbDataAdapter("SELECT DISTINCT P_UNITE.U_Intitule,P_UNITE.cbMarq FROM P_UNITE INNER JOIN F_ARTICLE ON P_UNITE.cbMarq=" & OledatableSchemaSage.Rows(i).Item("AR_UNITEVEN") & " AND F_ARTICLE.AR_REf='" & Element & "'", OleExcelConnectedArticle)
                                                    oleDataAdapeterLibelle.Fill(oleDataSetLibelle)
                                                    oleDataTable = oleDataSetLibelle.Tables(0)

                                                    If oleDataTable.Rows.Count <> 0 Then
                                                        Dim ValeurLue As String = oleDataTable.Rows(0).Item("U_Intitule")
                                                        If ValeurLue <> "" Then
                                                            Information &= Strings.LSet(oleDataTable.Rows(0).Item("U_Intitule"), PositionLeft)
                                                        Else
                                                            If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> Nothing Then ' s'il ya une valeur pas defaut
                                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                            Else
                                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                            End If
                                                        End If
                                                    Else
                                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> Nothing Then ' s'il ya une valeur pas defaut
                                                            Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                        Else
                                                            Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                        End If
                                                    End If
                                                Else
                                                    'si infos libre correspondant
                                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString) <> "1" Then
                                                        Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString), PositionLeft)
                                                    Else
                                                        Information &= Strings.LSet("", PositionLeft)
                                                    End If
                                                End If
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then ' s'il ya une valeur pas defaut
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                            Continue For '
                                        Case "DESC_UOM_BASE_CODE"
                                            PositionLeft = OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(1).ToString.Split(")")(0).Split(".")(0)
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString = "AR_UNITEVEN" Then

                                                    Dim oleDataAdapeterLibelle As OleDbDataAdapter
                                                    Dim oleDataSetLibelle As New DataSet
                                                    Dim oleDataTable As DataTable

                                                    oleDataAdapeterLibelle = New OleDbDataAdapter("SELECT * FROM P_UNITE INNER JOIN F_ARTICLE ON P_UNITE.cbMarq=" & OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString) & " AND F_ARTICLE.AR_REf='" & Element & "'", OleExcelConnectedArticle)
                                                    oleDataAdapeterLibelle.Fill(oleDataSetLibelle)
                                                    oleDataTable = oleDataSetLibelle.Tables(0)

                                                    If oleDataTable.Rows.Count <> 0 Then
                                                        Dim ValeurLue As String = oleDataTable.Rows(0).Item("U_Intitule")
                                                        If ValeurLue <> "" Then
                                                            Information &= Strings.LSet(oleDataTable.Rows(0).Item("U_Intitule"), PositionLeft)
                                                        Else
                                                            If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then ' s'il ya une valeur pas defaut
                                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                            Else
                                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                            End If
                                                        End If
                                                    Else
                                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then ' s'il ya une valeur pas defaut
                                                            Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                        Else
                                                            Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                        End If
                                                    End If
                                                Else
                                                    'si infos libre correspondant
                                                    If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString) <> "1" Then
                                                        Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString), PositionLeft)
                                                    Else
                                                        Information &= Strings.LSet("", PositionLeft)
                                                    End If
                                                End If
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then ' s'il ya une valeur pas defaut
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                            Continue For '
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
                                                    Dim Nbre As Integer = CounteLigne(OledatableSchemaSage.Rows(i).Item("AR_REF"))
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
                                                        Dim Nbre As Integer = CounteLigne(OledatableSchemaSage.Rows(i).Item("AR_REF"))
                                                        Information &= Strings.LSet(RenvoiValeurLue(Nbre, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
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
                                                        Dim Nbre As Integer = CounteLigne(OledatableSchemaSage.Rows(i).Item("AR_REF"))
                                                        Information &= Strings.LSet(RenvoiValeurLue(Nbre, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
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
                                                    Dim Nbre As Integer = CounteLigne(OledatableSchemaSage.Rows(i).Item("AR_REF"))
                                                    Dim ValeurLue As String = RenvoiValeurLue(Nbre, OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                                    Information &= Strings.LSet(ValeurLue, PositionLeft)
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
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                                If OledatableSchema.Rows(Ligne).Item("Cols").ToString = "LINE_COUNT" Then '
                                                    Dim Nbre As Integer = CounteLigne(OledatableSchemaSage.Rows(i).Item("AR_REF"))
                                                    Information &= Strings.LSet(RenvoiValeurLue(Nbre, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
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
                                Error_journalArt.WriteLine(Information)
                                Information = ""
                                CreationLigne(OledatableSchemaSage.Rows(i).Item("AR_REF"))
                                statut = True
                            Catch ex As Exception
                                MsgBox(ex.Message)
                            End Try
                            'Error_journalArt.WriteLine(Information)
                            'Information = ""
                            'CreationLigne(OledatableSchemaSage.Rows(i).Item("AR_REF"))
                        Else
                            '
                        End If
                    Else

                        If Ckmodifier.Checked = False And statut = False Then
                            Error_journalArt = File.AppendText(PathsFileFormatiers & "PRO" & version & Strings.Right(DateAndTime.Year(Now), 4) & Format(DateAndTime.Month(Now), "00") & Format(DateAndTime.Day(Now), "00") & Format(DateAndTime.Hour(Now), "00") & Format(DateAndTime.Minute(Now), "00") & Format(DateAndTime.Second(Now), "00") & ".txt")
                            statut = True
                        End If
                        Element = OledatableSchemaSage.Rows(i).Item(0)
                        cpteurs = i
                        If ListBox.InvokeRequired Then
                            Dim MonDelegate As New Evenement(AddressOf Traitement1)
                            ListBox.Invoke(MonDelegate)
                        Else
                            Traitement1()
                        End If
                        For Ligne As Integer = 0 To OledatableSchema.Rows.Count - 1
                            '-------------------------------------------------------------------------------------
                            If OledatableSchema.Rows(Ligne).Item("Cols").ToString = "OWNER_CODE" Then
                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                If Societe <> "" Then
                                    Information &= Strings.LSet(Societe, PositionLeft)
                                Else
                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                End If
                                Continue For
                            End If
                            Select Case OledatableSchema.Rows(Ligne).Item("Cols") '
                                Case "LOT_CONTROL"
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString = "AR_SuiviStock" Then
                                        PositionLeft = OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(1).ToString.Split(")")(0).Split(".")(0)
                                        If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString) = "5" Then
                                            Information &= Strings.LSet("1", PositionLeft)
                                        Else
                                            If OledatableSchema.Rows(Ligne).Item("DefaultValue") <> Nothing Then
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue"), PositionLeft)
                                            Else
                                                Information &= Strings.LSet("0", PositionLeft)
                                            End If
                                        End If
                                        Continue For '
                                    ElseIf OledatableSchema.Rows(Ligne).Item("InfosLibre").ToString = "True" Then
                                        PositionLeft = OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(1).ToString.Split(")")(0).Split(".")(0)
                                        If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString) = "5" Then
                                            Information &= Strings.LSet("1", PositionLeft)
                                        Else
                                            If OledatableSchema.Rows(Ligne).Item("DefaultValue") <> Nothing Then
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue"), PositionLeft)
                                            Else
                                                Information &= Strings.LSet("0", PositionLeft)
                                            End If
                                        End If
                                        Continue For '
                                    Else
                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue"), PositionLeft)
                                        Continue For
                                    End If
                                Case "SERIAL_NO_CONTROL"
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString = "AR_SuiviStock" Then
                                        PositionLeft = OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(1).ToString.Split(")")(0).Split(".")(0)
                                        If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString) = "1" Then
                                            Information &= Strings.LSet("1", PositionLeft)
                                        Else
                                            If OledatableSchema.Rows(Ligne).Item("DefaultValue") <> Nothing Then
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue"), PositionLeft)
                                            Else
                                                Information &= Strings.LSet("0", PositionLeft)
                                            End If
                                        End If
                                        Continue For
                                    ElseIf OledatableSchema.Rows(Ligne).Item("InfosLibre").ToString = "True" Then
                                        'si infos libre correspondant
                                        PositionLeft = OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(1).ToString.Split(")")(0).Split(".")(0)
                                        If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString) = "1" Then
                                            Information &= Strings.LSet("1", PositionLeft)
                                        Else
                                            If OledatableSchema.Rows(Ligne).Item("DefaultValue") <> Nothing Then
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue"), PositionLeft)
                                            Else
                                                Information &= Strings.LSet("0", PositionLeft)
                                            End If
                                        End If
                                        Continue For '
                                    Else
                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue"), PositionLeft)
                                        Continue For
                                    End If
                                Case "SUPPLIER_CODE"
                                    PositionLeft = OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(1).ToString.Split(")")(0).Split(".")(0)
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage") <> Nothing Then
                                        If OledatableSchema.Rows(Ligne).Item("ChampSage") <> "" Then
                                            OleAdaptaterschemaFourssAR = New OleDbDataAdapter("select CT_NUM FROM F_ARTFOURNISS  WHERE F_ARTFOURNISS.AR_Ref='" & OledatableSchemaSage.Rows(i).Item("AR_REF") & "' AND F_ARTFOURNISS.AF_PRINCIPAL =1", OleExcelConnectedArticle)
                                            OleSchemaDatasetFourssAR = New DataSet
                                            OleAdaptaterschemaFourssAR.Fill(OleSchemaDatasetFourssAR)
                                            OledatableSchemaFourssAR = OleSchemaDatasetFourssAR.Tables(0)
                                            If OledatableSchemaFourssAR.Rows.Count > 0 Then
                                                If OledatableSchemaFourssAR.Rows(0).Item("CT_NUM").ToString = "" Then
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue"), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchemaFourssAR.Rows(0).Item("CT_NUM").ToString, PositionLeft)
                                                End If
                                            Else
                                                Information &= Strings.LSet("", PositionLeft)
                                            End If
                                            Continue For
                                        Else
                                            Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue"), PositionLeft)
                                            Continue For
                                        End If
                                    Else
                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue"), PositionLeft)
                                        Continue For
                                    End If
                                Case "UOM_BASE_CODE"
                                    PositionLeft = OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(1).ToString.Split(")")(0).Split(".")(0)
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                        If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString = "AR_UNITEVEN" Then
                                            Dim oleDataAdapeterLibelle As OleDbDataAdapter
                                            Dim oleDataSetLibelle As New DataSet
                                            Dim oleDataTable As DataTable

                                            oleDataAdapeterLibelle = New OleDbDataAdapter("SELECT DISTINCT P_UNITE.U_Intitule,P_UNITE.cbMarq FROM P_UNITE INNER JOIN F_ARTICLE ON P_UNITE.cbMarq=" & OledatableSchemaSage.Rows(i).Item("AR_UNITEVEN") & " AND F_ARTICLE.AR_REf='" & Element & "'", OleExcelConnectedArticle)
                                            oleDataAdapeterLibelle.Fill(oleDataSetLibelle)
                                            oleDataTable = oleDataSetLibelle.Tables(0)

                                            If oleDataTable.Rows.Count <> 0 Then
                                                Dim ValeurLue As String = oleDataTable.Rows(0).Item("U_Intitule")
                                                If ValeurLue <> "" Then
                                                    Information &= Strings.LSet(oleDataTable.Rows(0).Item("U_Intitule"), PositionLeft)
                                                Else
                                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> Nothing Then ' s'il ya une valeur pas defaut
                                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                    Else
                                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                    End If
                                                End If
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> Nothing Then ' s'il ya une valeur pas defaut
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        Else
                                            'si infos libre correspondant
                                            If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString) <> "1" Then
                                                Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet("", PositionLeft)
                                            End If
                                        End If
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then ' s'il ya une valeur pas defaut
                                            Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                        Else
                                            Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                        End If
                                    End If
                                    Continue For '
                                Case "DESC_UOM_BASE_CODE"
                                    PositionLeft = OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(1).ToString.Split(")")(0).Split(".")(0)
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then
                                        If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString = "AR_UNITEVEN" Then

                                            Dim oleDataAdapeterLibelle As OleDbDataAdapter
                                            Dim oleDataSetLibelle As New DataSet
                                            Dim oleDataTable As DataTable

                                            oleDataAdapeterLibelle = New OleDbDataAdapter("SELECT * FROM P_UNITE INNER JOIN F_ARTICLE ON P_UNITE.cbMarq=" & OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString) & " AND F_ARTICLE.AR_REf='" & Element & "'", OleExcelConnectedArticle)
                                            oleDataAdapeterLibelle.Fill(oleDataSetLibelle)
                                            oleDataTable = oleDataSetLibelle.Tables(0)

                                            If oleDataTable.Rows.Count <> 0 Then
                                                Dim ValeurLue As String = oleDataTable.Rows(0).Item("U_Intitule")
                                                If ValeurLue <> "" Then
                                                    Information &= Strings.LSet(oleDataTable.Rows(0).Item("U_Intitule"), PositionLeft)
                                                Else
                                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then ' s'il ya une valeur pas defaut
                                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                    Else
                                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                    End If
                                                End If
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then ' s'il ya une valeur pas defaut
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        Else
                                            'si infos libre correspondant
                                            If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString) <> "1" Then
                                                Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet("", PositionLeft)
                                            End If
                                        End If
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then ' s'il ya une valeur pas defaut
                                            Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                        Else
                                            Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                        End If
                                    End If
                                    Continue For '
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
                                            Dim Nbre As Integer = CounteLigne(OledatableSchemaSage.Rows(i).Item("AR_REF"))
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
                                                Dim Nbre As Integer = CounteLigne(OledatableSchemaSage.Rows(i).Item("AR_REF"))
                                                Information &= Strings.LSet(RenvoiValeurLue(Nbre, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
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
                                                Dim Nbre As Integer = CounteLigne(OledatableSchemaSage.Rows(i).Item("AR_REF"))
                                                Information &= Strings.LSet(RenvoiValeurLue(Nbre, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
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
                                            Dim Nbre As Integer = CounteLigne(OledatableSchemaSage.Rows(i).Item("AR_REF"))
                                            Dim ValeurLue As String = RenvoiValeurLue(Nbre, OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                            Information &= Strings.LSet(ValeurLue, PositionLeft)
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
                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                        If OledatableSchema.Rows(Ligne).Item("Cols").ToString = "LINE_COUNT" Then '
                                            Dim Nbre As Integer = CounteLigne(OledatableSchemaSage.Rows(i).Item("AR_REF"))
                                            Information &= Strings.LSet(RenvoiValeurLue(Nbre, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
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
                        Error_journalArt.WriteLine(Information)
                        Information = ""
                        CreationLigne(OledatableSchemaSage.Rows(i).Item("AR_REF"))
                        statut = True
                    End If
                Next
                If statut = True Then
                    Error_journalArt.Close()
                End If
                If Ckmodifier.Checked Then
                    IDernierDate(1)
                    statut = False
                    Error_journalArt.Close()
                End If
            Else
                'cette requette ne possede pas ligne 
                EstVide = True
            End If
            Return True
        Catch ex As Exception
            'MsgBox(ex.Message)
            Return False
        End Try
    End Function
    Public Function EstCreerOuMoudiffier(ByVal code As String) As String
        Dim etat As Integer = 0
        Try
            Dim OleAdaptaterschemaM As OleDbDataAdapter
            Dim OleSchemaDatasetM As DataSet
            Dim oleSchemaTableM As DataTable
            OleSchemaDatasetM = New DataSet
            OleAdaptaterschemaM = New OleDbDataAdapter("SELECT " & FlagtamponArticle & " FROM F_ARTICLE WHERE  AR_REF='" & code & "' AND "& FlagtamponArticle &" IS NULL OR "& FlagtamponArticle & " =''", OleExcelConnectedArticle)
            OleAdaptaterschemaM.Fill(OleSchemaDatasetM)
            oleSchemaTableM = OleSchemaDatasetM.Tables(0)
            Dim dr = oleSchemaTableM.Rows(0).Item(0)
            If IsDBNull(oleSchemaTableM.Rows(0).Item(0)) = False Then
                Return "M"
            Else
                Return "A"
            End If
        Catch ex As Exception
            Return "M"
        End Try
    End Function
    Public EstVide As Boolean = False
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
                        For Cpte As Integer = 0 To Tableau(0).Split(")")(0) - 1
                            FormatFloat &= "0"
                        Next
                        Dim g = CDbl(Valeur.ToString.Replace(".", ",")).ToString(FormatFloat).Replace(",", "")
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
        Try
            EcritureEntete()
        Catch ex As Exception
        End Try
    End Sub
    Public Function CounteLigne(ByVal codeArticle As String) As Integer
        Try
            If ChBoxEC_Qté.Checked Then
                Dim MaCommande As New OleDbCommand("select Count(*) FROM F_CONDITION,F_ARTICLE WHERE F_CONDITION.EC_Quantite<>1 AND F_CONDITION.AR_REF=F_ARTICLE.AR_REF AND F_CONDITION.AR_REF='" & codeArticle & "'", OleExcelConnectedArticle)
                Dim h = MaCommande.ExecuteScalar()
                Return CInt(MaCommande.ExecuteScalar())
            Else
                Dim MaCommande As New OleDbCommand("select Count(*) FROM F_CONDITION,F_ARTICLE WHERE F_CONDITION.AR_REF=F_ARTICLE.AR_REF AND F_CONDITION.AR_REF='" & codeArticle & "'", OleExcelConnectedArticle)
                Return CInt(MaCommande.ExecuteScalar())
            End If
        Catch ex As Exception
        End Try
    End Function
    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            'PictureBox1.Visible = False
            If EstVide Then
                lblsms.Text = "Aucune donnée retournée par ce Filtre"
                lblsms.Visible = True
            Else
                If ifRowError > 0 Then
                    DataListeIntegrer.Rows(selectids(ips)).Cells("Status").Value = My.Resources.criticalind_status
                    ifRowError = 0
                Else
                    DataListeIntegrer.Rows(selectids(ips)).Cells("Status").Value = My.Resources.accepter
                    If TachePlanifie <> "" Then
                        'My.Computer.Audio.Play("CIRCUS1.wav")
                    End If
                End If
            End If
            If ips < selectids.Length - 1 Then
                ips = ips + 1
                BackgroundWorker1.RunWorkerAsync()
            Else
                ips = 0
                nbrelignes = 0
                selectindexe = ""
            End If
        Catch ex As Exception
            '   MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Me.Hide()
    End Sub
    Public TypeSuivi, InfosLibre As String
    Private Sub ComboSuivi_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboSuivi.SelectedValueChanged
        TypeSuivi = ""
        TypeSuivi = ComboSuivi.Text
    End Sub

    Private Sub ComboInfosLibre_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboInfosLibre.SelectedValueChanged
        InfosLibre = ""
        InfosLibre = ComboInfosLibre.Text
    End Sub

    Private Sub lblsms_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblsms.Click

    End Sub

    Private Sub BackgroundWorker4_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker4.DoWork
        Try
            Dim OleAdaptaterschema As OleDbDataAdapter
            Dim OleSchemaDataset As DataSet

            DataListeIntegrer.Rows.Clear()
            OleAdaptaterschema = New OleDbDataAdapter("select * from PARAMETRE WHERE nomtype='COMMERCIAL'", OleConnenectionArticle)
            OleSchemaDataset = New DataSet
            OleAdaptaterschema.Fill(OleSchemaDataset)
            OledatableSchema = OleSchemaDataset.Tables(0)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub BackgroundWorker4_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker4.RunWorkerCompleted
        Try
            AfficheSchemasConso()
        Catch ex As Exception

        End Try
    End Sub
End Class