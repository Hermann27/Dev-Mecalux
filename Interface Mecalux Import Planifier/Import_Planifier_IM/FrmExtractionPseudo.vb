Imports System.Data.OleDb
Imports System.IO
Public Class FrmExtractionPseudo
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
    Public ContinuTraitement As Integer = 0
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
            If TachePlanifie = "Export Pseudos" Then
                For Each Item As String In SocieteCyble
                    If Item = OledatableSchema.Rows(i).Item("Societe") Then
                        DataListeIntegrer.Rows(i).Cells("choix").Value = True
                    End If
                Next
            End If
        Next i
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
    Public Shared ip As Integer
    Public Sub BtnModif_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnModif.Click
        nbreligne = DataListeIntegrer.Rows.Count
        ListBox.Items.Clear()
        ContinuTraitement = 0
        NbInfosLibre = 0
        NbInfosLibreVue = 0
        lblsms.Text = "0/0"
        selectindexe1 = Nothing
        selectids1 = Nothing
        Dim Etat As Boolean = False
        Try
            For i As Integer = 0 To DataListeIntegrer.Rows.Count - 1
                If DataListeIntegrer.Rows(i).Cells("Choix").Value = True Then
                    selectindexe1 &= DataListeIntegrer.Rows(i).Index & ";"
                    Etat = True
                End If
                DataListeIntegrer.Rows(i).Cells("Status").Value = My.Resources.btFermer22
            Next
            If Etat = True Then
                lblInfos.Visible = False
                selectindexe1 = selectindexe1.Substring(0, selectindexe1.Length - 1)
                selectids1 = selectindexe1.Split(";")
                VerificateurChampObligatoire(True, False)
                If ContinuTraitement = 3 Then
                    EstInfosLibre(True, False, True)
                    If NbInfosLibre = NbInfosLibreVue Then 'NbInfosLibreVue
                        Try
                            'PictureBox1.Visible = True
                            ListeChampSageEntete = RecuperationColonneSage(True)
                            lblSne.Text = "Scénario Extraction des Pseudo"
                            ListeChampSageEntete = ListeChampSageEntete.Substring(0, ListeChampSageEntete.Length - 1)
                            BackgroundWorker1.RunWorkerAsync()
                        Catch ex As Exception
                            MsgBox(ex.Message)
                        End Try
                    End If
                End If
            Else
                lblInfos.Text = "Aucune Société Selectionnée !"
                lblInfos.Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Function CounteLigne(ByVal codeArticle As String) As Integer
        Try
            Dim MaCommande As New OleDbCommand("select Count(*) FROM F_CONDITION,F_ARTICLE WHERE F_CONDITION.AR_REF=F_ARTICLE.AR_REF AND F_CONDITION.AR_REF='" & codeArticle & "'", OleExcelConnectedPseudo)
            Return CInt(MaCommande.ExecuteScalar())
        Catch ex As Exception
        End Try
    End Function
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
            End If
            OleAdaptaterschemaVersion = New OleDbDataAdapter("select version from P_TABLECORRESP WHERE CodeTbls='ALI'", OleConnenectionArticle)
            OleSchemaDatasetVersion = New DataSet
            OleAdaptaterschemaVersion.Fill(OleSchemaDatasetVersion)
            OledatableSchemaVersion = OleSchemaDatasetVersion.Tables(0)
        Catch ex As Exception
        End Try
    End Sub
    Public Sub VerificateurChampObligatoire(ByVal Entete As Boolean, ByVal Ligne As Boolean)
        OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='ALI' AND Entete=" & Entete & " AND Ligne=" & Ligne & "  ORDER BY ORDRE", OleConnenectionPseudo)
        OleSchemaDataset = New DataSet
        OleAdaptaterschema.Fill(OleSchemaDataset)
        OledatableSchema = OleSchemaDataset.Tables(0)
        lblSne.Text = "Verification conformite des Pseudo"
        ListBox.Items.Add("<-----------------------------------------------------------------DEBUT TRAITEMENT DE VERIFICATION DES CHAMPS OBLIGATOIRES----------------------------------------------------------------->")
        If OledatableSchema.Rows.Count <> 0 Then
            For i As Integer = 0 To OledatableSchema.Rows.Count - 1
                Select Case OledatableSchema.Rows(i).Item("Cols")
                    Case "OPERATION"
                        If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False Then
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
                    Case "ALIAS_CODE"
                        If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False Then
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
                    Case "OWNER_CODE"
                        If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False Then
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
        Else
            ListBox.Items.Add("Aucun résultat n'est fourni pour ce Scénario de traitement")
        End If
        ListBox.Items.Add("<-----------------------------------------------------------------Fin----------------------------------------------------------------->")
        ListBox.Items.Add("")
    End Sub
    Public Societe As String = ""
    Public Function RecuperationColonneSage(ByVal Entete As Boolean) As String
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
        OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='ALI' AND Entete=" & Entete & " ORDER BY ORDRE", OleConnenectionPseudo)

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
                        Dim LaRequeteVerificationColonne As String = "SELECT COUNT(*) FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='F_ARTICLE' AND COLUMN_NAME = '" & OledatableSchema.Rows(i).Item("ChampSage").ToString & "'"
                        Dim MaCommande As New OleDbCommand(LaRequeteVerificationColonne, OleExcelConnectedPseudo)
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
                                    ListBox.Items.Add("Le Champ {" & OledatableSchema.Rows(i).Item("Cols") & "} EasyWMS ne possede pas de Mapping Sage infos sur le Pseudo ")
                                End If
                            Else
                                If OledatableSchema.Rows(i).Item("ChampSage") <> "" Then
                                    ListeChampSage &= OledatableSchema.Rows(i).Item("ChampSage") & ","
                                    ListBox.Items.Add("Recuperation de la colonne en Entete : " & OledatableSchema.Rows(i).Item("ChampSage"))
                                Else
                                    ListBox.Items.Add("Le Champ {" & OledatableSchema.Rows(i).Item("Cols") & "} EasyWMS ne possede pas de Mapping Sage infos sur le Pseudo")
                                End If
                            End If
                        Else
                            If i = OledatableSchema.Rows.Count - 1 Then
                                If OledatableSchema.Rows(i).Item("ChampSage") <> "" Then
                                    ListeChampSage &= OledatableSchema.Rows(i).Item("ChampSage")
                                    ListBox.Items.Add("Recuperation de la colonne en Ligne : " & OledatableSchema.Rows(i).Item("ChampSage"))
                                Else
                                    ListBox.Items.Add("Le Champ {" & OledatableSchema.Rows(i).Item("Cols") & "} EasyWMS ne possede pas de Mapping Sage infos sur le Pseudo ")
                                End If
                            Else
                                If OledatableSchema.Rows(i).Item("ChampSage") <> "" Then
                                    ListeChampSage &= OledatableSchema.Rows(i).Item("ChampSage") & ","
                                    ListBox.Items.Add("Recuperation de la colonne en Ligne : " & OledatableSchema.Rows(i).Item("ChampSage"))
                                Else
                                    ListBox.Items.Add("Le Champ {" & OledatableSchema.Rows(i).Item("Cols") & "} EasyWMS ne possede pas de Mapping Sage infos sur le Pseudo ")
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
        lblSne.Text = "Scenario verification infos libre"
        Try
            OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='ALI' AND Entete=" & Entete & " AND InfosLibre=" & InfosLibre & " AND Ligne=" & Ligne, OleConnenectionPseudo)
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
                                    ListBox.Items.Add("la zone infos libre {" & OledatableSchema.Rows(i).Item("ChampSage") & "} en Entete du Pseudo n'existe pas dans Sage")
                                Else
                                    ListBox.Items.Add("Traitement  de la zone infos libre {" & OledatableSchema.Rows(i).Item("ChampSage") & "} sur le Pseudo Existe [OK] dans Sage")
                                End If
                            Else
                                ListBox.Items.Add("<--Le Champ indiquant l'infos libre est couché mais ne possede pas de mapping Sage-->")
                            End If
                        ElseIf Ligne = True Then
                            If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False Then
                                If ExisteInfosLibre(OledatableSchema.Rows(i).Item("ChampSage")) = False Then
                                    ListBox.Items.Add("la zone infos libre {" & OledatableSchema.Rows(i).Item("ChampSage") & "} sur le Pseudo n'existe pas dans Sage")
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
                OleAdaptaterschemaSage = New OleDbDataAdapter("select * from cbSysLibre WHERE CB_File='F_ARTICLE' And CB_Name='" & Join(Split(Trim(InfosLibre), "'"), "''") & "'", OleExcelConnectedPseudo)
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
            Return False
        End Try
    End Function
    Dim EstChoisir As Boolean = False
    Public Sub ExtractionPseudo(ByVal SelectChamp As String, ByVal indice As Integer)
        Dim Information As String = ""
        Dim ExistColonne As Boolean = False
        Dim ColonnName As String = ""
        Dim PositionLeft As Object = Nothing
        Try
            If ExcelConnect(DataListeIntegrer.Rows(indice).Cells("Serveur1").Value, DataListeIntegrer.Rows(indice).Cells("Societe1").Value, DataListeIntegrer.Rows(indice).Cells("NomUtil").Value, DataListeIntegrer.Rows(indice).Cells("Mot").Value) Then
                Societe = DataListeIntegrer.Rows(indice).Cells("Societe1").Value
                If Ckmodifier.Checked Then
                    If CodeEDI = True Then
                        If SelectChamp IsNot Nothing Then
                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif,AR_CodeEdiED_Code1 FROM F_ARTICLE WHERE AR_SuiviStock<>0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedPseudo) ' AND AR_CodeBarre <>'' AND AR_CodeEdiED_Code1 <>''
                        Else
                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock<>0 ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedPseudo) 'AND AR_CodeBarre  <>'' AND AR_CodeEdiED_Code1 <>'' 
                        End If
                    Else
                        If SelectChamp IsNot Nothing Then
                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif,AR_CodeEdiED_Code1 FROM F_ARTICLE WHERE AR_SuiviStock<>0   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedPseudo)
                        Else
                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock<>0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedPseudo)
                        End If
                    End If
                Else
                    If CodeEDI = True Then
                        If SelectChamp IsNot Nothing Then
                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif,AR_CodeEdiED_Code1 FROM F_ARTICLE WHERE AR_SuiviStock<>0    ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedPseudo) ' AND AR_CodeBarre <>'' AND AR_CodeEdiED_Code1 <>''
                        Else
                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock<>0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedPseudo) 'AND AR_CodeBarre <>'' AND AR_CodeEdiED_Code1 <>''
                        End If
                    Else
                        If SelectChamp IsNot Nothing Then
                            OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",AR_DateModif,AR_CodeEdiED_Code1 FROM F_ARTICLE WHERE AR_SuiviStock<>0   ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedPseudo)
                        Else
                            OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_ARTICLE WHERE AR_SuiviStock<>0  ORDER BY F_ARTICLE.AR_REF ", OleExcelConnectedPseudo)
                        End If
                    End If
                End If
                OleSchemaDatasetSage = New DataSet
                OleAdaptaterschemaSage.Fill(OleSchemaDatasetSage)
                OledatableSchemaSage = OleSchemaDatasetSage.Tables(0)

                OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='ALI' AND Entete=True ORDER BY ORDRE ", OleConnenectionPseudo)
                OleSchemaDataset = New DataSet
                OleAdaptaterschema.Fill(OleSchemaDataset)
                OledatableSchema = OleSchemaDataset.Tables(0)
                EstChoisir = True
            Else
                ErreurJrnPseudo = File.AppendText(Pathsfilejournal & "ERREURALI09" & Strings.Right(DateAndTime.Year(Now), 4) & Format(DateAndTime.Month(Now), "00") & Format(DateAndTime.Day(Now), "00") & Format(DateAndTime.Hour(Now), "00") & Format(DateAndTime.Minute(Now), "00") & Format(DateAndTime.Second(Now), "00") & ".txt")
                ErreurJrnPseudo.WriteLine("Erreur de Connexion à la Société <[" & DataListeIntegrer.Rows(indice).Cells("Societe1").Value & "]>")
                ErreurJrnPseudo.Close()
            End If
        Catch ex As Exception
            'MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            ExtractionPseudo(ListeChampSageEntete, selectids1(ip))
        Catch ex As Exception
            ifRowerror = selectids1(ip)
            'DataListeIntegrer.Rows(selectids1(ip)).Cells("Status").Value = My.Resources.btSupprimer221
            ErreurJrnPseudo = File.AppendText(Pathsfilejournal & "ERREURALI09" & Strings.Right(DateAndTime.Year(Now), 4) & Format(DateAndTime.Month(Now), "00") & Format(DateAndTime.Day(Now), "00") & Format(DateAndTime.Hour(Now), "00") & Format(DateAndTime.Minute(Now), "00") & Format(DateAndTime.Second(Now), "00") & ".txt")
            ErreurJrnPseudo.WriteLine("ERREUR SYSTEME : " & ex.Message)
            ErreurJrnPseudo.Close()
        End Try

    End Sub
    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        BackgroundWorker2.RunWorkerAsync()
    End Sub
    Delegate Sub Evenement()
    Public Element As String
    Public Elements As String
    Public cpteur As Integer = 1
    Public EstVide As Boolean = False
    Public Sub Traitement1()
        ListBox.Items.Add("Extraction du Pseudo : <[" & Element & "]> ")
        lblsms.Text = cpteur + 1 & "/" & OledatableSchemaSage.Rows.Count
        ListBox.SelectedIndex = ListBox.Items.Count - 1
    End Sub
    Public Function RDernierDate(ByRef IDTraitement As Object) As Date
        Try
            Dim OleAdaptater As OleDbDataAdapter
            Dim OleAfficheDataset As DataSet
            Dim Oledatable As DataTable
            OleAdaptater = New OleDbDataAdapter("select * from WEAT_ADATE WHERE IDDossier=" & IDTraitement & " And TypeExport='Export Pseudo'", OleConnenectionPseudo)
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
    Public Function EcritureEntete() As Boolean
        Dim Information As String = ""
        Dim ExistColonne As Boolean = False
        Dim ColonnName As String = ""
        Dim PositionLeft As Object = Nothing
        Dim PositionLeftLigne As Integer = 0
        Dim ColonnNameLigne As String = ""
        Dim statut As Boolean = False
        Dim P_CONDITIONNEMENT As String = ""
        Dim Situation As Boolean = False
        Dim oleAdatarSchemaP_CON As OleDbDataAdapter = Nothing
        Dim oleShemanDataSetP_CON As DataSet
        Dim oleSchemaTable As DataTable
        Dim AR_CodeBarre, AR_CodeEdiED_Code1 As String
        Dim version As String = "09"
        AR_CodeBarre = ""
        AR_CodeEdiED_Code1 = ""
        cpteur = 0
        Try
            If OledatableSchemaSage.Rows.Count <> 0 Then
                EstVide = False
                version = OledatableSchemaVersion.Rows(0).Item("version")
                Dim DateTraitement As Date = RDernierDate(5)
                For i As Integer = 0 To OledatableSchemaSage.Rows.Count - 1
                    If Ckmodifier.Checked Then
                        If IsDBNull(OledatableSchemaSage.Rows(i).Item("AR_CodeBarre")) = False Then
                            AR_CodeBarre = OledatableSchemaSage.Rows(i).Item("AR_CodeBarre")
                        Else
                            AR_CodeBarre = ""
                        End If
                        If IsDBNull(OledatableSchemaSage.Rows(i).Item("AR_CodeEdiED_Code1")) = False Then
                            AR_CodeEdiED_Code1 = OledatableSchemaSage.Rows(i).Item("AR_CodeEdiED_Code1")
                        Else
                            AR_CodeEdiED_Code1 = ""
                        End If
                        If ArticleRecemmentModifier(DateTraitement, OledatableSchemaSage, i) = True Then
                            If Ckmodifier.Checked = True And statut = False Then
                                Error_journalPseudo = File.AppendText(PathsFileFormatiers & "ALI" & version & Strings.Right(DateAndTime.Year(Now), 4) & Format(DateAndTime.Month(Now), "00") & Format(DateAndTime.Day(Now), "00") & Format(DateAndTime.Hour(Now), "00") & Format(DateAndTime.Minute(Now), "00") & Format(DateAndTime.Second(Now), "00") & ".txt")
                                statut = True
                            End If
                            Element = OledatableSchemaSage.Rows(i).Item("AR_REF")
                            If ListBox.InvokeRequired Then
                                Dim MonDelegate As New Evenement(AddressOf Traitement1)
                                ListBox.Invoke(MonDelegate)
                            Else
                                Traitement1()
                            End If
                            cpteur = i + 1
                            If CodeEDI = True Then
                                cpteur = i + 1
                                For Ligne As Integer = 0 To OledatableSchema.Rows.Count - 1
                                    If OledatableSchema.Rows(Ligne).Item("Cols").ToString = "OWNER_CODE" Then
                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                        If Societe <> "" Then
                                            Information &= Strings.LSet(Societe, PositionLeft)
                                        Else
                                            Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                        End If
                                        Continue For
                                    End If
                                    If Ligne = 1 And Situation = True Then
                                        PositionLeft = GetPosition(OledatableSchema.Rows(0).Item("Format").ToString)
                                        Information &= Strings.LSet(OledatableSchema.Rows(0).Item("DefaultValue").ToString, PositionLeft)
                                    End If
                                    If OledatableSchema.Rows(Ligne).Item("Cols").ToString = "UOM_CODE" Then
                                        If Situation = False Then
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> Nothing Or OledatableSchema.Rows(Ligne).Item("ChampSage").ToString = Nothing Then
                                                oleAdatarSchemaP_CON = New OleDbDataAdapter("SELECT AR_uniteven FROM F_ARTICLE WHERE F_ARTICLE.AR_Ref='" & Element & "'", OleExcelConnectedPseudo)
                                                oleShemanDataSetP_CON = New DataSet

                                                oleAdatarSchemaP_CON.Fill(oleShemanDataSetP_CON)
                                                oleSchemaTable = oleShemanDataSetP_CON.Tables(0)
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                                ColonnNameLigne = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString

                                                If ColonnNameLigne <> Nothing Or ColonnNameLigne = Nothing Then
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                                    Else
                                                        If oleSchemaTable.Rows.Count <> 0 Then
                                                            P_CONDITIONNEMENT = oleSchemaTable.Rows(0).Item("AR_uniteven").ToString
                                                            'Information &= Strings.LSet(P_CONDITIONNEMENT, PositionLeft)
                                                            Dim oleDataAdapeterLibelle As OleDbDataAdapter
                                                            Dim oleDataSetLibelle As New DataSet
                                                            Dim oleDataTable As DataTable

                                                            oleDataAdapeterLibelle = New OleDbDataAdapter("SELECT U_Intitule FROM P_UNITE INNER JOIN F_ARTICLE ON P_UNITE.CbIndice=" & P_CONDITIONNEMENT & " AND F_ARTICLE.AR_REf='" & Element & "'", OleExcelConnectedPseudo)
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
                                                            Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                        End If
                                                    End If
                                                ElseIf OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> Nothing Then
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                    Else
                                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                    End If
                                                Else
                                                    'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                                        Information &= Strings.LSet("", PositionLeft)
                                                    Else
                                                        Information &= Strings.LSet("", PositionLeft)
                                                    End If
                                                End If
                                            ElseIf OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> Nothing Then
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            Else
                                                'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                                    Information &= Strings.LSet("", PositionLeft)
                                                Else
                                                    Information &= Strings.LSet("", PositionLeft)
                                                End If
                                            End If
                                            If AR_CodeBarre.Trim <> Nothing Then
                                                Error_journalPseudo.WriteLine(Information)
                                            End If
                                            Information = ""
                                            Situation = True
                                            If AR_CodeEdiED_Code1 <> Nothing Then
                                                Information &= Strings.LSet("F", 1)
                                                Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item("AR_CodeEdiED_Code1").ToString, 50)
                                                Ligne = 1
                                            Else
                                                Situation = False
                                            End If
                                            Continue For
                                        Else
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> Nothing Or OledatableSchema.Rows(Ligne).Item("ChampSage").ToString = Nothing Then
                                                oleAdatarSchemaP_CON = New OleDbDataAdapter("SELECT P_CONDITIONNEMENT FROM F_ARTICLE INNER JOIN P_CONDITIONNEMENT ON F_ARTICLE.AR_CONDITION = P_CONDITIONNEMENT.CbMarq AND F_ARTICLE.AR_Ref='" & Element & "'", OleExcelConnectedPseudo)
                                                oleShemanDataSetP_CON = New DataSet

                                                oleAdatarSchemaP_CON.Fill(oleShemanDataSetP_CON)
                                                oleSchemaTable = oleShemanDataSetP_CON.Tables(0)
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                                ColonnNameLigne = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString
                                                If ColonnNameLigne <> Nothing Or ColonnNameLigne = Nothing Then
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                                    Else
                                                        If oleSchemaTable.Rows.Count <> 0 Then
                                                            P_CONDITIONNEMENT = oleSchemaTable.Rows(0).Item("P_CONDITIONNEMENT").ToString
                                                            Information &= Strings.LSet(P_CONDITIONNEMENT, PositionLeft)
                                                        Else
                                                            Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                        End If
                                                    End If
                                                ElseIf OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> Nothing Then
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                    Else
                                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                    End If
                                                Else
                                                    'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                                        Information &= Strings.LSet("", PositionLeft)
                                                    Else
                                                        Information &= Strings.LSet("", PositionLeft)
                                                    End If
                                                End If
                                            ElseIf OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> Nothing Then
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            Else
                                                'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                                    Information &= Strings.LSet("", PositionLeft)
                                                Else
                                                    Information &= Strings.LSet("", PositionLeft)
                                                End If
                                            End If
                                            If AR_CodeEdiED_Code1 <> Nothing Then
                                                Error_journalPseudo.WriteLine(Information)
                                                Information = ""
                                                Situation = False
                                            End If
                                            Continue For
                                        End If
                                    End If
                                    '-------------------------------------******************************------------------------------------------
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> Nothing Then 's'il existe un mappig Sage 
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
                                            If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> Nothing Then
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> Nothing Then
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
                                    '-------------------------------------**********FIN********************------------------------------------------
                                Next
                            Else
                                '-----------------------------------------------------------SI CodeEDI = False---------------------
                                cpteur = i + 1
                                For Ligne As Integer = 0 To OledatableSchema.Rows.Count - 1
                                    If OledatableSchema.Rows(Ligne).Item("Cols").ToString = "OWNER_CODE" Then
                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                        If Societe <> "" Then
                                            Information &= Strings.LSet(Societe, PositionLeft)
                                        Else
                                            Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                        End If
                                        Continue For
                                    End If
                                    If Ligne = 1 And Situation = True Then
                                        PositionLeft = GetPosition(OledatableSchema.Rows(0).Item("Format").ToString)
                                        Information &= Strings.LSet(OledatableSchema.Rows(0).Item("DefaultValue").ToString, PositionLeft)
                                    End If
                                    If OledatableSchema.Rows(Ligne).Item("Cols").ToString = "UOM_CODE" Then
                                        If Situation = False Then
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> Nothing Or OledatableSchema.Rows(Ligne).Item("ChampSage").ToString = Nothing Then
                                                oleAdatarSchemaP_CON = New OleDbDataAdapter("SELECT AR_uniteven FROM F_ARTICLE WHERE F_ARTICLE.AR_Ref='" & Element & "'", OleExcelConnectedPseudo)
                                                oleShemanDataSetP_CON = New DataSet

                                                oleAdatarSchemaP_CON.Fill(oleShemanDataSetP_CON)
                                                oleSchemaTable = oleShemanDataSetP_CON.Tables(0)
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                                ColonnNameLigne = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString

                                                If ColonnNameLigne <> Nothing Or ColonnNameLigne = Nothing Then
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                                    Else
                                                        If oleSchemaTable.Rows.Count <> 0 Then
                                                            P_CONDITIONNEMENT = oleSchemaTable.Rows(0).Item("AR_uniteven").ToString
                                                            'Information &= Strings.LSet(P_CONDITIONNEMENT, PositionLeft)
                                                            Dim oleDataAdapeterLibelle As OleDbDataAdapter
                                                            Dim oleDataSetLibelle As New DataSet
                                                            Dim oleDataTable As DataTable

                                                            oleDataAdapeterLibelle = New OleDbDataAdapter("SELECT U_Intitule FROM P_UNITE INNER JOIN F_ARTICLE ON P_UNITE.CbIndice=" & P_CONDITIONNEMENT & " AND F_ARTICLE.AR_REf='" & Element & "'", OleExcelConnectedPseudo)
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
                                                            Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                        End If
                                                    End If
                                                ElseIf OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> Nothing Then
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                    Else
                                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                    End If
                                                Else
                                                    'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                                        Information &= Strings.LSet("", PositionLeft)
                                                    Else
                                                        Information &= Strings.LSet("", PositionLeft)
                                                    End If
                                                End If
                                            ElseIf OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> Nothing Then
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            Else
                                                'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                                    Information &= Strings.LSet("", PositionLeft)
                                                Else
                                                    Information &= Strings.LSet("", PositionLeft)
                                                End If
                                            End If
                                            If AR_CodeBarre.Trim <> Nothing Then
                                                Error_journalPseudo.WriteLine(Information)
                                            End If
                                            Information = ""
                                            Situation = True
                                            If AR_CodeEdiED_Code1 <> Nothing Then
                                                Information &= Strings.LSet("F", 1)
                                                Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item("AR_CodeEdiED_Code1").ToString, 50)
                                                Ligne = 1
                                            Else
                                                Situation = False
                                            End If
                                            Continue For
                                        Else
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> Nothing Or OledatableSchema.Rows(Ligne).Item("ChampSage").ToString = Nothing Then
                                                oleAdatarSchemaP_CON = New OleDbDataAdapter("SELECT P_CONDITIONNEMENT FROM F_ARTICLE INNER JOIN P_CONDITIONNEMENT ON F_ARTICLE.AR_CONDITION = P_CONDITIONNEMENT.CbMarq AND F_ARTICLE.AR_Ref='" & Element & "'", OleExcelConnectedPseudo)
                                                oleShemanDataSetP_CON = New DataSet

                                                oleAdatarSchemaP_CON.Fill(oleShemanDataSetP_CON)
                                                oleSchemaTable = oleShemanDataSetP_CON.Tables(0)
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                                ColonnNameLigne = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString
                                                If ColonnNameLigne <> Nothing Or ColonnNameLigne = Nothing Then
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                                    Else
                                                        If oleSchemaTable.Rows.Count <> 0 Then
                                                            P_CONDITIONNEMENT = oleSchemaTable.Rows(0).Item("P_CONDITIONNEMENT").ToString
                                                            Information &= Strings.LSet(P_CONDITIONNEMENT, PositionLeft)
                                                        Else
                                                            Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                        End If
                                                    End If
                                                ElseIf OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> Nothing Then
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                    Else
                                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                    End If
                                                Else
                                                    'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                        'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                                        Information &= Strings.LSet("", PositionLeft)
                                                    Else
                                                        Information &= Strings.LSet("", PositionLeft)
                                                    End If
                                                End If
                                            ElseIf OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> Nothing Then
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            Else
                                                'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                                    Information &= Strings.LSet("", PositionLeft)
                                                Else
                                                    Information &= Strings.LSet("", PositionLeft)
                                                End If
                                            End If
                                            If AR_CodeEdiED_Code1 <> Nothing Then
                                                Error_journalPseudo.WriteLine(Information)
                                                Information = ""
                                                Situation = False
                                            End If
                                            Continue For
                                        End If
                                    End If
                                    '-------------------------------------******************************------------------------------------------
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> Nothing Then 's'il existe un mappig Sage 
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
                                            If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> Nothing Then
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> Nothing Then
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
                                    '-------------------------------------**********FIN********************------------------------------------------
                                Next
                            End If
                        End If
                    Else
                        If IsDBNull(OledatableSchemaSage.Rows(i).Item("AR_CodeBarre")) = False Then
                            AR_CodeBarre = OledatableSchemaSage.Rows(i).Item("AR_CodeBarre")
                        Else
                            AR_CodeBarre = ""
                        End If

                        If IsDBNull(OledatableSchemaSage.Rows(i).Item("AR_CodeEdiED_Code1")) = False Then
                            AR_CodeEdiED_Code1 = OledatableSchemaSage.Rows(i).Item("AR_CodeEdiED_Code1")
                        Else
                            AR_CodeEdiED_Code1 = ""
                        End If
                        If Ckmodifier.Checked = False And statut = False Then
                            Error_journalPseudo = File.AppendText(PathsFileFormatiers & "ALI" & version & Strings.Right(DateAndTime.Year(Now), 4) & Format(DateAndTime.Month(Now), "00") & Format(DateAndTime.Day(Now), "00") & Format(DateAndTime.Hour(Now), "00") & Format(DateAndTime.Minute(Now), "00") & Format(DateAndTime.Second(Now), "00") & ".txt")
                            statut = True
                        End If
                        Element = OledatableSchemaSage.Rows(i).Item("AR_REF")
                        If ListBox.InvokeRequired Then
                            Dim MonDelegate As New Evenement(AddressOf Traitement1)
                            ListBox.Invoke(MonDelegate)
                        Else
                            Traitement1()
                        End If

                        If CodeEDI = True Then
                            cpteur = i + 1
                            For Ligne As Integer = 0 To OledatableSchema.Rows.Count - 1
                                If OledatableSchema.Rows(Ligne).Item("Cols").ToString = "OWNER_CODE" Then
                                    PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                    If Societe <> "" Then
                                        Information &= Strings.LSet(Societe, PositionLeft)
                                    Else
                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                    End If
                                    Continue For
                                End If
                                If Ligne = 1 And Situation = True Then
                                    PositionLeft = GetPosition(OledatableSchema.Rows(0).Item("Format").ToString)
                                    Information &= Strings.LSet(OledatableSchema.Rows(0).Item("DefaultValue").ToString, PositionLeft)
                                End If
                                If OledatableSchema.Rows(Ligne).Item("Cols").ToString = "UOM_CODE" Then
                                    If Situation = False Then
                                        If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> Nothing Or OledatableSchema.Rows(Ligne).Item("ChampSage").ToString = Nothing Then
                                            oleAdatarSchemaP_CON = New OleDbDataAdapter("SELECT AR_uniteven FROM F_ARTICLE WHERE F_ARTICLE.AR_Ref='" & Element & "'", OleExcelConnectedPseudo)
                                            oleShemanDataSetP_CON = New DataSet

                                            oleAdatarSchemaP_CON.Fill(oleShemanDataSetP_CON)
                                            oleSchemaTable = oleShemanDataSetP_CON.Tables(0)
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                            ColonnNameLigne = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString

                                            If ColonnNameLigne <> Nothing Or ColonnNameLigne = Nothing Then
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                                Else
                                                    If oleSchemaTable.Rows.Count <> 0 Then
                                                        P_CONDITIONNEMENT = oleSchemaTable.Rows(0).Item("AR_uniteven").ToString
                                                        'Information &= Strings.LSet(P_CONDITIONNEMENT, PositionLeft)
                                                        Dim oleDataAdapeterLibelle As OleDbDataAdapter
                                                        Dim oleDataSetLibelle As New DataSet
                                                        Dim oleDataTable As DataTable

                                                        oleDataAdapeterLibelle = New OleDbDataAdapter("SELECT U_Intitule FROM P_UNITE INNER JOIN F_ARTICLE ON P_UNITE.CbIndice=" & P_CONDITIONNEMENT & " AND F_ARTICLE.AR_REf='" & Element & "'", OleExcelConnectedPseudo)
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
                                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                    End If
                                                End If
                                            ElseIf OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> Nothing Then
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            Else
                                                'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                                    Information &= Strings.LSet("", PositionLeft)
                                                Else
                                                    Information &= Strings.LSet("", PositionLeft)
                                                End If
                                            End If
                                        ElseIf OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> Nothing Then
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        Else
                                            'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                                Information &= Strings.LSet("", PositionLeft)
                                            Else
                                                Information &= Strings.LSet("", PositionLeft)
                                            End If
                                        End If
                                        If AR_CodeBarre.Trim <> Nothing Then
                                            Error_journalPseudo.WriteLine(Information)
                                        End If
                                        Information = ""
                                        Situation = True
                                        If AR_CodeEdiED_Code1 <> Nothing Then
                                            Information &= Strings.LSet("F", 1)
                                            Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item("AR_CodeEdiED_Code1").ToString, 50)
                                            Ligne = 1
                                        Else
                                            Situation = False
                                        End If
                                        Continue For
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> Nothing Or OledatableSchema.Rows(Ligne).Item("ChampSage").ToString = Nothing Then
                                            oleAdatarSchemaP_CON = New OleDbDataAdapter("SELECT P_CONDITIONNEMENT FROM F_ARTICLE INNER JOIN P_CONDITIONNEMENT ON F_ARTICLE.AR_CONDITION = P_CONDITIONNEMENT.CbMarq AND F_ARTICLE.AR_Ref='" & Element & "'", OleExcelConnectedPseudo)
                                            oleShemanDataSetP_CON = New DataSet

                                            oleAdatarSchemaP_CON.Fill(oleShemanDataSetP_CON)
                                            oleSchemaTable = oleShemanDataSetP_CON.Tables(0)
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                            ColonnNameLigne = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString
                                            If ColonnNameLigne <> Nothing Or ColonnNameLigne = Nothing Then
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                                Else
                                                    If oleSchemaTable.Rows.Count <> 0 Then
                                                        P_CONDITIONNEMENT = oleSchemaTable.Rows(0).Item("P_CONDITIONNEMENT").ToString
                                                        Information &= Strings.LSet(P_CONDITIONNEMENT, PositionLeft)
                                                    Else
                                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                    End If
                                                End If
                                            ElseIf OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> Nothing Then
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            Else
                                                'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                                    Information &= Strings.LSet("", PositionLeft)
                                                Else
                                                    Information &= Strings.LSet("", PositionLeft)
                                                End If
                                            End If
                                        ElseIf OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> Nothing Then
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        Else
                                            'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                                Information &= Strings.LSet("", PositionLeft)
                                            Else
                                                Information &= Strings.LSet("", PositionLeft)
                                            End If
                                        End If
                                        If AR_CodeEdiED_Code1 <> Nothing Then
                                            Error_journalPseudo.WriteLine(Information)
                                            Information = ""
                                            Situation = False
                                        End If
                                        Continue For
                                    End If
                                End If
                                '-------------------------------------******************************------------------------------------------
                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> Nothing Then 's'il existe un mappig Sage 
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> Nothing Then
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
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> Nothing Then
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
                                '-------------------------------------**********FIN********************------------------------------------------
                            Next
                        Else '-------------------------------------------------------------------------------
                            cpteur = i + 1
                            For Ligne As Integer = 0 To OledatableSchema.Rows.Count - 1
                                If OledatableSchema.Rows(Ligne).Item("Cols").ToString = "OWNER_CODE" Then
                                    PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                    If Societe <> "" Then
                                        Information &= Strings.LSet(Societe, PositionLeft)
                                    Else
                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                    End If
                                    Continue For
                                End If
                                If Ligne = 1 And Situation = True Then
                                    PositionLeft = GetPosition(OledatableSchema.Rows(0).Item("Format").ToString)
                                    Information &= Strings.LSet(OledatableSchema.Rows(0).Item("DefaultValue").ToString, PositionLeft)
                                End If
                                If OledatableSchema.Rows(Ligne).Item("Cols").ToString = "UOM_CODE" Then
                                    If Situation = False Then
                                        If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> Nothing Or OledatableSchema.Rows(Ligne).Item("ChampSage").ToString = Nothing Then
                                            oleAdatarSchemaP_CON = New OleDbDataAdapter("SELECT AR_uniteven FROM F_ARTICLE WHERE F_ARTICLE.AR_Ref='" & Element & "'", OleExcelConnectedPseudo)
                                            oleShemanDataSetP_CON = New DataSet

                                            oleAdatarSchemaP_CON.Fill(oleShemanDataSetP_CON)
                                            oleSchemaTable = oleShemanDataSetP_CON.Tables(0)
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                            ColonnNameLigne = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString

                                            If ColonnNameLigne <> Nothing Or ColonnNameLigne = Nothing Then
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                                Else
                                                    If oleSchemaTable.Rows.Count <> 0 Then
                                                        P_CONDITIONNEMENT = oleSchemaTable.Rows(0).Item("AR_uniteven").ToString
                                                        'Information &= Strings.LSet(P_CONDITIONNEMENT, PositionLeft)
                                                        Dim oleDataAdapeterLibelle As OleDbDataAdapter
                                                        Dim oleDataSetLibelle As New DataSet
                                                        Dim oleDataTable As DataTable

                                                        oleDataAdapeterLibelle = New OleDbDataAdapter("SELECT U_Intitule FROM P_UNITE INNER JOIN F_ARTICLE ON P_UNITE.CbIndice=" & P_CONDITIONNEMENT & " AND F_ARTICLE.AR_REf='" & Element & "'", OleExcelConnectedPseudo)
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
                                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                    End If
                                                End If
                                            ElseIf OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> Nothing Then
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            Else
                                                'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                                    Information &= Strings.LSet("", PositionLeft)
                                                Else
                                                    Information &= Strings.LSet("", PositionLeft)
                                                End If
                                            End If
                                        ElseIf OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> Nothing Then
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        Else
                                            'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                                Information &= Strings.LSet("", PositionLeft)
                                            Else
                                                Information &= Strings.LSet("", PositionLeft)
                                            End If
                                        End If
                                        If AR_CodeBarre.Trim <> Nothing Then
                                            Error_journalPseudo.WriteLine(Information)
                                        End If
                                        Information = ""
                                        Situation = True
                                        If AR_CodeEdiED_Code1 <> Nothing Then
                                            Information &= Strings.LSet("F", 1)
                                            Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item("AR_CodeEdiED_Code1").ToString, 50)
                                            Ligne = 1
                                        Else
                                            Situation = False
                                        End If
                                        Continue For
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> Nothing Or OledatableSchema.Rows(Ligne).Item("ChampSage").ToString = Nothing Then
                                            oleAdatarSchemaP_CON = New OleDbDataAdapter("SELECT P_CONDITIONNEMENT FROM F_ARTICLE INNER JOIN P_CONDITIONNEMENT ON F_ARTICLE.AR_CONDITION = P_CONDITIONNEMENT.CbMarq AND F_ARTICLE.AR_Ref='" & Element & "'", OleExcelConnectedPseudo)
                                            oleShemanDataSetP_CON = New DataSet

                                            oleAdatarSchemaP_CON.Fill(oleShemanDataSetP_CON)
                                            oleSchemaTable = oleShemanDataSetP_CON.Tables(0)
                                            PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                            ColonnNameLigne = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString
                                            If ColonnNameLigne <> Nothing Or ColonnNameLigne = Nothing Then
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnNameLigne).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                                Else
                                                    If oleSchemaTable.Rows.Count <> 0 Then
                                                        P_CONDITIONNEMENT = oleSchemaTable.Rows(0).Item("P_CONDITIONNEMENT").ToString
                                                        Information &= Strings.LSet(P_CONDITIONNEMENT, PositionLeft)
                                                    Else
                                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                    End If
                                                End If
                                            ElseIf OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> Nothing Then
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            Else
                                                'ici le mapping Sage existe mais ne possede pas de valeur ni de valeur pas defaut parametre  et est un champ obligatoire
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                                    Information &= Strings.LSet("", PositionLeft)
                                                Else
                                                    Information &= Strings.LSet("", PositionLeft)
                                                End If
                                            End If
                                        ElseIf OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> Nothing Then
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            End If
                                        Else
                                            'ici le champ Sage ne possede pas de mapping ni de valeur pas defautparametre est un champ Obligatoire
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                'Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchemaLigne.Rows(Ligne).Item("Format").ToString), PositionLeftLigne)
                                                Information &= Strings.LSet("", PositionLeft)
                                            Else
                                                Information &= Strings.LSet("", PositionLeft)
                                            End If
                                        End If
                                        If AR_CodeEdiED_Code1 <> Nothing Then
                                            Error_journalPseudo.WriteLine(Information)
                                            Information = ""
                                            Situation = False
                                        End If
                                        Continue For
                                    End If
                                End If
                                '-------------------------------------******************************------------------------------------------
                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> Nothing Then 's'il existe un mappig Sage 
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> Nothing Then
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
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> Nothing Then
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
                                '-------------------------------------**********FIN********************------------------------------------------
                            Next
                        End If '--------------------------------------------------------------------------------------------------------------
                        statut = True
                    End If
                Next
                If statut = True Then
                    Error_journalPseudo.Close()
                End If
                If Ckmodifier.Checked Then
                    IDernierDate(5)
                    statut = False
                End If
            Else
                'cette requette ne possede pas ligne 
                EstVide = True
            End If
            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
            Error_journalPseudo.Close()
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
            OleAdaptater = New OleDbDataAdapter("select * from WEAT_ADATE WHERE IDDossier=" & IDTraitement & " And TypeExport='Export Pseudo'", OleConnenectionPseudo)
            OleAfficheDataset = New DataSet
            OleAdaptater.Fill(OleAfficheDataset)
            Oledatable = OleAfficheDataset.Tables(0)
            If Oledatable.Rows.Count <> 0 Then
                Insertion = " UPDATE WEAT_ADATE SET DateDern='" & Strings.FormatDateTime(Now, Microsoft.VisualBasic.DateFormat.ShortDate) & "' WHERE IDDossier=" & IDTraitement & " And TypeExport='Export Pseudo'"
                OleCmdIns = New OleDbCommand(Insertion)
                OleCmdIns.Connection = OleConnenectionPseudo
                OleCmdIns.ExecuteNonQuery()
            Else
                Insertion = " Insert Into WEAT_ADATE (TypeExport,IDDossier,DateDern) Values ('Export Pseudo'," & IDTraitement & ",'" & Strings.FormatDateTime(Now, Microsoft.VisualBasic.DateFormat.ShortDate) & "')"
                OleCmdIns = New OleDbCommand(Insertion)
                OleCmdIns.Connection = OleConnenectionPseudo
                OleCmdIns.ExecuteNonQuery()
            End If
        Catch ex As Exception
        End Try
    End Function
    Public Function GetPosition(ByVal Format As Object) As Integer
        Try
            Dim Position() As Object = Format.ToString.Split("(")(1).ToString.Split(")")(0).Split(".")

            If Position.Length = 2 Then
                Return Position(0)
            Else
                Return Position(0)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
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
        If EstChoisir Then
            EcritureEntete()
        End If
    End Sub
    Dim ifRowerror = 0
    Private Sub BackgroundWorker2_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        Try
            'PictureBox1.Visible = False
            If EstVide Or EstChoisir = False Then
                lblInfos.Text = "Aucune donnée retournée par ce Filtre"
                lblInfos.Visible = True
            Else
                If ifRowerror > 0 Then
                    DataListeIntegrer.Rows(selectids1(ip)).Cells("Status").Value = My.Resources.criticalind_status
                    ifRowerror = 0
                Else
                    DataListeIntegrer.Rows(selectids1(ip)).Cells("Status").Value = My.Resources.accepter
                End If
            End If
            If ip < selectids1.Length - 1 Then
                ip = ip + 1
                BackgroundWorker1.RunWorkerAsync()
            Else
                ip = 0
                nbreligne = 0
                selectindexe1 = ""
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

    Private Sub BackgroundWorker4_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker4.DoWork
        Try
            Dim OleAdaptaterschema As OleDbDataAdapter
            Dim OleSchemaDataset As DataSet
            DataListeIntegrer.Rows.Clear()
            OleAdaptaterschema = New OleDbDataAdapter("select * from PARAMETRE WHERE nomtype='COMMERCIAL'", OleConnenectionPseudo)
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

    Private Sub SplitContainer1_SplitterMoved(ByVal sender As System.Object, ByVal e As System.Windows.Forms.SplitterEventArgs) Handles SplitContainer1.SplitterMoved

    End Sub
End Class