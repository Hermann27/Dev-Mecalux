Imports System.Data.OleDb
Imports System.IO
Public Class FrmExtractionFournisseurs
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
            If TachePlanifie = "Export Fournisseur" Then
                For Each Item As String In SocieteCyble
                    If Item = OledatableSchema.Rows(i).Item("Societe") Then
                        DataListeIntegrer.Rows(i).Cells("choix").Value = True
                    End If
                Next
            End If
        Next i
    End Sub
    Public Sub BtnModif_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnModif.Click
        nbreligne = DataListeIntegrer.Rows.Count
        ListBox.Items.Clear()
        ContinuTraitement = 0
        NbInfosLibre = 0
        NbInfosLibreVue = 0
        lblsms.Text = "0/0"
        lblInfos.Visible = False
        selectindexe2 = Nothing
        selectids2 = Nothing
        Dim Etat As Boolean = False
        Try
            For i As Integer = 0 To DataListeIntegrer.Rows.Count - 1
                If DataListeIntegrer.Rows(i).Cells("Choix").Value = True Then
                    selectindexe2 &= DataListeIntegrer.Rows(i).Index & ";"
                    Etat = True
                End If
                DataListeIntegrer.Rows(i).Cells("Status").Value = My.Resources.btFermer22
            Next
            If Etat = True Then
                selectindexe2 = selectindexe2.Substring(0, selectindexe2.Length - 1)
                selectids2 = selectindexe2.Split(";")
                VerificateurChampObligatoire(True, False)
                If ContinuTraitement = 4 Then
                    EstInfosLibre(True, False, True)
                    If NbInfosLibre = NbInfosLibreVue Then 'NbInfosLibreVue
                        Try
                            'PictureBox1.Visible = True
                            ListeChampSageEntete = RecuperationColonneSage(True, False)
                            lblSne.Text = "Sc�nario Extraction des Fournisseurs"
                            ListeChampSageEntete = ListeChampSageEntete.Substring(0, ListeChampSageEntete.Length - 1)
                            BackgroundWorker1.RunWorkerAsync()
                        Catch ex As Exception
                            MsgBox(ex.Message)
                        End Try
                    End If
                End If
            Else
                lblInfos.Text = "Aucune Soci�t� Selectionn�e !"
                lblInfos.Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
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

                OleAdaptaterschemaVersion = New OleDbDataAdapter("select version from P_TABLECORRESP WHERE CodeTbls='SUP'", OleConnenectionArticle)
                OleSchemaDatasetVersion = New DataSet
                OleAdaptaterschemaVersion.Fill(OleSchemaDatasetVersion)
                OledatableSchemaVersion = OleSchemaDatasetVersion.Tables(0)
            End If
        Catch ex As Exception
        End Try
    End Sub
    Public ContinuTraitement As Integer = 0
    Public Sub VerificateurChampObligatoire(ByVal Entete As Boolean, ByVal Ligne As Boolean)
        OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='SUP' AND Entete=" & Entete & " AND Ligne=" & Ligne, OleConnenection)
        OleSchemaDataset = New DataSet
        OleAdaptaterschema.Fill(OleSchemaDataset)
        OledatableSchema = OleSchemaDataset.Tables(0)
        lblSne.Text = "Verification conformite des fournisseurs"
        ListBox.Items.Add("<-----------------------------------------------------------------DEBUT DU TRAITEMENT DE LA VERIFICATION DES CHAMPS OBLIGATOIRES----------------------------------------------------------------->")
        If OledatableSchema.Rows.Count <> 0 Then
            For i As Integer = 0 To OledatableSchema.Rows.Count - 1
                Select Case OledatableSchema.Rows(i).Item("Cols")
                    Case "OP�RATION"
                        If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False And OledatableSchema.Rows(i).Item("ChampSage") <> "" Then
                            ContinuTraitement += 1
                            ListBox.Items.Add("Un Mapping Sage {" & OledatableSchema.Rows(i).Item("ChampSage") & "} a �t� appliquer sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}*")
                        Else
                            If IsDBNull(OledatableSchema.Rows(i).Item("DefaultValue")) = False Then
                                ListBox.Items.Add("Une Valeur par defaut a �t� param�tr�e sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}*")
                                ContinuTraitement += 1
                            Else
                                ListBox.Items.Add("Aucun Mapping Sage n'est param�tre sur  le champ {" & OledatableSchema.Rows(i).Item("Cols") & "}* EasyWMS  ni de valeur par defaut n'a ete param�tr�e {Ce Champ est Obligatoire}")
                            End If
                        End If
                    Case "SUPPLIER_CODE"
                        If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False And OledatableSchema.Rows(i).Item("ChampSage") <> "" Then
                            ContinuTraitement += 1
                            ListBox.Items.Add("Un Mapping Sage {" & OledatableSchema.Rows(i).Item("ChampSage") & "} a �t� appliquer sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}*")
                        Else
                            If IsDBNull(OledatableSchema.Rows(i).Item("DefaultValue")) = False Then
                                ListBox.Items.Add("Une Valeur par defaut a �t� param�tr�e sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}*")
                                ContinuTraitement += 1
                            Else
                                ListBox.Items.Add("Aucun Mapping Sage n'est param�tre sur  le champ {" & OledatableSchema.Rows(i).Item("Cols") & "*} EasyWMS  ni de valeur par defaut n'a ete param�tr�e {Ce Champ est Obligatoire}")
                            End If
                        End If
                    Case "ADDR_NAME_FROM"
                        If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False And OledatableSchema.Rows(i).Item("ChampSage") <> "" Then
                            ContinuTraitement += 1
                            ListBox.Items.Add("Un Mapping Sage {" & OledatableSchema.Rows(i).Item("ChampSage") & "} a �t� appliquer sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}*")
                        Else
                            If IsDBNull(OledatableSchema.Rows(i).Item("DefaultValue")) = False Then
                                ListBox.Items.Add("Une Valeur par defaut a �t� param�tr�e sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}*")
                                ContinuTraitement += 1
                            Else
                                ListBox.Items.Add("Aucun Mapping Sage n'est param�tre sur  le champ {" & OledatableSchema.Rows(i).Item("Cols") & "}* EasyWMS  ni de valeur par defaut n'a ete param�tr�e {Ce Champ est Obligatoire}")
                            End If
                        End If
                    Case "ADDR_NAME_TO"
                        If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False And OledatableSchema.Rows(i).Item("ChampSage") <> "" Then
                            ContinuTraitement += 1
                            ListBox.Items.Add("Un Mapping Sage {" & OledatableSchema.Rows(i).Item("ChampSage") & "} a �t� appliquer sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}*")
                        Else
                            If IsDBNull(OledatableSchema.Rows(i).Item("DefaultValue")) = False Then
                                ListBox.Items.Add("Une Valeur par defaut a �t� param�tr�e sur  le champ EasyWMS {" & OledatableSchema.Rows(i).Item("Cols") & "}*")
                                ContinuTraitement += 1
                            Else
                                ListBox.Items.Add("Aucun Mapping Sage n'est param�tre sur  le champ {" & OledatableSchema.Rows(i).Item("Cols") & "}* EasyWMS  ni de valeur par defaut n'a ete param�tr�e {Ce Champ est Obligatoire}")
                            End If
                        End If
                End Select
            Next i
        Else
            ListBox.Items.Add("Aucun r�sultat n'est fourni pour ce Sc�nario de traitement")
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
            OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='SUP' AND (Entete=" & Entete & " OR InfosLibre=true) AND Ligne=False ORDER BY ORDRE", OleConnenection)
        Else
            OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='SUP' AND (Ligne=" & Ligne & " OR InfosLibre=true) AND Entete=False ORDER BY ORDRE", OleConnenection)
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
                        Dim MaCommande As New OleDbCommand(LaRequeteVerificationColonne, OleExcelConnectedFrs)
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
                                    ListBox.Items.Add("Le Champ {" & OledatableSchema.Rows(i).Item("Cols") & "} EasyWMS ne possede pas de Mapping Sage infos sur le Fournisseur ")
                                End If
                            Else
                                If OledatableSchema.Rows(i).Item("ChampSage") <> "" Then
                                    ListeChampSage &= OledatableSchema.Rows(i).Item("ChampSage") & ","
                                    ListBox.Items.Add("Recuperation de la colonne en Entete : " & OledatableSchema.Rows(i).Item("ChampSage"))
                                Else
                                    ListBox.Items.Add("Le Champ {" & OledatableSchema.Rows(i).Item("Cols") & "} EasyWMS ne possede pas de Mapping Sage infos sur le Fournisseur ")
                                End If
                            End If
                        Else
                            If i = OledatableSchema.Rows.Count - 1 Then
                                If OledatableSchema.Rows(i).Item("ChampSage") <> "" Then
                                    ListeChampSage &= OledatableSchema.Rows(i).Item("ChampSage")
                                    ListBox.Items.Add("Recuperation de la colonne en Ligne : " & OledatableSchema.Rows(i).Item("ChampSage"))
                                Else
                                    ListBox.Items.Add("Le Champ {" & OledatableSchema.Rows(i).Item("Cols") & "} EasyWMS ne possede pas de Mapping Sage infos sur le Fournisseur")
                                End If
                            Else
                                If OledatableSchema.Rows(i).Item("ChampSage") <> "" Then
                                    ListeChampSage &= OledatableSchema.Rows(i).Item("ChampSage") & ","
                                    ListBox.Items.Add("Recuperation de la colonne en Ligne : " & OledatableSchema.Rows(i).Item("ChampSage"))
                                Else
                                    ListBox.Items.Add("Le Champ {" & OledatableSchema.Rows(i).Item("Cols") & "} EasyWMS ne possede pas de Mapping Sage infos sur le Fournisseur")
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
            OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='SUP' AND Entete=" & Entete & " AND InfosLibre=" & InfosLibre & " AND Ligne=" & Ligne, OleConnenection)
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
                                    ListBox.Items.Add("Traitement  de la zone infos libre {" & OledatableSchema.Rows(i).Item("ChampSage") & "} sur le fournisseur Existe [OK] dans Sage")
                                End If
                            Else
                                ListBox.Items.Add("<--Le Champ indiquant l'infos libre est couch� mais ne possede pas de mapping Sage-->")
                            End If
                        ElseIf Ligne = True Then
                            If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False Then
                                If ExisteInfosLibre(OledatableSchema.Rows(i).Item("ChampSage")) = False Then
                                    ListBox.Items.Add("la zone infos libre {" & OledatableSchema.Rows(i).Item("ChampSage") & "} sur le fournisseur n'existe pas dans Sage")
                                End If
                            Else
                                ListBox.Items.Add("<--Le Champ indiquant l'infos libre est couch� mais ne possede pas de mapping Sage-->")
                            End If
                        Else
                            ListBox.Items.Add("<--Aucune information libre n'est parametr�e-->")
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
                OleAdaptaterschemaSage = New OleDbDataAdapter("select * from cbSysLibre WHERE CB_File='F_COMPTET' And CB_Name='" & Join(Split(Trim(InfosLibre), "'"), "''") & "'", OleExcelConnectedFrs)
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

    Public Sub ExtractionFournisseur(ByVal SelectChamp As String, ByVal indice As Integer)
        Dim Information As String = ""
        Dim ExistColonne As Boolean = False
        Dim ColonnName As String = ""
        Dim PositionLeft As Object = Nothing
        Try
            If ExcelConnect(DataListeIntegrer.Rows(indice).Cells("Serveur1").Value, DataListeIntegrer.Rows(indice).Cells("Societe1").Value, DataListeIntegrer.Rows(indice).Cells("NomUtil").Value, DataListeIntegrer.Rows(indice).Cells("Mot").Value) Then

                If Ckmodifier.Checked Then
                    If SelectChamp IsNot Nothing Then
                        OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",CBModification FROM F_COMPTET WHERE CT_Type=1  ORDER BY F_COMPTET.CT_NUM ", OleExcelConnectedFrs)
                    Else
                        OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_COMPTET WHERE CT_Type=1  ORDER BY F_COMPTET.CT_NUM ", OleExcelConnectedFrs)
                    End If
                Else
                    If SelectChamp IsNot Nothing Then
                        OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & ",CBModification FROM F_COMPTET WHERE CT_Type=1  ORDER BY F_COMPTET.CT_NUM", OleExcelConnectedFrs)
                    Else
                        OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_COMPTET WHERE CT_Type=1  ORDER BY F_COMPTET.CT_NUM", OleExcelConnectedFrs)
                    End If
                End If
                OleSchemaDatasetSage = New DataSet
                OleAdaptaterschemaSage.Fill(OleSchemaDatasetSage)
                OledatableSchemaSage = OleSchemaDatasetSage.Tables(0)

                OleAdaptaterschema = New OleDbDataAdapter("select * from TEMPLATE WHERE CodeTbls='SUP' AND Entete=True AND Societe='" & DataListeIntegrer.Rows(indice).Cells("Societe1").Value & "' ORDER BY ORDRE ", OleConnenection)
                OleSchemaDataset = New DataSet
                OleAdaptaterschema.Fill(OleSchemaDataset)
                OledatableSchema = OleSchemaDataset.Tables(0)

                If OledatableSchema.Rows.Count = 0 Then
                    OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='SUP' AND Entete=True ORDER BY ORDRE ", OleConnenection)
                    OleSchemaDataset = New DataSet
                    OleAdaptaterschema.Fill(OleSchemaDataset)
                    OledatableSchema = OleSchemaDataset.Tables(0)
                End If
                If OledatableSchemaSage.Rows.Count <> 0 Then
                    EstChoisir = True
                End If
            Else
                ErreurJrnFrs = File.AppendText(Pathsfilejournal & "ERREURSUP05" & Strings.Right(DateAndTime.Year(Now), 4) & Format(DateAndTime.Month(Now), "00") & Format(DateAndTime.Day(Now), "00") & Format(DateAndTime.Hour(Now), "00") & Format(DateAndTime.Minute(Now), "00") & Format(DateAndTime.Second(Now), "00") & ".txt")
                ErreurJrnFrs.WriteLine("Erreur de Connexion � la Soci�t� <[" & DataListeIntegrer.Rows(indice).Cells("Societe1").Value & "]>")
                ErreurJrnFrs.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Dim ifRowError = 0
    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Try
            ExtractionFournisseur(ListeChampSageEntete, selectids2(ip))
        Catch ex As Exception
            ifRowError = selectids2(ip)
            ErreurJrnFrs = File.AppendText(Pathsfilejournal & "ERREURSUP05" & Strings.Right(DateAndTime.Year(Now), 4) & Format(DateAndTime.Month(Now), "00") & Format(DateAndTime.Day(Now), "00") & Format(DateAndTime.Hour(Now), "00") & Format(DateAndTime.Minute(Now), "00") & Format(DateAndTime.Second(Now), "00") & ".txt")
            ErreurJrnFrs.WriteLine("ERREUR SYSTEME : " & ex.Message)
            ErreurJrnFrs.Close()
        End Try
    End Sub
    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As System.Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        BackgroundWorker2.RunWorkerAsync()
    End Sub
    Public EstVide As Boolean = False
    Delegate Sub Evenement()
    Public Element As String
    Public cpteur As Integer
    Public Sub Traitement1()
        ListBox.Items.Add("Extraction du Fournisseur : " & Element)
        lblsms.Text = cpteur + 1 & "/" & OledatableSchemaSage.Rows.Count
        ListBox.SelectedIndex = ListBox.Items.Count - 1
    End Sub
    Public Function RDernierDate(ByRef IDTraitement As Object) As Date
        Try
            Dim OleAdaptater As OleDbDataAdapter
            Dim OleAfficheDataset As DataSet
            Dim Oledatable As DataTable
            OleAdaptater = New OleDbDataAdapter("select * from WEAT_ADATE WHERE IDDossier=" & IDTraitement & " And TypeExport='Export Fournisseur'", OleConnenection)
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
            If DatDerntraitement <= OleArtModifierDt.Rows(m).Item("CBModification") Then
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
            OleAdaptater = New OleDbDataAdapter("select * from WEAT_ADATE WHERE IDDossier=" & IDTraitement & " And TypeExport='Export Fournisseur'", OleConnenection)
            OleAfficheDataset = New DataSet
            OleAdaptater.Fill(OleAfficheDataset)
            Oledatable = OleAfficheDataset.Tables(0)
            If Oledatable.Rows.Count <> 0 Then
                Insertion = " UPDATE WEAT_ADATE SET DateDern='" & Strings.FormatDateTime(Now, Microsoft.VisualBasic.DateFormat.ShortDate) & "' WHERE IDDossier=" & IDTraitement & " And TypeExport='Export Fournisseur'"
                OleCmdIns = New OleDbCommand(Insertion)
                OleCmdIns.Connection = OleConnenection
                OleCmdIns.ExecuteNonQuery()
            Else
                Insertion = " Insert Into WEAT_ADATE (TypeExport,IDDossier,DateDern) Values ('Export Fournisseur'," & IDTraitement & ",'" & Strings.FormatDateTime(Now, Microsoft.VisualBasic.DateFormat.ShortDate) & "')"
                OleCmdIns = New OleDbCommand(Insertion)
                OleCmdIns.Connection = OleConnenection
                OleCmdIns.ExecuteNonQuery()
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
        Dim version As String = "05"
        Try
            If OledatableSchemaSage.Rows.Count <> 0 Then
                Dim DateTraitement As Date = RDernierDate(3)
                Version = OledatableSchemaVersion.Rows(0).Item("version")
                For i As Integer = 0 To OledatableSchemaSage.Rows.Count - 1
                    If Ckmodifier.Checked Then
                        If ArticleRecemmentModifier(DateTraitement, OledatableSchemaSage, i) = True Then
                            If Ckmodifier.Checked And statut = False Then
                                Error_journalFrs = File.AppendText(PathsFileFormatiers & "SUP" & version & Strings.Right(DateAndTime.Year(Now), 4) & Format(DateAndTime.Month(Now), "00") & Format(DateAndTime.Day(Now), "00") & Format(DateAndTime.Hour(Now), "00") & Format(DateAndTime.Minute(Now), "00") & Format(DateAndTime.Second(Now), "00") & ".txt")
                                statut = True
                            End If
                            Element = OledatableSchemaSage.Rows(i).Item("CT_NUM")
                            cpteur = i
                            If ListBox.InvokeRequired Then
                                Dim MonDelegate As New Evenement(AddressOf Traitement1)
                                ListBox.Invoke(MonDelegate)
                            Else
                                Traitement1()
                            End If

                            Try
                                For Ligne As Integer = 0 To OledatableSchema.Rows.Count - 1
                                    '-------------------------------------------------------------------------------------
                                    Dim CARACTERE As String = OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0).ToString
                                    If CARACTERE <> "V" Then
                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString)
                                    End If
                                    Select Case OledatableSchema.Rows(Ligne).Item("Cols") '
                                        Case "OP�RATION"
                                            If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            Else
                                                Information &= Strings.LSet("F", PositionLeft)
                                            End If
                                            Continue For
                                        Case "SUPPLIER_CODE"
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                                If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                                    ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                        Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                        Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                                    Else 'sinon le format n'est pas du type variable
                                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                        Case "DESCRIPTION"
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                                If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                                    ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                        Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                        Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                                    Else 'sinon le format n'est pas du type variable
                                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                        Case "SUPPLIERCLASS_CODE"
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                                If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                                    ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                        Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                        Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                                    Else 'sinon le format n'est pas du type variable
                                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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

                                        Case "ADDR_NAME_FROM"
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                                If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                                    ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                        Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                        Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                                    Else 'sinon le format n'est pas du type variable
                                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                        Case "ADDR_LINE1_FROM"
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                                If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                                    ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                        Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                        Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                                    Else 'sinon le format n'est pas du type variable
                                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                        Case "ADDR_LINE2_FROM"
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                                If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                                    ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                        Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                        Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                                    Else 'sinon le format n'est pas du type variable
                                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                                            Information &= Strings.LSet("", PositionLeft)
                                                        End If
                                                    End If
                                                End If
                                            End If
                                            Continue For
                                        Case "CITY_FROM"
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                                If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                                    ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                        Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                        Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                                    Else 'sinon le format n'est pas du type variable
                                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                        Case "STATE_FROM"
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                                If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                                    ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                        Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                        Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                                    Else 'sinon le format n'est pas du type variable
                                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                        Case "COUNTRY_FROM"
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                                If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                                    ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                        Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                        Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                                    Else 'sinon le format n'est pas du type variable
                                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                        Case "ZIP_FROM"
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                                If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                                    ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                        Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                        Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                                    Else 'sinon le format n'est pas du type variable
                                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                        Case "CONTACT_NAME_FROM"
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                                If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                                    ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                        Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                        Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                                    Else 'sinon le format n'est pas du type variable
                                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                        Case "CONTACT_TEL_FROM"
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                                If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                                    ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                            Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString.Replace(".", "").Replace(" ", "").Replace("-", ""), PositionLeft)
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
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                        Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                        Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                                    Else 'sinon le format n'est pas du type variable
                                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                        Case "CONTACT_EXT_FROM"
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                                If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                                    ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                        Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                        Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                                    Else 'sinon le format n'est pas du type variable
                                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                        Case "CONTACT_FAX_FROM"
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                                If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                                    ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                            Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString.Replace(".", "").Replace(" ", "").Replace("-", ""), PositionLeft)
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
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                        Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                        Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                                    Else 'sinon le format n'est pas du type variable
                                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                        Case "COMMENT_FROM"
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                                If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                                    ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                        Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                        Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                                    Else 'sinon le format n'est pas du type variable
                                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                        Case "ADDR_NAME_TO"
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> Nothing Then ' si valeur Sage n'est pas null
                                                    ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                            Dim OleAdaptaterschemaSageDepot As OleDbDataAdapter = Nothing
                                                            Dim OleSchemaDatasetSageDepot As DataSet = New DataSet
                                                            Dim OledatableSchemaSageDepot As DataTable
                                                            OleAdaptaterschemaSageDepot = New OleDbDataAdapter("SELECT DE_Intitule FROM F_COMPTET INNER JOIN F_DEPOT ON F_COMPTET.DE_No =F_DEPOT .DE_No AND F_COMPTET.CT_Num ='" & OledatableSchemaSage.Rows(i).Item("CT_NUM").ToString & "'", OleExcelConnectedFrs)
                                                            OleAdaptaterschemaSageDepot.Fill(OleSchemaDatasetSageDepot)
                                                            OledatableSchemaSageDepot = OleSchemaDatasetSageDepot.Tables(0)
                                                            If OledatableSchemaSageDepot.Rows.Count <> 0 Then
                                                                Information &= Strings.LSet(OledatableSchemaSageDepot.Rows(0).Item("DE_Intitule").ToString, PositionLeft)
                                                            Else
                                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
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
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                        Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                        Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                                    Else 'sinon le format n'est pas du type variable
                                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                        Case "ADDR_LINE1_TO"
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> Nothing Then ' si valeur Sage n'est pas null
                                                    ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                            Dim OleAdaptaterschemaSageDepot As OleDbDataAdapter = Nothing
                                                            Dim OleSchemaDatasetSageDepot As DataSet = New DataSet
                                                            Dim OledatableSchemaSageDepot As DataTable
                                                            OleAdaptaterschemaSageDepot = New OleDbDataAdapter("SELECT DE_Adresse FROM F_COMPTET INNER JOIN F_DEPOT ON F_COMPTET.DE_No =F_DEPOT .DE_No AND F_COMPTET.CT_Num ='" & OledatableSchemaSage.Rows(i).Item("CT_NUM").ToString & "'", OleExcelConnectedFrs)
                                                            OleAdaptaterschemaSageDepot.Fill(OleSchemaDatasetSageDepot)
                                                            OledatableSchemaSageDepot = OleSchemaDatasetSageDepot.Tables(0)
                                                            If OledatableSchemaSageDepot.Rows.Count <> 0 Then
                                                                Information &= Strings.LSet(OledatableSchemaSageDepot.Rows(0).Item("DE_Adresse").ToString, PositionLeft)
                                                            Else
                                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
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
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                        Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                        Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                                    Else 'sinon le format n'est pas du type variable
                                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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

                                        Case "ADDR_LINE2_TO"
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> Nothing Then ' si valeur Sage n'est pas null
                                                    ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                            Dim OleAdaptaterschemaSageDepot As OleDbDataAdapter = Nothing
                                                            Dim OleSchemaDatasetSageDepot As DataSet = New DataSet
                                                            Dim OledatableSchemaSageDepot As DataTable
                                                            OleAdaptaterschemaSageDepot = New OleDbDataAdapter("SELECT DE_Complement FROM F_COMPTET INNER JOIN F_DEPOT ON F_COMPTET.DE_No =F_DEPOT .DE_No AND F_COMPTET.CT_Num ='" & OledatableSchemaSage.Rows(i).Item("CT_NUM").ToString & "'", OleExcelConnectedFrs)
                                                            OleAdaptaterschemaSageDepot.Fill(OleSchemaDatasetSageDepot)
                                                            OledatableSchemaSageDepot = OleSchemaDatasetSageDepot.Tables(0)
                                                            If OledatableSchemaSageDepot.Rows.Count <> 0 Then
                                                                Information &= Strings.LSet(OledatableSchemaSageDepot.Rows(0).Item("DE_Complement").ToString, PositionLeft)
                                                            Else
                                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
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
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                        Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                        Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                                    Else 'sinon le format n'est pas du type variable
                                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                        Case "CITY_TO"
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> Nothing Then ' si valeur Sage n'est pas null
                                                    ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                            Dim OleAdaptaterschemaSageDepot As OleDbDataAdapter = Nothing
                                                            Dim OleSchemaDatasetSageDepot As DataSet = New DataSet
                                                            Dim OledatableSchemaSageDepot As DataTable
                                                            OleAdaptaterschemaSageDepot = New OleDbDataAdapter("SELECT DE_Ville FROM F_COMPTET INNER JOIN F_DEPOT ON F_COMPTET.DE_No =F_DEPOT .DE_No AND F_COMPTET.CT_Num ='" & OledatableSchemaSage.Rows(i).Item("CT_NUM").ToString & "'", OleExcelConnectedFrs)
                                                            OleAdaptaterschemaSageDepot.Fill(OleSchemaDatasetSageDepot)
                                                            OledatableSchemaSageDepot = OleSchemaDatasetSageDepot.Tables(0)
                                                            If OledatableSchemaSageDepot.Rows.Count <> 0 Then
                                                                Information &= Strings.LSet(OledatableSchemaSageDepot.Rows(0).Item("DE_Ville").ToString, PositionLeft)
                                                            Else
                                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
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
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                        Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                        Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                                    Else 'sinon le format n'est pas du type variable
                                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                        Case "STATE_TO"
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> Nothing Then ' si valeur Sage n'est pas null
                                                    ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                            Dim OleAdaptaterschemaSageDepot As OleDbDataAdapter = Nothing
                                                            Dim OleSchemaDatasetSageDepot As DataSet = New DataSet
                                                            Dim OledatableSchemaSageDepot As DataTable
                                                            OleAdaptaterschemaSageDepot = New OleDbDataAdapter("SELECT DE_Region FROM F_COMPTET INNER JOIN F_DEPOT ON F_COMPTET.DE_No =F_DEPOT .DE_No AND F_COMPTET.CT_Num ='" & OledatableSchemaSage.Rows(i).Item("CT_NUM").ToString & "'", OleExcelConnectedFrs)
                                                            OleAdaptaterschemaSageDepot.Fill(OleSchemaDatasetSageDepot)
                                                            OledatableSchemaSageDepot = OleSchemaDatasetSageDepot.Tables(0)
                                                            If OledatableSchemaSageDepot.Rows.Count <> 0 Then
                                                                Information &= Strings.LSet(OledatableSchemaSageDepot.Rows(0).Item("DE_Region").ToString, PositionLeft)
                                                            Else
                                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
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
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                        Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                        Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                                    Else 'sinon le format n'est pas du type variable
                                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                        Case "COUNTRY_TO"
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> Nothing Then ' si valeur Sage n'est pas null
                                                    ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                            Dim OleAdaptaterschemaSageDepot As OleDbDataAdapter = Nothing
                                                            Dim OleSchemaDatasetSageDepot As DataSet = New DataSet
                                                            Dim OledatableSchemaSageDepot As DataTable
                                                            OleAdaptaterschemaSageDepot = New OleDbDataAdapter("SELECT DE_Pays FROM F_COMPTET INNER JOIN F_DEPOT ON F_COMPTET.DE_No =F_DEPOT .DE_No AND F_COMPTET.CT_Num ='" & OledatableSchemaSage.Rows(i).Item("CT_NUM").ToString & "'", OleExcelConnectedFrs)
                                                            OleAdaptaterschemaSageDepot.Fill(OleSchemaDatasetSageDepot)
                                                            OledatableSchemaSageDepot = OleSchemaDatasetSageDepot.Tables(0)
                                                            If OledatableSchemaSageDepot.Rows.Count <> 0 Then
                                                                Information &= Strings.LSet(OledatableSchemaSageDepot.Rows(0).Item("DE_Pays").ToString, PositionLeft)
                                                            Else
                                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
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
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                        Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                        Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                                    Else 'sinon le format n'est pas du type variable
                                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                        Case "ZIP_TO"
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                                If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> Nothing Then ' si valeur Sage n'est pas null
                                                    ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                            Dim OleAdaptaterschemaSageDepot As OleDbDataAdapter = Nothing
                                                            Dim OleSchemaDatasetSageDepot As DataSet = New DataSet
                                                            Dim OledatableSchemaSageDepot As DataTable
                                                            OleAdaptaterschemaSageDepot = New OleDbDataAdapter("SELECT DE_CodePostal FROM F_COMPTET INNER JOIN F_DEPOT ON F_COMPTET.DE_No =F_DEPOT .DE_No AND F_COMPTET.CT_Num ='" & OledatableSchemaSage.Rows(i).Item("CT_NUM").ToString & "'", OleExcelConnectedFrs)
                                                            OleAdaptaterschemaSageDepot.Fill(OleSchemaDatasetSageDepot)
                                                            OledatableSchemaSageDepot = OleSchemaDatasetSageDepot.Tables(0)
                                                            If OledatableSchemaSageDepot.Rows.Count <> 0 Then
                                                                Information &= Strings.LSet(OledatableSchemaSageDepot.Rows(0).Item("DE_CodePostal").ToString, PositionLeft)
                                                            Else
                                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
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
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                        Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                        Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                                    Else 'sinon le format n'est pas du type variable
                                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                                            Information &= Strings.LSet("", PositionLeft)
                                                        End If
                                                    End If
                                                End If
                                            End If
                                            Continue For
                                        Case "CONTACT_NAME_TO"
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                                If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                                    ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                        Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                        Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                                    Else 'sinon le format n'est pas du type variable
                                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                        Case "CONTACT_TEL_TO"
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                                If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                                    ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                        Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                        Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                                    Else 'sinon le format n'est pas du type variable
                                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                        Case "CONTACT_EXT_TO"
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                                If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                                    ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                        Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                        Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                                    Else 'sinon le format n'est pas du type variable
                                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                        Case "CONTACT_FAX_TO"
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                                If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                                    ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                        Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                        Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                                    Else 'sinon le format n'est pas du type variable
                                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                        Case "COMMENT_TO"
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                                If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                                    ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                        Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                        Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                                    Else 'sinon le format n'est pas du type variable
                                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                            PositionLeft = OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(1).ToString.Split(")")(0).Split(".")(0)
                                            OleAdaptaterschemaFourssAR = New OleDbDataAdapter("SELECT P_EXPEDITION.E_Intitule  FROM F_COMPTET INNER JOIN P_EXPEDITION  ON F_COMPTET.CT_Num='" & OledatableSchemaSage.Rows(i).Item("CT_NUM").ToString & "' AND F_COMPTET.cbMarq =P_EXPEDITION.cbMarq  AND F_COMPTET.CT_Type =1", OleExcelConnectedFrs)
                                            OleSchemaDatasetFourssAR = New DataSet
                                            OleAdaptaterschemaFourssAR.Fill(OleSchemaDatasetFourssAR)
                                            OledatableSchemaFourssAR = OleSchemaDatasetFourssAR.Tables(0)
                                            If OledatableSchemaFourssAR.Rows.Count > 0 Then
                                                If OledatableSchemaFourssAR.Rows(0).Item("E_Intitule") = "" Then
                                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                    Else
                                                        Information &= Strings.LSet("", PositionLeft)
                                                    End If
                                                Else
                                                    Information &= Strings.LSet(OledatableSchemaFourssAR.Rows(0).Item("E_Intitule"), PositionLeft)
                                                End If
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                Else
                                                    Information &= Strings.LSet("", PositionLeft)
                                                End If
                                            End If
                                            Continue For
                                        Case "CARRIER_NAME"
                                            PositionLeft = OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(1).ToString.Split(")")(0).Split(".")(0)
                                            OleAdaptaterschemaFourssAR = New OleDbDataAdapter("SELECT P_EXPEDITION.E_Intitule  FROM F_COMPTET INNER JOIN P_EXPEDITION  ON F_COMPTET.CT_Num='" & OledatableSchemaSage.Rows(i).Item("CT_NUM").ToString & "' AND F_COMPTET.cbMarq =P_EXPEDITION.cbMarq  AND F_COMPTET.CT_Type =1", OleExcelConnectedFrs)
                                            OleSchemaDatasetFourssAR = New DataSet
                                            OleAdaptaterschemaFourssAR.Fill(OleSchemaDatasetFourssAR)
                                            OledatableSchemaFourssAR = OleSchemaDatasetFourssAR.Tables(0)

                                            If OledatableSchemaFourssAR.Rows.Count > 0 Then
                                                If OledatableSchemaFourssAR.Rows(0).Item("E_Intitule") = "" Then
                                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                    Else
                                                        Information &= Strings.LSet("", PositionLeft)
                                                    End If
                                                Else
                                                    Information &= Strings.LSet(OledatableSchemaFourssAR.Rows(0).Item("E_Intitule"), PositionLeft)
                                                End If
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                Else
                                                    Information &= Strings.LSet("", PositionLeft)
                                                End If
                                            End If
                                            Continue For
                                        Case "DESC_CARRIER"
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                                If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                                    ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                        Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                        Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                                    Else 'sinon le format n'est pas du type variable
                                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                        Case "TERMS"
                                            PositionLeft = OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(1).ToString.Split(")")(0).Split(".")(0)
                                            OleAdaptaterschemaFourssAR = New OleDbDataAdapter("SELECT P_CONDLIVR.C_Intitule from F_COMPTET INNER JOIN P_CONDLIVR ON F_COMPTET.N_condition =P_CONDLIVR .cbMarq AND F_COMPTET.CT_Num='" & OledatableSchemaSage.Rows(i).Item("CT_NUM").ToString & "'", OleExcelConnectedFrs)
                                            OleSchemaDatasetFourssAR = New DataSet
                                            OleAdaptaterschemaFourssAR.Fill(OleSchemaDatasetFourssAR)
                                            OledatableSchemaFourssAR = OleSchemaDatasetFourssAR.Tables(0)

                                            If OledatableSchemaFourssAR.Rows.Count > 0 Then
                                                If OledatableSchemaFourssAR.Rows(0).Item("C_Intitule") = "" Then
                                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                    Else
                                                        Information &= Strings.LSet("", PositionLeft)
                                                    End If
                                                Else
                                                    Information &= Strings.LSet(OledatableSchemaFourssAR.Rows(0).Item("C_Intitule"), PositionLeft)
                                                End If
                                            Else
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                Else
                                                    Information &= Strings.LSet("", PositionLeft)
                                                End If
                                            End If
                                            Continue For
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                                If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                                    ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                            Dim ValeurLue As String = ""
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
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                        Dim ValeurLue As String = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                        Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                                    Else 'sinon le format n'est pas du type variable
                                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
                                                        If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                            Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                        Else
                                                            Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                        End If
                                                    End If
                                                Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                                    'Dim traitementAutres As String = 0
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                        Dim ValeurLue As String = ""
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
                                        Case "DAMAGED_OWNER_CODE"
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                                If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                                    ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                        Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                        Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                                    Else 'sinon le format n'est pas du type variable
                                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                        Case Else
                                            If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                                If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                                    ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                        Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                        Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                                    Else 'sinon le format n'est pas du type variable
                                                        PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                    '-------------------------------------**********FIN********************------------------------------------------
                                Next
                                statut = True
                                If statut = True Then
                                    Error_journalFrs.WriteLine(Information)
                                    Information = ""
                                End If
                            Catch ex As Exception
                                MsgBox(ex.Message)
                            End Try
                        End If
                    Else
                        If Ckmodifier.Checked = False And statut = False Then
                            Error_journalFrs = File.AppendText(PathsFileFormatiers & "SUP" & version & Strings.Right(DateAndTime.Year(Now), 4) & Format(DateAndTime.Month(Now), "00") & Format(DateAndTime.Day(Now), "00") & Format(DateAndTime.Hour(Now), "00") & Format(DateAndTime.Minute(Now), "00") & Format(DateAndTime.Second(Now), "00") & ".txt")
                            statut = True
                        End If
                        Element = OledatableSchemaSage.Rows(i).Item("CT_NUM")
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
                                Case "OP�RATION"
                                    If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                    Else
                                        Information &= Strings.LSet("F", PositionLeft)
                                    End If
                                    Continue For
                                Case "SUPPLIER_CODE"
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                        If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                            ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else 'sinon le format n'est pas du type variable
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                Case "DESCRIPTION"
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                        If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                            ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else 'sinon le format n'est pas du type variable
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                Case "SUPPLIERCLASS_CODE"
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                        If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                            ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else 'sinon le format n'est pas du type variable
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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

                                Case "ADDR_NAME_FROM"
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                        If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                            ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else 'sinon le format n'est pas du type variable
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                Case "ADDR_LINE1_FROM"
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                        If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                            ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else 'sinon le format n'est pas du type variable
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                Case "ADDR_LINE2_FROM"
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                        If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                            ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else 'sinon le format n'est pas du type variable
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                                    Information &= Strings.LSet("", PositionLeft)
                                                End If
                                            End If
                                        End If
                                    End If
                                    Continue For
                                Case "CITY_FROM"
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                        If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                            ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else 'sinon le format n'est pas du type variable
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                Case "STATE_FROM"
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                        If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                            ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else 'sinon le format n'est pas du type variable
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                Case "COUNTRY_FROM"
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                        If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                            ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else 'sinon le format n'est pas du type variable
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                Case "ZIP_FROM"
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                        If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                            ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else 'sinon le format n'est pas du type variable
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                Case "CONTACT_NAME_FROM"
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                        If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                            ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else 'sinon le format n'est pas du type variable
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                Case "CONTACT_TEL_FROM"
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                        If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                            ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                    Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString.Replace(".", "").Replace(" ", "").Replace("-", ""), PositionLeft)
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else 'sinon le format n'est pas du type variable
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                Case "CONTACT_EXT_FROM"
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                        If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                            ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else 'sinon le format n'est pas du type variable
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                Case "CONTACT_FAX_FROM"
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                        If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                            ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                    Information &= Strings.LSet(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString.Replace(".", "").Replace(" ", "").Replace("-", ""), PositionLeft)
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else 'sinon le format n'est pas du type variable
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                Case "COMMENT_FROM"
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                        If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                            ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else 'sinon le format n'est pas du type variable
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                Case "ADDR_NAME_TO"
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                        If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> Nothing Then ' si valeur Sage n'est pas null
                                            ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                    Dim OleAdaptaterschemaSageDepot As OleDbDataAdapter = Nothing
                                                    Dim OleSchemaDatasetSageDepot As DataSet = New DataSet
                                                    Dim OledatableSchemaSageDepot As DataTable
                                                    OleAdaptaterschemaSageDepot = New OleDbDataAdapter("SELECT DE_Intitule FROM F_COMPTET INNER JOIN F_DEPOT ON F_COMPTET.DE_No =F_DEPOT .DE_No AND F_COMPTET.CT_Num ='" & OledatableSchemaSage.Rows(i).Item("CT_NUM").ToString & "'", OleExcelConnectedFrs)
                                                    OleAdaptaterschemaSageDepot.Fill(OleSchemaDatasetSageDepot)
                                                    OledatableSchemaSageDepot = OleSchemaDatasetSageDepot.Tables(0)
                                                    If OledatableSchemaSageDepot.Rows.Count <> 0 Then
                                                        Information &= Strings.LSet(OledatableSchemaSageDepot.Rows(0).Item("DE_Intitule").ToString, PositionLeft)
                                                    Else
                                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else 'sinon le format n'est pas du type variable
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                Case "ADDR_LINE1_TO"
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                        If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> Nothing Then ' si valeur Sage n'est pas null
                                            ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                    Dim OleAdaptaterschemaSageDepot As OleDbDataAdapter = Nothing
                                                    Dim OleSchemaDatasetSageDepot As DataSet = New DataSet
                                                    Dim OledatableSchemaSageDepot As DataTable
                                                    OleAdaptaterschemaSageDepot = New OleDbDataAdapter("SELECT DE_Adresse FROM F_COMPTET INNER JOIN F_DEPOT ON F_COMPTET.DE_No =F_DEPOT .DE_No AND F_COMPTET.CT_Num ='" & OledatableSchemaSage.Rows(i).Item("CT_NUM").ToString & "'", OleExcelConnectedFrs)
                                                    OleAdaptaterschemaSageDepot.Fill(OleSchemaDatasetSageDepot)
                                                    OledatableSchemaSageDepot = OleSchemaDatasetSageDepot.Tables(0)
                                                    If OledatableSchemaSageDepot.Rows.Count <> 0 Then
                                                        Information &= Strings.LSet(OledatableSchemaSageDepot.Rows(0).Item("DE_Adresse").ToString, PositionLeft)
                                                    Else
                                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else 'sinon le format n'est pas du type variable
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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

                                Case "ADDR_LINE2_TO"
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                        If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> Nothing Then ' si valeur Sage n'est pas null
                                            ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                    If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                        Information &= Strings.LSet(RenvoiValeurLue(OledatableSchemaSage.Rows(i).Item(ColonnName).ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                    Else
                                                        Dim OleAdaptaterschemaSageDepot As OleDbDataAdapter = Nothing
                                                        Dim OleSchemaDatasetSageDepot As DataSet = New DataSet
                                                        Dim OledatableSchemaSageDepot As DataTable
                                                        OleAdaptaterschemaSageDepot = New OleDbDataAdapter("SELECT DE_Complement FROM F_COMPTET INNER JOIN F_DEPOT ON F_COMPTET.DE_No =F_DEPOT .DE_No AND F_COMPTET.CT_Num ='" & OledatableSchemaSage.Rows(i).Item("CT_NUM").ToString & "'", OleExcelConnectedFrs)
                                                        OleAdaptaterschemaSageDepot.Fill(OleSchemaDatasetSageDepot)
                                                        OledatableSchemaSageDepot = OleSchemaDatasetSageDepot.Tables(0)
                                                        If OledatableSchemaSageDepot.Rows.Count <> 0 Then
                                                            Information &= Strings.LSet(OledatableSchemaSageDepot.Rows(0).Item("DE_Complement").ToString, PositionLeft)
                                                        Else
                                                            Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else 'sinon le format n'est pas du type variable
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                Case "CITY_TO"
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                        If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> Nothing Then ' si valeur Sage n'est pas null
                                            ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                    Dim OleAdaptaterschemaSageDepot As OleDbDataAdapter = Nothing
                                                    Dim OleSchemaDatasetSageDepot As DataSet = New DataSet
                                                    Dim OledatableSchemaSageDepot As DataTable
                                                    OleAdaptaterschemaSageDepot = New OleDbDataAdapter("SELECT DE_Ville FROM F_COMPTET INNER JOIN F_DEPOT ON F_COMPTET.DE_No =F_DEPOT .DE_No AND F_COMPTET.CT_Num ='" & OledatableSchemaSage.Rows(i).Item("CT_NUM").ToString & "'", OleExcelConnectedFrs)
                                                    OleAdaptaterschemaSageDepot.Fill(OleSchemaDatasetSageDepot)
                                                    OledatableSchemaSageDepot = OleSchemaDatasetSageDepot.Tables(0)
                                                    If OledatableSchemaSageDepot.Rows.Count <> 0 Then
                                                        Information &= Strings.LSet(OledatableSchemaSageDepot.Rows(0).Item("DE_Ville").ToString, PositionLeft)
                                                    Else
                                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else 'sinon le format n'est pas du type variable
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                Case "STATE_TO"
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                        If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> Nothing Then ' si valeur Sage n'est pas null
                                            ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                    Dim OleAdaptaterschemaSageDepot As OleDbDataAdapter = Nothing
                                                    Dim OleSchemaDatasetSageDepot As DataSet = New DataSet
                                                    Dim OledatableSchemaSageDepot As DataTable
                                                    OleAdaptaterschemaSageDepot = New OleDbDataAdapter("SELECT DE_Region FROM F_COMPTET INNER JOIN F_DEPOT ON F_COMPTET.DE_No =F_DEPOT .DE_No AND F_COMPTET.CT_Num ='" & OledatableSchemaSage.Rows(i).Item("CT_NUM").ToString & "'", OleExcelConnectedFrs)
                                                    OleAdaptaterschemaSageDepot.Fill(OleSchemaDatasetSageDepot)
                                                    OledatableSchemaSageDepot = OleSchemaDatasetSageDepot.Tables(0)
                                                    If OledatableSchemaSageDepot.Rows.Count <> 0 Then
                                                        Information &= Strings.LSet(OledatableSchemaSageDepot.Rows(0).Item("DE_Region").ToString, PositionLeft)
                                                    Else
                                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else 'sinon le format n'est pas du type variable
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                Case "COUNTRY_TO"
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                        If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> Nothing Then ' si valeur Sage n'est pas null
                                            ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                    Dim OleAdaptaterschemaSageDepot As OleDbDataAdapter = Nothing
                                                    Dim OleSchemaDatasetSageDepot As DataSet = New DataSet
                                                    Dim OledatableSchemaSageDepot As DataTable
                                                    OleAdaptaterschemaSageDepot = New OleDbDataAdapter("SELECT DE_Pays FROM F_COMPTET INNER JOIN F_DEPOT ON F_COMPTET.DE_No =F_DEPOT .DE_No AND F_COMPTET.CT_Num ='" & OledatableSchemaSage.Rows(i).Item("CT_NUM").ToString & "'", OleExcelConnectedFrs)
                                                    OleAdaptaterschemaSageDepot.Fill(OleSchemaDatasetSageDepot)
                                                    OledatableSchemaSageDepot = OleSchemaDatasetSageDepot.Tables(0)
                                                    If OledatableSchemaSageDepot.Rows.Count <> 0 Then
                                                        Information &= Strings.LSet(OledatableSchemaSageDepot.Rows(0).Item("DE_Pays").ToString, PositionLeft)
                                                    Else
                                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else 'sinon le format n'est pas du type variable
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                Case "ZIP_TO"
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                        If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> Nothing Then ' si valeur Sage n'est pas null
                                            ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                    Dim OleAdaptaterschemaSageDepot As OleDbDataAdapter = Nothing
                                                    Dim OleSchemaDatasetSageDepot As DataSet = New DataSet
                                                    Dim OledatableSchemaSageDepot As DataTable
                                                    OleAdaptaterschemaSageDepot = New OleDbDataAdapter("SELECT DE_CodePostal FROM F_COMPTET INNER JOIN F_DEPOT ON F_COMPTET.DE_No =F_DEPOT .DE_No AND F_COMPTET.CT_Num ='" & OledatableSchemaSage.Rows(i).Item("CT_NUM").ToString & "'", OleExcelConnectedFrs)
                                                    OleAdaptaterschemaSageDepot.Fill(OleSchemaDatasetSageDepot)
                                                    OledatableSchemaSageDepot = OleSchemaDatasetSageDepot.Tables(0)
                                                    If OledatableSchemaSageDepot.Rows.Count <> 0 Then
                                                        Information &= Strings.LSet(OledatableSchemaSageDepot.Rows(0).Item("DE_CodePostal").ToString, PositionLeft)
                                                    Else
                                                        Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else 'sinon le format n'est pas du type variable
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                                    Information &= Strings.LSet("", PositionLeft)
                                                End If
                                            End If
                                        End If
                                    End If
                                    Continue For
                                Case "CONTACT_NAME_TO"
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                        If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                            ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else 'sinon le format n'est pas du type variable
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                Case "CONTACT_TEL_TO"
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                        If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                            ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else 'sinon le format n'est pas du type variable
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                Case "CONTACT_EXT_TO"
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                        If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                            ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else 'sinon le format n'est pas du type variable
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                Case "CONTACT_FAX_TO"
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                        If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                            ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else 'sinon le format n'est pas du type variable
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                Case "COMMENT_TO"
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                        If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                            ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else 'sinon le format n'est pas du type variable
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                    PositionLeft = OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(1).ToString.Split(")")(0).Split(".")(0)
                                    OleAdaptaterschemaFourssAR = New OleDbDataAdapter("SELECT P_EXPEDITION.E_Intitule  FROM F_COMPTET INNER JOIN P_EXPEDITION  ON F_COMPTET.CT_Num='" & OledatableSchemaSage.Rows(i).Item("CT_NUM").ToString & "' AND F_COMPTET.cbMarq =P_EXPEDITION.cbMarq  AND F_COMPTET.CT_Type =1", OleExcelConnectedFrs)
                                    OleSchemaDatasetFourssAR = New DataSet
                                    OleAdaptaterschemaFourssAR.Fill(OleSchemaDatasetFourssAR)
                                    OledatableSchemaFourssAR = OleSchemaDatasetFourssAR.Tables(0)
                                    If OledatableSchemaFourssAR.Rows.Count > 0 Then
                                        If OledatableSchemaFourssAR.Rows(0).Item("E_Intitule") = "" Then
                                            If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            Else
                                                Information &= Strings.LSet("", PositionLeft)
                                            End If
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaFourssAR.Rows(0).Item("E_Intitule"), PositionLeft)
                                        End If
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                        Else
                                            Information &= Strings.LSet("", PositionLeft)
                                        End If
                                    End If
                                    Continue For
                                Case "CARRIER_NAME"
                                    PositionLeft = OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(1).ToString.Split(")")(0).Split(".")(0)
                                    OleAdaptaterschemaFourssAR = New OleDbDataAdapter("SELECT P_EXPEDITION.E_Intitule  FROM F_COMPTET INNER JOIN P_EXPEDITION  ON F_COMPTET.CT_Num='" & OledatableSchemaSage.Rows(i).Item("CT_NUM").ToString & "' AND F_COMPTET.cbMarq =P_EXPEDITION.cbMarq  AND F_COMPTET.CT_Type =1", OleExcelConnectedFrs)
                                    OleSchemaDatasetFourssAR = New DataSet
                                    OleAdaptaterschemaFourssAR.Fill(OleSchemaDatasetFourssAR)
                                    OledatableSchemaFourssAR = OleSchemaDatasetFourssAR.Tables(0)

                                    If OledatableSchemaFourssAR.Rows.Count > 0 Then
                                        If OledatableSchemaFourssAR.Rows(0).Item("E_Intitule") = "" Then
                                            If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            Else
                                                Information &= Strings.LSet("", PositionLeft)
                                            End If
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaFourssAR.Rows(0).Item("E_Intitule"), PositionLeft)
                                        End If
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                        Else
                                            Information &= Strings.LSet("", PositionLeft)
                                        End If
                                    End If
                                    Continue For
                                Case "DESC_CARRIER"
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                        If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                            ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else 'sinon le format n'est pas du type variable
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                Case "TERMS"
                                    PositionLeft = OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(1).ToString.Split(")")(0).Split(".")(0)
                                    OleAdaptaterschemaFourssAR = New OleDbDataAdapter("SELECT P_CONDLIVR.C_Intitule from F_COMPTET INNER JOIN P_CONDLIVR ON F_COMPTET.N_condition =P_CONDLIVR .cbMarq AND F_COMPTET.CT_Num='" & OledatableSchemaSage.Rows(i).Item("CT_NUM").ToString & "'", OleExcelConnectedFrs)
                                    OleSchemaDatasetFourssAR = New DataSet
                                    OleAdaptaterschemaFourssAR.Fill(OleSchemaDatasetFourssAR)
                                    OledatableSchemaFourssAR = OleSchemaDatasetFourssAR.Tables(0)

                                    If OledatableSchemaFourssAR.Rows.Count > 0 Then
                                        If OledatableSchemaFourssAR.Rows(0).Item("C_Intitule") = "" Then
                                            If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                                Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                            Else
                                                Information &= Strings.LSet("", PositionLeft)
                                            End If
                                        Else
                                            Information &= Strings.LSet(OledatableSchemaFourssAR.Rows(0).Item("C_Intitule"), PositionLeft)
                                        End If
                                    Else
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then
                                            Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                        Else
                                            Information &= Strings.LSet("", PositionLeft)
                                        End If
                                    End If
                                    Continue For
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                        If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                            ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                                    Dim ValeurLue As String = ""
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                Dim ValeurLue As String = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else 'sinon le format n'est pas du type variable
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
                                                If OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "N" Or OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(0) = "D" Then
                                                    Information &= Strings.LSet(RenvoiValeurLue(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                                Else
                                                    Information &= Strings.LSet(OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString, PositionLeft)
                                                End If
                                            End If
                                        Else 'Prevue pour la non existance de mapping Sage et  de valeur pas defaut
                                            'Dim traitementAutres As String = 0
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then
                                                Dim ValeurLue As String = ""
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
                                Case "DAMAGED_OWNER_CODE"
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                        If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                            ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else 'sinon le format n'est pas du type variable
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                                Case Else
                                    If OledatableSchema.Rows(Ligne).Item("ChampSage").ToString <> "" Then 's'il existe un mappig Sage 
                                        If OledatableSchemaSage.Rows(i).Item(OledatableSchema.Rows(Ligne).Item("ChampSage").ToString).ToString <> "" Then ' si valeur Sage n'est pas null
                                            ColonnName = OledatableSchema.Rows(Ligne).Item("ChampSage").ToString ' affecter de la colonne dans � la variable colonne
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
                                        If OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString <> "" Then 's'il existe une valeur pas defaut param�tr�e
                                            If OledatableSchema.Rows(Ligne).Item("Format").ToString = "V" Then 'si le format lue est du type Variable(V) si oui
                                                Dim ValeurLue = OledatableSchema.Rows(Ligne).Item("DefaultValue").ToString
                                                Information &= Strings.LSet("[" & ValeurLue & "]", ValeurLue.Length + 2)
                                            Else 'sinon le format n'est pas du type variable
                                                PositionLeft = GetPosition(OledatableSchema.Rows(Ligne).Item("Format").ToString) 'recuperation de la longueur du Champ grace � la fonction GetPosition qui prend en param�tre le format du champ
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
                            '-------------------------------------**********FIN********************------------------------------------------
                        Next
                        statut = True
                        If statut = True Then
                            Error_journalFrs.WriteLine(Information)
                            Information = ""
                        End If
                    End If
                Next
                If statut = True Then
                    Error_journalFrs.Close()
                End If
                If Ckmodifier.Checked Then
                    IDernierDate(3)
                    statut = False
                End If
            Else
                'cette requette ne possede pas ligne 
                EstVide = False
            End If
            Return True
        Catch ex As Exception
            'MsgBox(ex.Message)
            Return False
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
                lblInfos.Text = "Aucune donn�e retourn�e par ce Filtre"
                lblInfos.Visible = True
            Else
                If ifRowError > 0 Then
                    DataListeIntegrer.Rows(selectids2(ip)).Cells("Status").Value = My.Resources.criticalind_status
                    ifRowError = 0
                Else
                    DataListeIntegrer.Rows(selectids2(ip)).Cells("Status").Value = My.Resources.accepter
                End If
            End If
            If ip < selectids2.Length - 1 Then
                ip = ip + 1
                BackgroundWorker1.RunWorkerAsync()
            Else
                ip = 0
                nbreligne = 0
                selectindexe2 = ""
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

    Private Sub BackgroundWorker4_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker4.DoWork
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
End Class