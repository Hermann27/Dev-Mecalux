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

    Private Sub BtnModif_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnModif.Click
        ListBox.Items.Clear()
        ContinuTraitement = 0
        NbInfosLibre = 0
        NbInfosLibreVue = 0
        Try
            VerificateurChampObligatoire(True, False)
            If ContinuTraitement = 4 Then
                EstInfosLibre(True, False, True)
                If NbInfosLibre = NbInfosLibreVue Then 'NbInfosLibreVue
                    Try

                        PictureBox1.Visible = True
                        ListeChampSageEntete = RecuperationColonneSage(True, False)
                        lblSne.Text = "Scénario Extraction des Fournisseurs"
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
        OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='SUP' AND Entete=" & Entete & " AND Ligne=" & Ligne, OleConnenection)
        OleSchemaDataset = New DataSet
        OleAdaptaterschema.Fill(OleSchemaDataset)
        OledatableSchema = OleSchemaDataset.Tables(0)
        lblSne.Text = "Verification conformite des fournisseurs"
        ListBox.Items.Add("<-----------------------------------------------------------------DEBUT DU TRAITEMENT DE LA VERIFICATION DES CHAMPS OBLIGATOIRES----------------------------------------------------------------->")
        If OledatableSchema.Rows.Count <> 0 Then
            For i As Integer = 0 To OledatableSchema.Rows.Count - 1
                Select Case OledatableSchema.Rows(i).Item("Cols")
                    Case "OPÉRATION"
                        If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False And OledatableSchema.Rows(i).Item("ChampSage") <> "" Then
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
                    Case "SUPPLIER_CODE"
                        If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False And OledatableSchema.Rows(i).Item("ChampSage") <> "" Then
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
                    Case "ADDR_NAME_FROM"
                        If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False And OledatableSchema.Rows(i).Item("ChampSage") <> "" Then
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
                    Case "ADDR_NAME_TO"
                        If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False And OledatableSchema.Rows(i).Item("ChampSage") <> "" Then
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
                                ListBox.Items.Add("<--Le Champ indiquant l'infos libre est couché mais ne possede pas de mapping Sage-->")
                            End If
                        ElseIf Ligne = True Then
                            If IsDBNull(OledatableSchema.Rows(i).Item("ChampSage")) = False Then
                                If ExisteInfosLibre(OledatableSchema.Rows(i).Item("ChampSage")) = False Then
                                    ListBox.Items.Add("la zone infos libre {" & OledatableSchema.Rows(i).Item("ChampSage") & "} sur le fournisseur n'existe pas dans Sage")
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
                OleAdaptaterschemaSage = New OleDbDataAdapter("select * from cbSysLibre WHERE CB_File='F_COMPTET' And CB_Name='" & Join(Split(Trim(InfosLibre), "'"), "''") & "'", OleExcelConnect)
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

    Public Sub ExtractionFournisseur(ByVal SelectChamp As String)
        Dim Information As String = ""
        Dim ExistColonne As Boolean = False
        Dim ColonnName As String = ""
        Dim PositionLeft As Object = Nothing

        Try
            If ExcelConnect(NomServersql, NomBaseCpta, Nom_Utilsql, Mot_Passql) Then
                If RbARcreation.IsChecked Then
                    If SelectChamp IsNot Nothing Then
                        OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & " FROM F_COMPTET WHERE CT_Type=1 AND CT_DateCreate = CBModification ORDER BY F_COMPTET.CT_NUM ", OleExcelConnect)
                    Else
                        OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_COMPTET WHERE CT_Type=1 AND CT_DateCreate = CBModification ORDER BY F_COMPTET.CT_NUM ", OleExcelConnect)
                    End If
                ElseIf RbARmodifier.IsChecked Then
                    If SelectChamp IsNot Nothing Then
                        OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & " FROM F_COMPTET WHERE CT_Type=1 AND CT_DateCreate <> CBModification ORDER BY F_COMPTET.CT_NUM ", OleExcelConnect)
                    Else
                        OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_COMPTET WHERE CT_Type=1 AND CT_DateCreate <> CBModification ORDER BY F_COMPTET.CT_NUM ", OleExcelConnect)
                    End If
                ElseIf (RbAR2action.IsChecked) Then
                    If SelectChamp IsNot Nothing Then
                        OleAdaptaterschemaSage = New OleDbDataAdapter("select " & SelectChamp & " FROM F_COMPTET WHERE CT_Type=1 ORDER BY F_COMPTET.CT_NUM", OleExcelConnect)
                    Else
                        OleAdaptaterschemaSage = New OleDbDataAdapter("select * FROM F_COMPTET WHERE CT_Type=1 ORDER BY F_COMPTET.CT_NUM", OleExcelConnect)
                    End If
                End If

                OleSchemaDatasetSage = New DataSet
                OleAdaptaterschemaSage.Fill(OleSchemaDatasetSage)
                OledatableSchemaSage = OleSchemaDatasetSage.Tables(0)

                OleAdaptaterschema = New OleDbDataAdapter("select * from P_COLONNEST WHERE CodeTbls='SUP' AND Entete=True ORDER BY ORDRE ", OleConnenection)
                OleSchemaDataset = New DataSet
                OleAdaptaterschema.Fill(OleSchemaDataset)
                OledatableSchema = OleSchemaDataset.Tables(0)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        ExtractionFournisseur(ListeChampSageEntete)
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
        Error_journal = File.AppendText(PathsFileFormatiers & "SUP09" & Strings.Right(DateAndTime.Year(Now), 4) & Format(DateAndTime.Month(Now), "00") & Format(DateAndTime.Day(Now), "00") & Format(DateAndTime.Hour(Now), "00") & Format(DateAndTime.Minute(Now), "00") & Format(DateAndTime.Second(Now), "00") & ".txt")
        Try
            If OledatableSchemaSage.Rows.Count <> 0 Then
                For i As Integer = 0 To OledatableSchemaSage.Rows.Count - 1
                    ListBox.Items.Add("Extraction du Fournisseur : " & OledatableSchemaSage.Rows(i).Item("CT_NUM"))
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
                                                Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(0.0, PositionLeft)
                                            End If
                                        End If
                                    End If
                                End If
                                Continue For
                            Case "SUPPLIERCLASS_CODE"
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

                            Case "ADDR_NAME_FROM"
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
                            Case "ADDR_LINE1_FROM"
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
                            Case "ADDR_LINE2_FROM"
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
                            Case "CITY_FROM"
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
                            Case "STATE_FROM"
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
                            Case "COUNTRY_FROM"
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
                            Case "ZIP_FROM"
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
                            Case "CONTACT_NAME_FROM"
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
                            Case "CONTACT_TEL_FROM"
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
                            Case "CONTACT_EXT_FROM"
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
                            Case "CONTACT_FAX_FROM"
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
                            Case "COMMENT_FROM"
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
                            Case "ADDR_NAME_TO"
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
                            Case "ADDR_LINE1_TO"
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

                            Case "ADDR_LINE2_TO"
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
                            Case "CITY_TO"
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
                            Case "STATE_TO"
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
                            Case "COUNTRY_TO"
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
                            Case "ZIP_TO"
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
                            Case "CONTACT_NAME_TO"
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
                            Case "CONTACT_TEL_TO"
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

                            Case "CONTACT_EXT_TO"
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
                            Case "CONTACT_FAX_TO"
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
                            Case "COMMENT_TO"
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
                            Case "CARRIER_CODE"
                                PositionLeft = OledatableSchema.Rows(Ligne).Item("Format").ToString.Split("(")(1).ToString.Split(")")(0).Split(".")(0)
                                OleAdaptaterschemaFourssAR = New OleDbDataAdapter("SELECT P_EXPEDITION.E_Intitule  FROM F_COMPTET INNER JOIN P_EXPEDITION  ON F_COMPTET.CT_Num='" & OledatableSchemaSage.Rows(i).Item("CT_NUM").ToString & "' AND F_COMPTET.cbMarq =P_EXPEDITION.cbMarq  AND F_COMPTET.CT_Type =1", OleExcelConnect)
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
                                OleAdaptaterschemaFourssAR = New OleDbDataAdapter("SELECT P_EXPEDITION.E_Intitule  FROM F_COMPTET INNER JOIN P_EXPEDITION  ON F_COMPTET.CT_Num='" & OledatableSchemaSage.Rows(i).Item("CT_NUM").ToString & "' AND F_COMPTET.cbMarq =P_EXPEDITION.cbMarq  AND F_COMPTET.CT_Type =1", OleExcelConnect)
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
                            Case "TERMS"
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
                            Case "DAMAGED_OWNER_CODE"
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
                                                Information &= Strings.LSet(RenvoiValeurLue(0.0, OledatableSchema.Rows(Ligne).Item("Format").ToString), PositionLeft)
                                            Else
                                                Information &= Strings.LSet(0.0, PositionLeft)
                                            End If
                                        End If
                                    End If
                                End If
                                Continue For
                        End Select
                        '-------------------------------------**********FIN********************------------------------------------------
                    Next
                    Error_journal.WriteLine(Information)
                    Information = ""
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
