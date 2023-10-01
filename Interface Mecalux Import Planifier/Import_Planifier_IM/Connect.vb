Option Explicit On
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Data.OleDb
Imports Objets100Lib
Imports System.Data.SqlClient
Imports System.Security.Cryptography
Module Connect
    Public Declare Function CloseClipboard Lib "user32" () As Long
    Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare Function EmptyClipboard Lib "user32" () As Long
    Public Idexe As Integer = 0
    Public Declare Ansi Function GetPrivateProfileString Lib "kernel32" _
           Alias "GetPrivateProfileStringA" (ByVal D_Pougoue As String, _
           ByVal P_Hermann As String, ByVal H_Junior As String, _
           ByVal J_Djoumdjeu As String, ByVal H_Pougoue As Integer, _
           ByVal D_Junior As String) As Long

    Public Declare Ansi Function WritePrivateProfileString Lib "kernel32" _
            Alias "WritePrivateProfileStringA" (ByVal App_Section As String, ByVal App_Cle As String, ByVal App_Valeur As String, ByVal App_Path As String) As Boolean

    Public BaseCpta, OM_BaseCpta As New BSCPTAApplication3
    Public BaseCial As New BSCIALApplication3
    Public Error_journal, FichierCSO, ErreurJrn, _
    Error_journalArt, ErreurJrnArt, _
    Error_journalClient, ErreurJrnClient, Error_journalBCF, ErreurJrnBCF, _
    Error_journalFrs, ErreurJrnFrs, _
    Error_journalPseudo, ErreurJrnPseudo, _
    Error_journalCC, ErreurJrnCC, _
    Error_journalCR, ErreurJrnCR, _
    Error_journalVS, ErreurJrnVS As StreamWriter

    'variable de connexion SQL sur different ecran 
    Public OleExcelConnect As New OleDbConnection
    Public OleExcelConnected, OleExcelConnectedArticle, _
    OleExcelConnectedPseudo, OleExcelConnectedClient, _
    OleExcelConnectedFrs, OleExcelConnectedBC, OleExcelConnectBCF, _
    OleExcelConnectedCC, OleExcelConnectedCR, OleExcelConnectedVS As New OleDbConnection

    'variable de connexion SQL sur different ecran 
    Public OleConnenection, OleConnenectionArticle, _
           OleConnenectionPseudo, OleConnenectionClient, _
           OleConnenectionFrs, OleConnenectionBC, OleConnenectionBCF, _
           OleConnenectionCC, OleConnenectionCR, OleConnenectionVS, BaseSQLConnection As New OleDbConnection

    Public NomBaseGesCom, Nom_UtilGes, Mot_PassqlGes, PathsFileCSO, PathsFileCRP, PathsFileVST, PathsFileXFERT, PathsFileMECALUX, PathsFileCSV_ERROR, Pougoue_Fichier As String
    Public NomBaseCpta, Pathsfilejournal, PathsfileSave As String
    Public PathsFileFormatiers, PatchImportftp, DO_StatutClient, DO_StatutFournisseur, PathsFileRecuperer, PathsFileFormatArticle As String
    Public fonctionnement, DecimalMone, DecimalNomb As String

    'Public Nom_Cession, Nom_Etablissement, Nom_A_Nouveau, Nom_Tiers, Nom_Cloture, Nom_Section As String
    Public NomServersql, PathsBaseCpta, StatutConsolider, DatabaseUrl, Nom_Util As String
    Public Mot_Passql, Flagtampon, FlagtamponArticle, Nom_Utilsql, Mot_Pas, PathsFileAccess, PathsfileExport As String
    Public Comptabool, Sqlbool, AccessData, bool As Boolean
    ' Public Article As IBOArticle3
    ''Pour se connecter à la table F_COMPTEG de la base maître
    Public IsJournalExport As Integer = 0
    Public IsJournalExports As Integer
    Public OleSocieteConnect As OleDbConnection
    'Autres variables et constantes
    Public chemin As String
    Public user As String
    Public pwd As String
    Public stat As String
    Public JournalConnection As OleDbConnection
    Public CodeEDI, LOT As Boolean

    Public nbreligne As Integer = 0
    Public selectindex As String = ""
    Public selectindexe, selectindexe1, selectindexe2, selectindexe3, selectindexe4, selectindexe5, selectindexe6, selectindexe7, selectindexe8, selectindexe9, selectindexe10, selectindexe11 As String
    Public selectid As String()
    Public selectids, selectids1, selectids2, selectids3, selectids4, selectids5, selectids6, selectids7, selectids8, selectids9 As String()
    Public Function Base64Encode(ByVal Texte As String) As String
        Try
            Dim texteBytes As Byte() = Encoding.ASCII.GetBytes(Texte)
            If texteBytes.Length = 0 Then
                Return ""
            Else
                Return Convert.ToBase64String(texteBytes)
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function
    Public Function Base64Decode(ByVal Texte As String) As String
        Try
            If Texte.Length = 0 Then
                Return ""
            Else
                Return Encoding.ASCII.GetString(Convert.FromBase64String(Texte))
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function
    Public Function ExisteTableSQL(ByRef TableSQL As String) As Boolean
        Try
            Dim ServeurAdap As OleDbDataAdapter
            Dim ServeurDs As DataSet
            Dim ServeurTab As DataTable
            ServeurAdap = New OleDbDataAdapter("select TOP(1)* from  sysobjects  where name ='" & Trim(TableSQL) & "'", BaseSQLConnection)
            ServeurDs = New DataSet
            ServeurAdap.Fill(ServeurDs)
            ServeurTab = ServeurDs.Tables(0)
            If ServeurTab.Rows.Count <> 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Function RenvoiMontantConditionnement(ByVal Valeur As Object, ByVal Decimale As Integer, ByRef Sepnombre As String, ByVal SepMonetaire As String) As Double
        If Sepnombre = SepMonetaire Then
            Valeur = CDbl(Join(Split(Join(Split(Valeur, "."), Trim(SepMonetaire)), ","), Trim(SepMonetaire)))
        Else
            Valeur = CDbl(Join(Split(Valeur, "."), ","))
        End If
        If Decimale = 0 Then
            Valeur = Math.Round(Valeur, MidpointRounding.AwayFromZero)
        Else
            Valeur = Math.Round(Valeur, Decimale)
        End If
        RenvoiMontantConditionnement = Valeur
    End Function
    Public Function Renvoietypeformat(ByRef Formatintegrer As String) As String
        If Trim(Formatintegrer) = "Point virgule" Then

            Return "Délimité"
        Else
            Return Trim(Formatintegrer)
        End If
    End Function
    Public Function LoginAuFichierExcel(ByRef sPatch As String) As Boolean
        Try
            OleExcelConnected = New OleDbConnection("provider=Microsoft.ACE.OLEDB.12.0; " & "data source=" _
                              & sPatch & "; " & "Extended Properties=""Excel 12.0;HDR=NO;IMEX=1;""")
            OleExcelConnected.Open()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Function RetourneServeurFtp(ByRef GlobalPatch As String) As String
        Dim ServeurFtp() As String = Nothing
        ServeurFtp = Split(GlobalPatch, "/")
        If UBound(ServeurFtp) >= 3 Then
            ServeurFtp = Split(ServeurFtp(2), "@")
            If UBound(ServeurFtp) >= 1 Then
                Return Trim(ServeurFtp(UBound(ServeurFtp)))
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If
    End Function
    Public Function RetourneServeurFtping(ByRef GlobalPatch As String) As String
        Dim ServeurFtp() As String = Nothing
        ServeurFtp = Split(GlobalPatch, "/")
        If UBound(ServeurFtp) >= 3 Then
            ServeurFtp = Split(ServeurFtp(2), "@")
            If UBound(ServeurFtp) >= 1 Then
                ServeurFtp = Split(Trim(ServeurFtp(UBound(ServeurFtp))), ":")
                If UBound(ServeurFtp) >= 0 Then
                    Return Trim(ServeurFtp(0))
                Else
                    Return Nothing
                End If
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If
    End Function
    Public Function RetourneUserFtp(ByRef GlobalPatch As String) As String
        Dim UserFtp() As String = Nothing
        UserFtp = Split(GlobalPatch, ":")
        If UBound(UserFtp) >= 3 Then
            UserFtp = Split(UserFtp(1), "//")
            If UBound(UserFtp) >= 1 Then
                Return Trim(UserFtp(UBound(UserFtp)))
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If
    End Function
    Public Function RetournePassWordFtp(ByRef GlobalPatch As String) As String
        Dim PassWordFtp() As String = Nothing
        PassWordFtp = Split(GlobalPatch, ":")
        If UBound(PassWordFtp) >= 3 Then
            PassWordFtp = Split(PassWordFtp(2), "@")
            If UBound(PassWordFtp) >= 1 Then
                Return Trim(PassWordFtp(0))
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If
    End Function
    Public Function RetourneDirectoryFtp(ByRef GlobalPatch As String) As String
        Dim DirectoryFtp() As String = Nothing
        DirectoryFtp = Split(GlobalPatch, ":")
        If UBound(DirectoryFtp) >= 3 Then
            If InStr(Trim(DirectoryFtp(3)), "/") <> 0 Then
                Return Strings.Right(Trim(DirectoryFtp(3)), Len(Trim(DirectoryFtp(3))) - InStr(Trim(DirectoryFtp(3)), "/"))
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If
    End Function
    Public Function RetourneServeurSQL(ByRef GlobalPatch As String) As String
        Dim PassWordFtp() As String = Nothing
        PassWordFtp = Split(GlobalPatch, "//")
        If UBound(PassWordFtp) >= 1 Then
            PassWordFtp = Split(PassWordFtp(1), "/")
            If UBound(PassWordFtp) >= 2 Then
                If ExisteServeurSQL(Trim(PassWordFtp(0))) Is Nothing Then
                    Return Nothing
                Else
                    Return Trim(PassWordFtp(0))
                End If
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If
    End Function
    Public Function RetourneBaseSQL(ByRef GlobalPatch As String) As String
        Dim PassWordFtp() As String = Nothing
        PassWordFtp = Split(GlobalPatch, "//")
        If UBound(PassWordFtp) >= 1 Then
            PassWordFtp = Split(PassWordFtp(1), "/")
            If UBound(PassWordFtp) >= 2 Then
                Return Trim(PassWordFtp(1))
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If
    End Function
    Public Function RetourneTableSQL(ByRef GlobalPatch As String) As String
        Dim PassWordFtp() As String = Nothing
        PassWordFtp = Split(GlobalPatch, "//")
        If UBound(PassWordFtp) >= 1 Then
            PassWordFtp = Split(PassWordFtp(1), "/")
            If UBound(PassWordFtp) >= 2 Then
                Return Trim(PassWordFtp(2))
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If
    End Function
    Public Sub CreateComboBoxColumn(ByRef Dataobject As DataGridView, ByRef Colname As String, ByRef HeaderName As String)
        Dim ocolumn As New DataGridViewTextBoxColumn
        With ocolumn
            .Name = HeaderName
            .HeaderText = Colname
            .Width = 100
            .Visible = True
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
            .ReadOnly = True
            .SortMode = DataGridViewColumnSortMode.NotSortable
        End With
        Dataobject.Columns.Add(ocolumn)
    End Sub
    Function LireChaine(ByVal Kamen_Fichier_Ini As String, ByVal Pouliyou_Section As String, ByVal Djantcheu_Key As String) As String
        Dim X As Long
        Dim Ham_buffer As String

        Ham_buffer = Space(300)
        X = GetPrivateProfileString(Pouliyou_Section, Djantcheu_Key, "", Ham_buffer, Len(Ham_buffer), Kamen_Fichier_Ini)
        If Len(Trim(Left(Ham_buffer, 295))) > 0 Then
            LireChaine = Left(Trim(Left(Ham_buffer, 295)), Len(Trim(Left(Ham_buffer, 295))) - 1)
        Else
            LireChaine = Nothing
        End If
    End Function
    'Pour fermer la base esclave
    Public Sub FermeBaseCpta(ByRef BaseCpta As BSCPTAApplication3)
        Try
            If BaseCpta.IsOpen = True Then
                BaseCpta.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Public Function OpenBaseCpta(ByRef BaseCpta As BSCPTAApplication3, ByRef FichierCpta As String, Optional ByVal Utilisateur As String = "", Optional ByVal MotDePasse As String = "") As Boolean
        Try
            BaseCpta.Name = FichierCpta
            If Utilisateur <> "" Then
                BaseCpta.Loggable.UserName = Utilisateur
                BaseCpta.Loggable.UserPwd = MotDePasse
            End If
            BaseCpta.Open()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Function OuvreBaseCial(ByRef BaseCial As BSCIALApplication3, ByRef BaseCpta As BSCPTAApplication3, ByVal GescomChemin As String, ByVal FichierCpta As String, Optional ByVal GescomUserName As String = "", Optional ByVal GescomPasswd As String = "", Optional ByVal ComptaUserName As String = "", Optional ByVal ComptaPasswd As String = "") As Boolean
        Try
            BaseCpta.Name = FichierCpta
            If ComptaUserName <> "" Then
                BaseCpta.Loggable.UserName = ComptaUserName
                BaseCpta.Loggable.UserPwd = ComptaPasswd
            End If
            BaseCial.CptaApplication = BaseCpta
            BaseCial.Name = GescomChemin
            If GescomUserName <> "" Then
                BaseCial.Loggable.UserName = GescomUserName
                BaseCial.Loggable.UserPwd = GescomPasswd
            End If
            BaseCial.Open()
            Return True
        Catch ex As Exception
            Return False
            MessageBox.Show(ex.Message)
        End Try
    End Function
    Public Function FermeBaseCial(ByRef BaseCial As BSCIALApplication3) As Boolean
        Try
            If BaseCial.IsOpen = True Then
                BaseCial.Close()
                Return True
            End If
            If BaseCpta.IsOpen = True Then
                BaseCpta.Close()
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Function LirefichierConfig()
        Pougoue_Fichier = My.Application.Info.DirectoryPath & "\ConnectAPI.Ini"
        Flagtampon = LireChaine(Pougoue_Fichier, "CONNECTION", "FLAGTAMPON")
        FlagtamponArticle = LireChaine(Pougoue_Fichier, "CONNECTION", "FLAG")
        NomBaseCpta = LireChaine(Pougoue_Fichier, "CONNECTION", "BASE DE DONNEES COMPTA")
        PathsBaseCpta = LireChaine(Pougoue_Fichier, "CONNECTION", "CHEMIN DU FICHIER COMPTA")
        Nom_Util = LireChaine(Pougoue_Fichier, "CONNECTION", "UTILISATEUR")
        Mot_Pas = LireChaine(Pougoue_Fichier, "CONNECTION", "MOT DE PASSE")
        NomServersql = LireChaine(Pougoue_Fichier, "CONNECTION", "SERVEUR SQL")
        Mot_Passql = LireChaine(Pougoue_Fichier, "CONNECTION", "MOT DE PASSE SQL")
        Nom_Utilsql = LireChaine(Pougoue_Fichier, "CONNECTION", "UTILISATEUR SQL")
        PathsfileSave = LireChaine(Pougoue_Fichier, "CONNECTION", "REPERTOIRE SAUVEGARDE")
        PathsFileAccess = LireChaine(Pougoue_Fichier, "CONNECTION", "NOM FICHIER ACCESS")
        PathsFileFormatiers = LireChaine(Pougoue_Fichier, "CONNECTION", "REPERTOIRE FORMAT TIERS")




        DO_StatutClient = LireChaine(Pougoue_Fichier, "STATUT DU DOCUMENT", "DO_STATUT CLIENT")
        DO_StatutFournisseur = LireChaine(Pougoue_Fichier, "STATUT DU DOCUMENT", "DO_STATUT FOURNISSEUR")
        PathsFileCSO = LireChaine(Pougoue_Fichier, "PARAMETRAGE", "REPERTOIRE(ACHAT) TEMPORAIRE")
        PathsFileCRP = LireChaine(Pougoue_Fichier, "PARAMETRAGE", "REPERTOIRE(VENTE) TEMPORAIRE")
        PathsFileVST = LireChaine(Pougoue_Fichier, "PARAMETRAGE", "REPERTOIRE(MVT STOCK) TEMPORAIRE")

        PathsFileXFERT = LireChaine(Pougoue_Fichier, "PARAMETRAGE", "REPERTOIRE(TRANSFERT) TEMPORAIRE")
        PathsFileCSV_ERROR = LireChaine(Pougoue_Fichier, "PARAMETRAGE", "REPERTOIRE(ERREUR TRANSFORMATION) TEMPORAIRE")
        PathsFileMECALUX = LireChaine(Pougoue_Fichier, "PARAMETRAGE", "REPERTOIRE(SAUVEGADE FICHIERS MECALUX) TEMPORAIRE")


        PathsFileFormatArticle = LireChaine(Pougoue_Fichier, "CONNECTION", "REPERTOIRE FORMAT ARTICLES")
        Pathsfilejournal = LireChaine(Pougoue_Fichier, "CONNECTION", "REPERTOIRE JOURNAL")
        PathsfileExport = LireChaine(Pougoue_Fichier, "CONNECTION", "REPERTOIRE FORMAT ARTICLES")
        DatabaseUrl = LireChaine(Pougoue_Fichier, "CONNECTION", "DATABASEURL")
        If LireChaine(Pougoue_Fichier, "AUTRES", "CODE EDI") <> Nothing Then
            CodeEDI = LireChaine(Pougoue_Fichier, "AUTRES", "CODE EDI")
        End If
        If LireChaine(Pougoue_Fichier, "COMMENTAIRE", "LOT") <> Nothing Then
            LOT = LireChaine(Pougoue_Fichier, "COMMENTAIRE", "LOT")
        End If
        LirefichierConfig = Nothing
    End Function
    Public Function RenvoiMontant(ByVal Valeur As Object, ByVal Decimale As Integer, ByRef Sepnombre As String, ByVal SepMonetaire As String, Optional ByVal Descriptif_decimal As Integer = 0) As Double
        If Sepnombre = SepMonetaire Then
            Valeur = CDbl(Join(Split(Join(Split(Valeur, "."), Trim(SepMonetaire)), ","), Trim(SepMonetaire)))
        Else
            Valeur = CDbl(Join(Split(Valeur, "."), ","))
        End If
        If Decimale = 0 Then
            Valeur = Math.Round(Valeur / Math.Pow(10, Descriptif_decimal), 0)
        Else
            Valeur = Math.Round(Valeur / Math.Pow(10, Descriptif_decimal), Decimale)
        End If
        RenvoiMontant = Valeur
    End Function
    Public Function RenvoiTaux(ByVal Valeur As Object, ByRef Sepnombre As String, ByVal SepMonetaire As String) As Double
        If Sepnombre = SepMonetaire Then
            Valeur = CDbl(Join(Split(Join(Split(Valeur, "."), Trim(SepMonetaire)), ","), Trim(SepMonetaire)))
        Else
            Valeur = CDbl(Join(Split(Valeur, "."), ","))
        End If
        RenvoiTaux = Valeur
    End Function
    Public Function EstNumeric(ByVal Valeur As Object, ByRef Sepnombre As String, ByVal SepMonetaire As String) As Boolean
        EstNumeric = False
        If Sepnombre = SepMonetaire Then
            Valeur = Join(Split(Join(Split(Valeur, "."), Trim(SepMonetaire)), ","), Trim(SepMonetaire))
        Else
            Valeur = Join(Split(Valeur, "."), ",")
        End If
        If IsNumeric(Valeur) = True Then
            EstNumeric = True
        End If
    End Function
    Public Function Connected() As Boolean
        Try
            OleConnenection = New OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;Data Source=" & PathsFileAccess & "")
            OleConnenection.Open()
            OleConnenectionArticle = OleConnenection
            OleConnenectionPseudo = OleConnenection
            OleConnenectionClient = OleConnenection
            OleConnenectionFrs = OleConnenection
            OleConnenectionBC = OleConnenection
            OleConnenectionBCF = OleConnenection
            OleConnenectionCC = OleConnenection
            OleConnenectionCR = OleConnenection
            OleConnenectionvs = OleConnenection
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Function Formatage_Chaine(ByRef Chaine As Object) As Object
        Dim test As String = """"
        Formatage_Chaine = Strings.UCase(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Trim(Chaine), "²"), ""), "&"), ""), "é"), ""), "~"), ""), "#"), ""), "'"), ""), "{"), ""), "("), ""), "["), ""), "-"), ""), "|"), ""), "è"), ""), "`"), ""), "_"), ""), "\"), ""), "ç"), ""), "^"), ""), "à"), ""), "@"), ""), ")"), ""), "]"), ""), "="), ""), "}"), ""), "€"), ""), "^"), ""), "¨"), ""), "$"), ""), "£"), ""), "¤"), ""), "ù"), ""), "%"), ""), "*"), ""), "µ"), ""), "<"), ""), ">"), ""), ","), ""), "?"), ""), ";"), ""), "."), ""), ":"), ""), "/"), ""), "!"), ""), "§"), ""), "+"), ""), test), ""), " "), ""))
    End Function
    Public Function Formatage_Article(ByRef Chaine As Object) As Object
        Dim test As String = """"
        Formatage_Article = Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Join(Split(Trim(Chaine), "²"), ""), "&"), ""), "é"), ""), "~"), ""), "#"), ""), "'"), ""), "{"), ""), "("), ""), "["), ""), "|"), ""), "è"), ""), "`"), ""), "\"), ""), "ç"), ""), "^"), ""), "à"), ""), "@"), ""), ")"), ""), "]"), ""), "="), ""), "}"), ""), "€"), ""), "^"), ""), "¨"), ""), "£"), ""), "¤"), ""), "ù"), ""), "%"), ""), "*"), ""), "µ"), ""), "<"), ""), ">"), ""), ","), ""), "?"), ""), ";"), ""), ":"), ""), "!"), ""), "§"), ""), test), ""), " "), "")
    End Function
    Public Function GetArrayFile(ByVal sPath As String, Optional ByRef aLines() As String = Nothing) As Object
        Dim ArrayFill() As String
        Dim etat As Boolean = False
        GetArrayFile = File.ReadAllLines(sPath, Encoding.Default)
        Dim NewArrayFill(UBound(File.ReadAllLines(sPath, Encoding.Default))) As String
        ArrayFill = File.ReadAllLines(sPath, Encoding.Default)
        If ArrayFill(UBound(File.ReadAllLines(sPath, Encoding.Default))) = "" Then
            ReDim Preserve NewArrayFill(UBound(File.ReadAllLines(sPath, Encoding.Default)) - 1)
            etat = True
        End If
        If etat = True Then
            For i As Integer = 0 To UBound(File.ReadAllLines(sPath, Encoding.Default)) - 1
                If ArrayFill(i) <> "" Then
                    NewArrayFill(i) = ArrayFill(i)
                End If
            Next
        Else
            For i As Integer = 0 To UBound(File.ReadAllLines(sPath, Encoding.UTF8))
                If ArrayFill(i) <> "" Then
                    NewArrayFill(i) = ArrayFill(i)
                End If
            Next
        End If
        GetArrayFile = File.ReadAllLines(sPath, Encoding.Default)
        aLines = GetArrayFile
        aLines = NewArrayFill
        Return aLines
    End Function
    Public Function ExcelConnect(ByRef Server As String, ByRef basededonne As String, ByRef utilisateur As String, ByRef motdepasse As String) As Boolean
        Try
            OleExcelConnect = New OleDbConnection("provider=SQLOLEDB;UID=" & Trim(utilisateur) & ";Pwd=" & Trim(motdepasse) & ";Initial Catalog=" & Trim(basededonne.ToUpper) & ";Data Source=" & Trim(Server) & "")
            OleExcelConnect.Open()
            OleExcelConnectedArticle = OleExcelConnect
            OleExcelConnectedPseudo = OleExcelConnect
            OleExcelConnectedClient = OleExcelConnect
            OleExcelConnectedFrs = OleExcelConnect
            OleExcelConnectedBC = OleExcelConnect
            OleExcelConnectBCF = OleExcelConnect
            OleExcelConnectedCC = OleExcelConnect
            OleExcelConnectedCR = OleExcelConnect
            OleExcelConnectedVS = OleExcelConnect
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function ExcelConnectS(ByRef Server As String, ByRef basededonne As String, ByRef utilisateur As String, ByRef motdepasse As String) As Boolean
        Try
            OleExcelConnectedArticle = New OleDbConnection("provider=SQLOLEDB;UID=" & Trim(utilisateur) & ";Pwd=" & Trim(motdepasse) & ";Initial Catalog=" & Trim(basededonne.ToUpper) & ";Data Source=" & Trim(Server) & "")
            OleExcelConnectedArticle.Open()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Function BaseSQLConnexion(ByRef Server As String, ByRef basededonne As String, ByRef utilisateur As String, ByRef motdepasse As String) As Boolean
        Try
            BaseSQLConnection = New OleDbConnection("provider=SQLOLEDB;UID=" & Trim(utilisateur) & ";Pwd=" & Trim(motdepasse) & ";Initial Catalog=" & Trim(basededonne) & ";Data Source=" & Trim(Server) & "")
            BaseSQLConnection.Open()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Function Afficheauuser(ByRef Formataffiche As String) As String
        If Trim(Formataffiche) = "Délimité" Then

            Return "Point virgule"
        Else
            Return Trim(Formataffiche)
        End If
    End Function
    Public Sub DowloadFtp(ByRef Schema As String, ByRef IDdossier As Integer, ByRef Repertoiredesfichier As String)
        Try
            Dim ArtAdaptater As OleDbDataAdapter
            Dim ArtDataset As DataSet
            Dim Artdatatable As DataTable
            Dim k As Integer
            Call LirefichierConfig()
            ArtAdaptater = New OleDbDataAdapter("select * from  " & Schema & " WHERE Cible='FTP' And IDDossier=" & IDdossier & "", OleConnenection)
            ArtDataset = New DataSet
            ArtAdaptater.Fill(ArtDataset)
            Artdatatable = ArtDataset.Tables(0)
            If Artdatatable.Rows.Count <> 0 Then
                encours.Show()
                If Directory.Exists(Pathsfilejournal) = True Then
                    ErreurJrn = File.AppendText(Pathsfilejournal & "LOG_DOWNLOAD_FTP_" & Format(DateAndTime.Year(Now), "0000") & "" & Format(DateAndTime.Month(Now), "00") & "" & Format(DateAndTime.Day(Now), "00") & "_" & "" & Format(DateAndTime.Hour(Now), "00") & "_" & Format(DateAndTime.Minute(Now), "00") & "_" & Format(DateAndTime.Second(Now), "00") & ".txt")
                    For k = 0 To Artdatatable.Rows.Count - 1
                        encours.Refresh()
                        Dim dossierImport, dossierFtp, FTPserveur, FTPuser, FTPpwd As String
                        Dim verifFtp As Boolean
                        If Schema = "SOCIETEIMPORT_FICHEXML" Then
                            dossierFtp = RetourneDirectoryFtp(Trim(Artdatatable.Rows(k).Item("ImportDos")))
                            FTPserveur = RetourneServeurFtp(Trim(Artdatatable.Rows(k).Item("ImportDos")))
                            FTPuser = RetourneUserFtp(Trim(Artdatatable.Rows(k).Item("ImportDos")))
                            FTPpwd = RetournePassWordFtp(Trim(Artdatatable.Rows(k).Item("ImportDos")))
                        Else
                            dossierFtp = RetourneDirectoryFtp(Trim(Artdatatable.Rows(k).Item("CheminFilexport")))
                            FTPserveur = RetourneServeurFtp(Trim(Artdatatable.Rows(k).Item("CheminFilexport")))
                            FTPuser = RetourneUserFtp(Trim(Artdatatable.Rows(k).Item("CheminFilexport")))
                            FTPpwd = RetournePassWordFtp(Trim(Artdatatable.Rows(k).Item("CheminFilexport")))
                        End If
                        If dossierFtp <> "" Then
                            Dim strUrlFichier, strUrlDossier As String
                            Dim listefichierFtp() As String
                            Dim h As Integer
                            strUrlDossier = "FTP://" & FTPserveur & "/" & dossierFtp
                            listefichierFtp = listeFichiers(strUrlDossier, FTPuser, FTPpwd)
                            If listefichierFtp IsNot Nothing = True Then
                                For h = 0 To listefichierFtp.Length - 1
                                    encours.Refresh()
                                    dossierImport = Repertoiredesfichier
                                    strUrlFichier = strUrlDossier & "/" & listefichierFtp(h)
                                    If File.Exists(dossierImport & "\" & System.IO.Path.GetFileName(strUrlFichier)) = False Then
                                        dossierImport = dossierImport & "\" & System.IO.Path.GetFileName(strUrlFichier)
                                        verifFtp = downloadFichier(strUrlFichier, dossierImport, FTPuser, FTPpwd)
                                        If verifFtp = True Then
                                            'effaceFichier(strUrlFichier, FTPuser, FTPpwd, ErreurJrn)
                                        ElseIf verifFtp = False Then
                                            If File.Exists(dossierImport) = True Then
                                                File.Delete(dossierImport)
                                            End If
                                        End If
                                    End If
                                    encours.Refresh()
                                Next h
                            End If
                        End If
                        encours.Refresh()
                    Next k
                End If
            End If
        Catch ex As Exception

        Finally
            If ErreurJrn IsNot Nothing = True Then
                ErreurJrn.Close()
            End If
        End Try
    End Sub
    Public Function downloadFichier(ByVal strUrlFichier As String, ByVal strCheminDestinationFichier As String, ByVal identifiant As String, ByVal motDePasse As String) As Boolean
        Dim monUriFichier As New System.Uri(strUrlFichier)
        Dim monUriDestinationFichier As New System.Uri(strCheminDestinationFichier)
        If Not (monUriFichier.Scheme = Uri.UriSchemeFtp) Then
            ErreurJrn.WriteLine("L'Uri du fichier sur le serveur FTP n'est pas valide", _
                            "Une erreur est surevnue")
            Return False
            Exit Function
        End If
        If Not (monUriDestinationFichier.Scheme = Uri.UriSchemeFile) Then
            ErreurJrn.WriteLine("Le chemin de destination n'est pas valide !", _
                            "Une erreur est surevnue")
            Return False
            Exit Function
        End If

        Dim monResponseStream As Stream = Nothing
        Dim monFileStream As FileStream = Nothing
        Dim monReader As StreamReader = Nothing
        Try
            Dim downloadRequest As FtpWebRequest = CType(WebRequest.Create(monUriFichier), FtpWebRequest)
            If Not identifiant.Length = 0 Then
                Dim monCompteFtp As New NetworkCredential(identifiant, motDePasse)
                downloadRequest.Credentials = monCompteFtp
            End If

            Dim downloadResponse As FtpWebResponse = CType(downloadRequest.GetResponse(), FtpWebResponse)
            monResponseStream = downloadResponse.GetResponseStream()
            Dim nomFichier As String = monUriDestinationFichier.LocalPath.ToString
            monFileStream = File.Create(nomFichier)
            Dim monBuffer(1024) As Byte
            Dim octetsLus As Integer
            While True
                octetsLus = monResponseStream.Read(monBuffer, 0, monBuffer.Length)
                If octetsLus = 0 Then
                    Exit While
                End If
                monFileStream.Write(monBuffer, 0, octetsLus)
            End While
            ErreurJrn.WriteLine("Téléchargement du fichier " & System.IO.Path.GetFileName(strUrlFichier) & "effectué avec succès")
            Return True
            ' Gestion des exceptions
        Catch ex As UriFormatException
            ErreurJrn.WriteLine("Téléchargement du fichier " & System.IO.Path.GetFileName(strUrlFichier) & "effectué avec echec")
            ErreurJrn.WriteLine(ex.Message)
            Return False
        Catch ex As WebException
            ErreurJrn.WriteLine("Téléchargement du fichier " & System.IO.Path.GetFileName(strUrlFichier) & "effectué avec echec")
            ErreurJrn.WriteLine(ex.Message)
            Return False
        Catch ex As IOException
            ErreurJrn.WriteLine("Téléchargement du fichier " & System.IO.Path.GetFileName(strUrlFichier) & "effectué avec echec")
            ErreurJrn.WriteLine(ex.Message)
            Return False
        Finally
            If monReader IsNot Nothing Then
                monReader.Close()
            ElseIf monResponseStream IsNot Nothing Then
                monResponseStream.Close()
            End If
            If monFileStream IsNot Nothing Then
                monFileStream.Close()
            End If
        End Try
    End Function
    Public Function RetourneUserSQL(ByRef GlobalServeur As String) As String
        Dim ServeurAdap As OleDbDataAdapter
        Dim ServeurDs As DataSet
        Dim ServeurTab As DataTable
        ServeurAdap = New OleDbDataAdapter("select * from SERVEURSQL where Ser_ver='" & Trim(GlobalServeur) & "'", OleConnenection)
        ServeurDs = New DataSet
        ServeurAdap.Fill(ServeurDs)
        ServeurTab = ServeurDs.Tables(0)
        If ServeurTab.Rows.Count <> 0 Then
            If Convert.IsDBNull(ServeurTab.Rows(0).Item("Utili_sateur")) = False Then
                Return Trim(ServeurTab.Rows(0).Item("Utili_sateur"))
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If
    End Function
    Public Function ExisteFlag(ByRef TableSQL As String, ByRef Flag As String) As Boolean

        Try
            Dim ServeurAdap As OleDbDataAdapter
            Dim ServeurDs As DataSet
            Dim ServeurTab As DataTable
            ServeurAdap = New OleDbDataAdapter("select TOP(1)* from  " & Trim(TableSQL) & " ", BaseSQLConnection)
            ServeurDs = New DataSet
            ServeurAdap.Fill(ServeurDs)
            ServeurTab = ServeurDs.Tables(0)
            If ServeurTab.Columns.Count <> 0 Then
                If ServeurTab.Columns.Contains(Flag) = True Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Function RetournePasseSQL(ByRef GlobalServeur As String) As String
        Dim ServeurAdap As OleDbDataAdapter
        Dim ServeurDs As DataSet
        Dim ServeurTab As DataTable
        ServeurAdap = New OleDbDataAdapter("select * from SERVEURSQL where Ser_ver='" & Trim(GlobalServeur) & "'", OleConnenection)
        ServeurDs = New DataSet
        ServeurAdap.Fill(ServeurDs)
        ServeurTab = ServeurDs.Tables(0)
        If ServeurTab.Rows.Count <> 0 Then
            If Convert.IsDBNull(ServeurTab.Rows(0).Item("Mot_Passe")) = False Then
                Return ServeurTab.Rows(0).Item("Mot_Passe")
            Else
                Return Nothing
            End If
        Else
            Return Nothing
        End If
    End Function
    Public Sub Updatetampon(ByRef tablestampon As String, ByRef NumLot As String, ByRef Flagchamp As String, ByRef ChampLot As String, ByRef Champtraitement As String, ByRef Itraitement As Integer)
        Try
            Dim createtablesql As String
            Dim OleCreateTable As OleDbCommand
            createtablesql = "SET ARITHABORT ON"
            OleCreateTable = New OleDbCommand(createtablesql)
            OleCreateTable.Connection = BaseSQLConnection
            OleCreateTable.ExecuteNonQuery()
            createtablesql = "UPDATE  " & tablestampon & "  SET " & Flagchamp & "='WAZA' WHERE (" & ChampLot & " ='" & Join(Split(NumLot, "'"), "''") & "') And (" & Champtraitement & " ='" & Join(Split(Itraitement, "'"), "''") & "')"
            OleCreateTable = New OleDbCommand(createtablesql)
            OleCreateTable.Connection = BaseSQLConnection
            OleCreateTable.ExecuteNonQuery()
            createtablesql = "SET ARITHABORT OFF"
            OleCreateTable = New OleDbCommand(createtablesql)
            OleCreateTable.Connection = BaseSQLConnection
            OleCreateTable.ExecuteNonQuery()
        Catch ex As Exception
        End Try
    End Sub
    Public Function RenvoieID(ByRef Schema As String) As Integer
        Dim OleAdaptater As OleDbDataAdapter
        Dim OleAfficheDataset As DataSet
        Dim Oledatable As DataTable
        OleAdaptater = New OleDbDataAdapter("select max(IDDossier) As ID  from " & Schema & "", OleConnenection)
        OleAfficheDataset = New DataSet
        OleAdaptater.Fill(OleAfficheDataset)
        Oledatable = OleAfficheDataset.Tables(0)
        If Oledatable.Rows.Count <> 0 Then
            If Convert.IsDBNull(Oledatable.Rows(0).Item("ID")) = False Then
                RenvoieID = Oledatable.Rows(0).Item("ID") + 1
            Else
                RenvoieID = 1
            End If
        Else

            RenvoieID = 1
        End If
    End Function
    Public Function ExisteServeurSQL(ByRef GlobalServeur As String) As String
        Dim ServeurAdap As OleDbDataAdapter
        Dim ServeurDs As DataSet
        Dim ServeurTab As DataTable
        ServeurAdap = New OleDbDataAdapter("select * from SERVEURSQL where Ser_ver='" & Trim(GlobalServeur) & "'", OleConnenection)
        ServeurDs = New DataSet
        ServeurAdap.Fill(ServeurDs)
        ServeurTab = ServeurDs.Tables(0)
        If ServeurTab.Rows.Count <> 0 Then
            Return Trim(GlobalServeur)
        Else
            Return Nothing
        End If
    End Function
    Public Function ExisteFiltreTable(ByRef TableSQL As String, ByRef NomFiltre As String) As Boolean
        Dim ServeurAdap As OleDbDataAdapter
        Dim ServeurDs As DataSet
        Dim ServeurTab As DataTable
        ServeurAdap = New OleDbDataAdapter("select TOP(1)* from sys.columns  where name ='" & Join(Split(NomFiltre, "'"), "''") & "' And object_id=(select object_id from sys.objects  Where name='" & Join(Split(TableSQL, "'"), "''") & "')", BaseSQLConnection)
        ServeurDs = New DataSet
        ServeurAdap.Fill(ServeurDs)
        ServeurTab = ServeurDs.Tables(0)
        If ServeurTab.Rows.Count <> 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Function Creationtablejournal(ByRef tablesjournal As String) As Boolean
        Try
            If Trim(tablesjournal) <> "" And Len(Trim(tablesjournal)) > 4 Then

            Else
                tablesjournal = "WAZA_JOURNAL"
            End If
            Dim createtablesql As String
            Dim OleCreateTable As OleDbCommand
            createtablesql = "SET ARITHABORT ON"
            OleCreateTable = New OleDbCommand(createtablesql)
            OleCreateTable.Connection = JournalConnection
            OleCreateTable.ExecuteNonQuery()
            createtablesql = "IF NOT EXISTS (select TOP(1)* from  sysobjects  where name ='" & Trim(tablesjournal) & "')" & _
            " BEGIN " & _
            "CREATE TABLE [dbo].[" & tablesjournal & "](" & _
            "[PDate] [datetime] NULL, " & _
            "[LDate] [datetime] NULL, " & _
            "[Lot] [varchar](max) NULL, " & _
            "[Traitement] [varchar](max)  NULL, " & _
            "[Bloquant] [smallint] NULL, " & _
            "[Categorie] [smallint] NULL, " & _
            "[Nature] [varchar](max)  NULL, " & _
            "[cbMarq] [int] IDENTITY(1,1) NOT NULL, " & _
            "CONSTRAINT [PK_CBMARQ_WAZA_JOURNAL] PRIMARY KEY CLUSTERED " & _
            "( [cbMarq] ASC ) " & _
            ")" & _
            " END"
            OleCreateTable = New OleDbCommand(createtablesql)
            OleCreateTable.Connection = JournalConnection
            OleCreateTable.ExecuteNonQuery()
            createtablesql = "SET ARITHABORT OFF"
            OleCreateTable = New OleDbCommand(createtablesql)
            OleCreateTable.Connection = JournalConnection
            OleCreateTable.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Function JournalConnexion(ByRef Server As String, ByRef basededonne As String, ByRef utilisateur As String, ByRef motdepasse As String) As Boolean
        Try
            JournalConnection = New OleDbConnection("provider=SQLOLEDB;UID=" & Trim(utilisateur) & ";Pwd=" & Trim(motdepasse) & ";Initial Catalog=" & Trim(basededonne) & ";Data Source=" & Trim(Server) & "")
            JournalConnection.Open()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Function listeFichiers(ByVal serveurCible As String, ByVal identifiant As String, ByVal motDePasse As String) As Array
        Dim monResponseStream As Stream = Nothing
        Dim monStreamReader As StreamReader = Nothing
        Dim monResultat As Array = Nothing
        Dim monUriServeur As New System.Uri(serveurCible)
        If Not (monUriServeur.Scheme = Uri.UriSchemeFtp) Then
            ErreurJrn.WriteLine("L'Uri du serveur FTP n'est pas valide", _
                            "Une erreur est surevnue")
            'Si Uri non valide, arrêt du listage
            Return monResultat
            Exit Function
        End If
        Try

            Dim maRequeteListe As FtpWebRequest = CType(WebRequest.Create(monUriServeur), FtpWebRequest)
            maRequeteListe.Method = WebRequestMethods.Ftp.ListDirectoryDetails
            If Not identifiant.Length = 0 Then
                Dim monCompteFtp As New NetworkCredential(identifiant, motDePasse)
                maRequeteListe.Credentials = monCompteFtp
            End If

            Dim maResponseListe As FtpWebResponse = CType(maRequeteListe.GetResponse, FtpWebResponse)
            monStreamReader = New StreamReader(maResponseListe.GetResponseStream, _
                                               Encoding.Default)
            Dim listeBrute As String = monStreamReader.ReadToEnd
            Dim separateur() As String = {Environment.NewLine} ' -> retour chariot
            Dim tableauListe() As String = listeBrute.Split(separateur, _
                                           StringSplitOptions.RemoveEmptyEntries)
            Dim listeFinale As New List(Of String)
            Dim i As Integer = 0
            While i < tableauListe.Length
                If Not tableauListe(i).StartsWith("d") Then
                    listeFinale.Add(tableauListe(i).Substring(tableauListe(i).LastIndexOf(" ") + 1))
                End If
                i += 1
            End While
            monResultat = listeFinale.ToArray
            ErreurJrn.WriteLine("Liste terminée")
            ' Gestion des exceptions
        Catch ex As UriFormatException
            ErreurJrn.WriteLine(ex.Message)
        Catch ex As WebException
            ErreurJrn.WriteLine(ex.Message)
        Finally
            If monResponseStream IsNot Nothing Then
                monResponseStream.Close()
            End If
            If monStreamReader IsNot Nothing Then
                monStreamReader.Close()
            End If
        End Try
        Return monResultat
    End Function
    Public Sub DowloadFtpMvt(ByRef Schema As String, ByRef Categorie As String)
        Try
            Dim ArtAdaptater As OleDbDataAdapter
            Dim ArtDataset As DataSet
            Dim Artdatatable As DataTable
            Dim k As Integer
            Call LirefichierConfig()
            ArtAdaptater = New OleDbDataAdapter("select * from  " & Schema & " WHERE Cible='FTP' And Categorie='" & Categorie & "'", OleConnenection)
            ArtDataset = New DataSet
            ArtAdaptater.Fill(ArtDataset)
            Artdatatable = ArtDataset.Tables(0)
            If Artdatatable.Rows.Count <> 0 Then
                encours.Show()
                If Directory.Exists(Pathsfilejournal) = True Then
                    ErreurJrn = File.AppendText(Pathsfilejournal & "LOGIMPORT_FTP_" & Format(DateAndTime.Year(Now), "0000") & "" & Format(DateAndTime.Month(Now), "00") & "" & Format(DateAndTime.Day(Now), "00") & "_" & "" & Format(DateAndTime.Hour(Now), "00") & "_" & Format(DateAndTime.Minute(Now), "00") & "_" & Format(DateAndTime.Second(Now), "00") & ".txt")
                    For k = 0 To Artdatatable.Rows.Count - 1
                        encours.Refresh()
                        Dim dossierImport, dossierFtp, FTPserveur, FTPuser, FTPpwd As String
                        Dim verifFtp As Boolean
                        dossierFtp = Artdatatable.Rows(k).Item("RepertoireFTP").ToString
                        FTPserveur = Artdatatable.Rows(k).Item("ServeurFtp").ToString
                        FTPuser = Artdatatable.Rows(k).Item("UserFtp").ToString
                        FTPpwd = Artdatatable.Rows(k).Item("PwdFtp").ToString
                        If dossierFtp <> "" Then
                            Dim strUrlFichier, strUrlDossier As String
                            Dim listefichierFtp() As String
                            Dim h As Integer
                            strUrlDossier = "FTP://" & FTPserveur & "/" & dossierFtp
                            listefichierFtp = listeFichiers(strUrlDossier, FTPuser, FTPpwd)
                            If listefichierFtp IsNot Nothing = True Then
                                For h = 0 To listefichierFtp.Length - 1
                                    encours.Refresh()

                                    dossierImport = Trim(Artdatatable.Rows(k).Item("CheminFilexport"))
                                    strUrlFichier = strUrlDossier & "/" & listefichierFtp(h)
                                    If File.Exists(dossierImport & System.IO.Path.GetFileName(strUrlFichier)) = False Then
                                        dossierImport = dossierImport & System.IO.Path.GetFileName(strUrlFichier)
                                        verifFtp = downloadFichier(strUrlFichier, dossierImport, FTPuser, FTPpwd)
                                        If verifFtp = True Then
                                            'effaceFichier(strUrlFichier, FTPuser, FTPpwd)
                                        ElseIf verifFtp = False Then
                                            If File.Exists(dossierImport) = True Then
                                                File.Delete(dossierImport)
                                            End If
                                        End If
                                    End If
                                    encours.Refresh()
                                Next h
                            End If
                        End If
                        encours.Refresh()
                    Next k
                End If
            End If
        Catch ex As Exception

        Finally
            If ErreurJrn IsNot Nothing = True Then
                ErreurJrn.Close()
            End If
        End Try
    End Sub
    Public Function effaceFichier(ByVal uriFichier As String, ByVal identifiant As String, ByVal motDePasse As String, ByRef WrteJournal As StreamWriter) As Boolean
        WrteJournal.WriteLine("-----------------------------------------------------------------------------------------------------")
        WrteJournal.WriteLine("Action de suppresion du fichier : ")
        Dim monUriFichier As New Uri(uriFichier)
        If Not (monUriFichier.Scheme = Uri.UriSchemeFtp) Then
            WrteJournal.WriteLine("L'URI du fichier à supprimer n'est pas valide ou le fichier est introuvable")
            Return False
            Exit Function
        End If
        Try
            Dim maRequeteEffacement As FtpWebRequest = CType(WebRequest.Create(uriFichier), FtpWebRequest)
            maRequeteEffacement.Method = WebRequestMethods.Ftp.DeleteFile
            If Not identifiant.Length = 0 Then
                Dim monCompteFtp As New NetworkCredential(identifiant, motDePasse)
                maRequeteEffacement.Credentials = monCompteFtp
            End If
            Dim maResponseFtp As FtpWebResponse = CType(maRequeteEffacement.GetResponse, FtpWebResponse)
            WrteJournal.WriteLine("Action de suppression : " & maResponseFtp.StatusDescription)
            Return True
        Catch ex As Exception
            WrteJournal.WriteLine("Une erreur est surevnue message systeme : " & ex.Message)
            Return False
        End Try
    End Function


    Public Function uploadFichier(ByVal cheminSource As String, _
                              ByVal urlDestination As String, _
                              ByVal identifiant As String, _
                              ByVal motDePasse As String, ByRef WrteJournal As StreamWriter) As Boolean
        Dim monUriFichierLocal As System.Uri = New System.Uri(cheminSource)
        Dim monUriFichierDistant As System.Uri = New System.Uri(urlDestination)
        If Not (monUriFichierLocal.Scheme = Uri.UriSchemeFile) Then
            WrteJournal.WriteLine("Le chemin du fichier local n'est pas valide !")
            Return False
            Exit Function
        End If
        If Not (monUriFichierDistant.Scheme = Uri.UriSchemeFtp) Then
            WrteJournal.WriteLine("Le chemin du fichier sur le serveur FTP n'est pas valide !")
            Return False
            Exit Function
        End If

        Dim monRequestStream As Stream = Nothing
        Dim fileStream As FileStream = Nothing
        Dim uploadResponse As FtpWebResponse = Nothing
        Try
            Dim uploadRequest As FtpWebRequest = CType(WebRequest.Create(urlDestination), FtpWebRequest)
            If Not identifiant.Length = 0 Then
                Dim monCompte As New NetworkCredential(identifiant, motDePasse)
                uploadRequest.Credentials = monCompte
            End If

            uploadRequest.Method = WebRequestMethods.Ftp.UploadFile
            uploadRequest.Proxy = Nothing
            monRequestStream = uploadRequest.GetRequestStream()
            fileStream = File.Open(cheminSource, FileMode.Open)
            Dim buffer(1024) As Byte
            Dim bytesRead As Integer
            While True
                bytesRead = fileStream.Read(buffer, 0, buffer.Length)
                If bytesRead = 0 Then
                    Exit While
                End If
                monRequestStream.Write(buffer, 0, bytesRead)
            End While
            monRequestStream.Close()
            uploadResponse = CType(uploadRequest.GetResponse(), FtpWebResponse)
            WrteJournal.WriteLine("Upload terminé.")
            Return True
            ' Gestion des exceptions
        Catch ex As UriFormatException
            ErreurJrn.WriteLine(ex.Message)
            Return False
        Catch ex As WebException
            ErreurJrn.WriteLine(ex.Message)
            Return False
        Catch ex As IOException
            ErreurJrn.WriteLine(ex.Message)
            Return False
        Finally
            If uploadResponse IsNot Nothing Then
                uploadResponse.Close()
            End If
            If fileStream IsNot Nothing Then
                fileStream.Close()
            End If
            If monRequestStream IsNot Nothing Then
                monRequestStream.Close()
            End If
        End Try
    End Function
End Module
