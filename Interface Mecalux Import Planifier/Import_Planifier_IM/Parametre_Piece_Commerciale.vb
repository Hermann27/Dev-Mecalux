Imports System.Data.OleDb
Imports System.IO
Imports System.Net.NetworkInformation
Public Class Parametre_Piece_Commerciale
    Private Sub AfficheSchemasConso()
        Dim i As Integer
        Dim OleAdaptaterschema As OleDbDataAdapter
        Dim OleSchemaDataset As DataSet
        Dim OledatableSchema As DataTable
        DataListeIntegrer.Rows.Clear()

        OleAdaptaterschema = New OleDbDataAdapter("select * from PARACOMMERCIAL", OleConnenection)
        OleSchemaDataset = New DataSet
        OleAdaptaterschema.Fill(OleSchemaDataset)
        OledatableSchema = OleSchemaDataset.Tables(0)
        DataListeIntegrer.RowCount = OledatableSchema.Rows.Count
        For i = 0 To OledatableSchema.Rows.Count - 1
            DataListeIntegrer.Rows(i).Cells("basecpta1").Value = OledatableSchema.Rows(i).Item("BaseCpta")
            DataListeIntegrer.Rows(i).Cells("basegescom1").Value = OledatableSchema.Rows(i).Item("BaseCial")
            DataListeIntegrer.Rows(i).Cells("dostraite1").Value = OledatableSchema.Rows(i).Item("DosTraite")
            DataListeIntegrer.Rows(i).Cells("dosdest1").Value = OledatableSchema.Rows(i).Item("DosSauve")
            DataListeIntegrer.Rows(i).Cells("Piece1").Value = OledatableSchema.Rows(i).Item("Piece")
            DataListeIntegrer.Rows(i).Cells("Format1").Value = OledatableSchema.Rows(i).Item("Format")
            DataListeIntegrer.Rows(i).Cells("IDDossier1").Value = OledatableSchema.Rows(i).Item("IDDossier")
            DataListeIntegrer.Rows(i).Cells("Fournisseur1").Value = OledatableSchema.Rows(i).Item("Fournisseur")


            DataListeIntegrer.Rows(i).Cells("RepertoireFTP1").Value = OledatableSchema.Rows(i).Item("RepertoireFTP")
            DataListeIntegrer.Rows(i).Cells("ServeurFtp1").Value = OledatableSchema.Rows(i).Item("ServeurFtp")
            DataListeIntegrer.Rows(i).Cells("UserFtp1").Value = OledatableSchema.Rows(i).Item("UserFtp")
            DataListeIntegrer.Rows(i).Cells("PwdFtp1").Value = OledatableSchema.Rows(i).Item("PwdFtp")

        Next i
    End Sub

    Private Sub AfficheSocieteCpta()
        Dim OleAdaptater As OleDbDataAdapter
        Dim OleAfficheDataset As DataSet
        Dim Oledatable As DataTable
        Dim i As Integer
        DataListeSchema.Rows.Clear()
        basecpta.Items.Clear()

        Try
            OleAdaptater = New OleDbDataAdapter("select * from PARAMETRE where nomtype='COMPTABILITE'", OleConnenection)
            OleAfficheDataset = New DataSet
            OleAdaptater.Fill(OleAfficheDataset)
            Oledatable = OleAfficheDataset.Tables(0)
            For i = 0 To Oledatable.Rows.Count - 1
                If Trim(Oledatable.Rows(i).Item("Societe")) <> "" Then
                    basecpta.Items.AddRange(New String() {Oledatable.Rows(i).Item("Societe")})
                End If
            Next i

        Catch ex As Exception

        End Try
    End Sub

    Private Sub AfficheSocieteCial()
        Dim OleAdaptater As OleDbDataAdapter
        Dim OleAfficheDataset As DataSet
        Dim Oledatable As DataTable
        Dim i As Integer
        DataListeSchema.Rows.Clear()
        basegescom.Items.Clear()

        Try
            OleAdaptater = New OleDbDataAdapter("select * from PARAMETRE where nomtype='COMMERCIAL'", OleConnenection)
            OleAfficheDataset = New DataSet
            OleAdaptater.Fill(OleAfficheDataset)
            Oledatable = OleAfficheDataset.Tables(0)
            For i = 0 To Oledatable.Rows.Count - 1
                If Trim(Oledatable.Rows(i).Item("Societe")) <> "" Then
                    basegescom.Items.AddRange(New String() {Oledatable.Rows(i).Item("Societe")})
                End If
            Next i

        Catch ex As Exception

        End Try
    End Sub


    Private Sub Delete_DataListeSch()
        Dim i As Integer
        Dim OleAdaptaterDelete As OleDbDataAdapter
        Dim OleDeleteDataset As DataSet
        Dim OledatableDelete As DataTable
        Dim OleCommandDelete As OleDbCommand
        Dim DelFile As String
        For i = 0 To DataListeIntegrer.RowCount - 1
            If DataListeIntegrer.Rows(i).Cells("Supprimer").Value = True Then
                OleAdaptaterDelete = New OleDbDataAdapter("select * from PARACOMMERCIAL where BaseCpta='" & DataListeIntegrer.Rows(i).Cells("basecpta1").Value & "' and BaseCial='" & DataListeIntegrer.Rows(i).Cells("basegescom1").Value & "' and  Piece='" & DataListeIntegrer.Rows(i).Cells("Piece1").Value & "' and DosTraite='" & DataListeIntegrer.Rows(i).Cells("dostraite1").Value & "'", OleConnenection)
                OleDeleteDataset = New DataSet
                OleAdaptaterDelete.Fill(OleDeleteDataset)
                OledatableDelete = OleDeleteDataset.Tables(0)
                If OledatableDelete.Rows.Count <> 0 Then
                    DelFile = "Delete From PARACOMMERCIAL where BaseCpta='" & DataListeIntegrer.Rows(i).Cells("basecpta1").Value & "'and BaseCial='" & DataListeIntegrer.Rows(i).Cells("basegescom1").Value & "' and Piece='" & DataListeIntegrer.Rows(i).Cells("Piece1").Value & "' and DosTraite='" & DataListeIntegrer.Rows(i).Cells("dostraite1").Value & "'"
                    OleCommandDelete = New OleDbCommand(DelFile)
                    OleCommandDelete.Connection = OleConnenection
                    OleCommandDelete.ExecuteNonQuery()
                    DelFile = "Delete From PLANIFICATION where (Intitule='Import EANCOM Devis' Or Intitule='Import EANCOM Commande Vente' Or Intitule='Import EANCOM Bon de Livraison' Or Intitule='Import EANCOM Facture Vente' Or Intitule='Import EANCOM Commande Achat' Or Intitule='Import EANCOM Bon de Reception' Or Intitule='Import EANCOM Facture Achat' Or Intitule='Import EAN96 Devis' Or Intitule='Import EAN96 Commande Vente' Or Intitule='Import EAN96 Bon de Livraison' Or Intitule='Import EAN96 Facture Vente' Or Intitule='Import EAN96 Commande Achat' Or Intitule='Import EAN96 Bon de Reception' Or Intitule='Import EAN96 Facture Achat' Or Intitule='Import Sage 100 Devis' Or Intitule='Import Sage 100 Commande Vente' Or Intitule='Import Sage 100 Bon de Livraison' Or Intitule='Import Sage 100 Facture Vente' Or Intitule='Import Sage 100 Commande Achat' Or Intitule='Import Sage 100 Bon de Reception' Or Intitule='Import Sage 100 Facture Achat') and IDDossier=" & CInt(DataListeIntegrer.Rows(i).Cells("IDDossier1").Value) & ""
                    OleCommandDelete = New OleDbCommand(DelFile)
                    OleCommandDelete.Connection = OleConnenection
                    OleCommandDelete.ExecuteNonQuery()
                End If
            End If
        Next i
        AfficheSchemasConso()
    End Sub
    Private Sub EnregistrerLeSchema()
        Dim n As Integer
        Dim OleAdaptaterEnreg As OleDbDataAdapter
        Dim OleEnregDataset As DataSet
        Dim OledatableEnreg As DataTable
        Dim OleCommandEnreg As OleDbCommand
        Dim Insert As Boolean = False
        Dim Insertion As String
        If DataListeSchema.RowCount >= 1 Then

            For n = 0 To DataListeSchema.RowCount - 1
                If IsNumeric(DataListeSchema.Rows(n).Cells("IDDossier").Value) = True Then
                    If Trim(DataListeSchema.Rows(n).Cells("basecpta").Value) <> "" And Trim(DataListeSchema.Rows(n).Cells("basegescom").Value) <> "" And Trim(DataListeSchema.Rows(n).Cells("Piece").Value) <> "" And Trim(DataListeSchema.Rows(n).Cells("dostraite").Value) <> "" Then
                        OleAdaptaterEnreg = New OleDbDataAdapter("select * From PARACOMMERCIAL WHERE BaseCpta='" & DataListeSchema.Rows(n).Cells("basecpta").Value & "'and BaseCial='" & DataListeSchema.Rows(n).Cells("basegescom").Value & "'and Piece='" & DataListeSchema.Rows(n).Cells("Piece").Value & "'and DosTraite='" & DataListeSchema.Rows(n).Cells("dostraite").Value & "'", OleConnenection)
                        OleEnregDataset = New DataSet
                        OleAdaptaterEnreg.Fill(OleEnregDataset)
                        OledatableEnreg = OleEnregDataset.Tables(0)
                        If OledatableEnreg.Rows.Count <> 0 Then
                            MsgBox("Cette Société Comptable Existe déja!", MsgBoxStyle.Information, "Creation Societe")
                        Else
                            OleAdaptaterEnreg = New OleDbDataAdapter("select * From PARACOMMERCIAL WHERE IDDossier=" & CInt(DataListeSchema.Rows(n).Cells("IDDossier").Value) & " ", OleConnenection)
                            OleEnregDataset = New DataSet
                            OleAdaptaterEnreg.Fill(OleEnregDataset)
                            OledatableEnreg = OleEnregDataset.Tables(0)
                            If OledatableEnreg.Rows.Count <> 0 Then
                            Else
                                If (Trim(DataListeSchema.Rows(n).Cells("basecpta").Value) <> "" And Trim(DataListeSchema.Rows(n).Cells("basegescom").Value) <> "") Then
                                    Insertion = "Insert Into PARACOMMERCIAL (BaseCpta,ServeurFtp,RepertoireFTP,UserFtp,PwdFtp,BaseCial,DosTraite,DosSauve,Piece,Format,Fournisseur,IDDossier) VALUES ('" & DataListeSchema.Rows(n).Cells("basecpta").Value & "','" & DataListeSchema.Rows(n).Cells("ServeurFtp").Value & "','" & DataListeSchema.Rows(n).Cells("RepertoireFTP").Value & "','" & DataListeSchema.Rows(n).Cells("UserFtp").Value & "','" & DataListeSchema.Rows(n).Cells("PwdFtp").Value & "','" & DataListeSchema.Rows(n).Cells("basegescom").Value & "','" & DataListeSchema.Rows(n).Cells("dostraite").Value & "','" & DataListeSchema.Rows(n).Cells("dosdest").Value & "','" & DataListeSchema.Rows(n).Cells("Piece").Value & "','" & DataListeSchema.Rows(n).Cells("Format").Value & "','" & DataListeSchema.Rows(n).Cells("Fournisseur").Value & "'," & CInt(DataListeSchema.Rows(n).Cells("IDDossier").Value) & ")"
                                    OleCommandEnreg = New OleDbCommand(Insertion)
                                    OleCommandEnreg.Connection = OleConnenection
                                    OleCommandEnreg.ExecuteNonQuery()
                                    Insert = True
                                End If
                            End If
                        End If
                    End If
                Else
                    MsgBox("L'ID : " & DataListeSchema.Rows(n).Cells("IDDossier").Value & " du dossier doit être un entier !", MsgBoxStyle.Information, "Insertion des Schemas d'integration")
                End If
            Next n
            If Insert = True Then
                MsgBox("Insertion Reussie", MsgBoxStyle.Information, "Parametrage de la Consolidation")
                DataListeSchema.Rows.Clear()
            End If
        End If
    End Sub
    Private Sub MiseàjourLeSchema()
        Dim n As Integer
        Dim OleAdaptaterEnreg As OleDbDataAdapter
        Dim OleEnregDataset As DataSet
        Dim OledatableEnreg As DataTable
        Dim OleCommandEnreg As OleDbCommand
        Dim Insert As Boolean = False
        Dim Insertion As String
        Try
            If DataListeIntegrer.RowCount >= 0 Then
                For n = 0 To DataListeIntegrer.RowCount - 1
                    OleAdaptaterEnreg = New OleDbDataAdapter("select * From PARACOMMERCIAL WHERE BaseCpta='" & DataListeIntegrer.Rows(n).Cells("basecpta1").Value & "' and BaseCial='" & DataListeIntegrer.Rows(n).Cells("basegescom1").Value & "'and Piece='" & DataListeIntegrer.Rows(n).Cells("Piece1").Value & "' and DosTraite='" & DataListeIntegrer.Rows(n).Cells("dostraite1").Value & "'", OleConnenection)
                    OleEnregDataset = New DataSet
                    OleAdaptaterEnreg.Fill(OleEnregDataset)
                    OledatableEnreg = OleEnregDataset.Tables(0)
                    If OledatableEnreg.Rows.Count <> 0 Then
                        Insertion = "UPDATE PARACOMMERCIAL SET DosSauve='" & DataListeIntegrer.Rows(n).Cells("dosdest1").Value & "',Fournisseur='" & DataListeIntegrer.Rows(n).Cells("Fournisseur1").Value & "'where BaseCpta='" & DataListeIntegrer.Rows(n).Cells("basecpta1").Value & "'and BaseCial='" & DataListeIntegrer.Rows(n).Cells("basegescom1").Value & "'and Piece='" & DataListeIntegrer.Rows(n).Cells("Piece1").Value & "'and DosTraite='" & DataListeIntegrer.Rows(n).Cells("dostraite1").Value & "'"
                        OleCommandEnreg = New OleDbCommand(Insertion)
                        OleCommandEnreg.Connection = OleConnenection
                        OleCommandEnreg.ExecuteNonQuery()
                        Insert = True
                    End If
                Next n
                If Insert = True Then
                    MsgBox("Mise à Jour Reussie", MsgBoxStyle.Information, "Mise à Jour des pieces Commerciales")
                End If
            End If
            AfficheSchemasConso()
        Catch ex As Exception

        End Try

    End Sub
    Private Sub BT_Quit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Quit.Click
        Me.Close()
    End Sub
    Private Sub BT_Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Save.Click
        EnregistrerLeSchema()
        AfficheSchemasConso()
    End Sub
    Private Sub BT_DelRow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_DelRow.Click
        Dim first As Integer
        Dim last As Integer
        first = DataListeSchema.Rows.GetFirstRow(DataGridViewElementStates.Displayed)
        last = DataListeSchema.Rows.GetLastRow(DataGridViewElementStates.Displayed)
        If last >= 0 Then
            If last - first >= 0 Then
                DataListeSchema.Rows.RemoveAt(DataListeSchema.CurrentRow.Index)
            End If
        End If
        For i As Integer = 0 To DataListeSchema.RowCount - 1
            DataListeSchema.Rows(i).Cells("IDDossier").Value = RenvoieID("PARACOMMERCIAL") + i
        Next i
    End Sub
    Private Sub BT_ADD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_ADD.Click
        Dim i As Integer = DataListeSchema.Rows.Add()
        DataListeSchema.Rows(i).Cells("IDDossier").Value = RenvoieID("PARACOMMERCIAL") + i
        If RenvoieID("PARACOMMERCIAL") = DataListeSchema.Rows(0).Cells("IDDossier").Value Then

        Else
            For j As Integer = 0 To DataListeSchema.RowCount - 1
                DataListeSchema.Rows(j).Cells("IDDossier").Value = RenvoieID("PARACOMMERCIAL") + j
            Next j
        End If
    End Sub

    Private Sub Parametre_Piece_Commerciale_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized
        GroupBox4.Width = Me.Width
        Dim boll As Boolean
        boll = Connected()
        AfficheSchemasConso()
        AfficheSocieteCpta()
        AfficheSocieteCial()

    End Sub

    Private Sub DataListeSchema_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataListeSchema.CellClick
        Dim NameFormat As String
        Try
            If e.RowIndex >= 0 Then
                If DataListeSchema.Columns(e.ColumnIndex).Name = "par3" Then
                    FolderRepListeFile.Description = "Chemin du Dossier de Traitement"
                    If FolderRepListeFile.ShowDialog = Windows.Forms.DialogResult.OK Then
                        NameFormat = Trim(FolderRepListeFile.SelectedPath)
                        Do While InStr(Trim(NameFormat), "\") <> 0
                            NameFormat = Strings.Right(NameFormat, Strings.Len(Trim(NameFormat)) - InStr(Trim(NameFormat), "\"))
                        Loop
                        DataListeSchema.Rows(e.RowIndex).Cells("dostraite").Value = FolderRepListeFile.SelectedPath & "\"
                        '    DataListeSchema.Rows(e.RowIndex).Cells("DossierExport").Value = NameFormat
                    End If
                End If

                If DataListeSchema.Columns(e.ColumnIndex).Name = "Ping" Then
                    Dim strUrlServeur, resultat As String
                    strUrlServeur = DataListeSchema.Rows(e.RowIndex).Cells("ServeurFtp").Value
                    resultat = getPingTime(strUrlServeur)
                    MsgBox(resultat, MsgBoxStyle.Information, "ping du serveur " & strUrlServeur)
                End If
            End If

            If DataListeSchema.Columns(e.ColumnIndex).Name = "Par4" Then
                FolderRepListeFile.Description = "Chemin du Dossier de Sauvegarde"
                If FolderRepListeFile.ShowDialog = Windows.Forms.DialogResult.OK Then
                    NameFormat = Trim(FolderRepListeFile.SelectedPath)
                    Do While InStr(Trim(NameFormat), "\") <> 0
                        NameFormat = Strings.Right(NameFormat, Strings.Len(Trim(NameFormat)) - InStr(Trim(NameFormat), "\"))
                    Loop
                    DataListeSchema.Rows(e.RowIndex).Cells("dosdest").Value = FolderRepListeFile.SelectedPath & "\"
                    '    DataListeSchema.Rows(e.RowIndex).Cells("DossierExport").Value = NameFormat
                End If
            End If



        Catch ex As Exception

        End Try
    End Sub

   


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        MiseàjourLeSchema()
    End Sub

    Private Sub DataListeIntegrer_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataListeIntegrer.CellClick

        Dim NameFormat As String
        Try
            If e.RowIndex >= 0 Then

                If DataListeIntegrer.Columns(e.ColumnIndex).Name = "par2" Then
                    FolderRepListeFile.Description = "Chemin du Dossier de Sauvegarde"
                    If FolderRepListeFile.ShowDialog = Windows.Forms.DialogResult.OK Then
                        NameFormat = Trim(FolderRepListeFile.SelectedPath)
                        Do While InStr(Trim(NameFormat), "\") <> 0
                            NameFormat = Strings.Right(NameFormat, Strings.Len(Trim(NameFormat)) - InStr(Trim(NameFormat), "\"))
                        Loop
                        DataListeIntegrer.Rows(e.RowIndex).Cells("dosdest1").Value = FolderRepListeFile.SelectedPath & "\"
                        '    DataListeSchema.Rows(e.RowIndex).Cells("DossierExport").Value = NameFormat
                    End If
                End If
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub BT_Delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Delete.Click
        Delete_DataListeSch()
    End Sub

    Private Function getPingTime(ByVal adresseIP As String) As String
        Dim monPing As New Ping
        Dim maReponsePing As PingReply
        Dim resultatPing As String = Nothing
        Try
            maReponsePing = monPing.Send(adresseIP, Nothing)
            resultatPing = "Réponse de " & adresseIP & " en " & maReponsePing.RoundtripTime.ToString & " ms."
            Return resultatPing
        Catch ex As PingException
            resultatPing = "Impossible de joindre l'hôte : " & ex.Message
            Return resultatPing
        End Try
    End Function
End Class