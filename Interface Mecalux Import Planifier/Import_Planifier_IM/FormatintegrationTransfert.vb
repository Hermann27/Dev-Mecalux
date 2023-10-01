Imports Objets100Lib
Imports System
Imports System.Data.OleDb
Imports System.Collections
Imports System.Windows.Forms
Imports System.IO
Imports System.Xml
Imports Microsoft.VisualBasic
Public Class FormatintegrationTransfert
    Public CellCol As Integer
    Public Index As Integer
    Public xdoc As XmlDocument
    Public racine As XmlElement
    Public nodelist As XmlNodeList
    Public nodelist2 As XmlNodeList
    Private Sub AfficheColDispo(ByRef Fichier As String)
        Dim i As Integer
        Dim Documentvente As Object
        Dim OleAdaptater As OleDbDataAdapter
        Dim OleAfficheDataset As DataSet
        Dim Oledatable As DataTable
        DataDispo.Rows.Clear()
        If Trim(Fichier) = "Document Ligne" Then
            Documentvente = "F_DOCLIGNE"
        Else
            Documentvente = "F_DOCENTETE"
        End If
        Try
            OleAdaptater = New OleDbDataAdapter("select * from WIT_COL  Where Fichier='" & Trim(Documentvente) & "' Order BY ColDispo,Libelle", OleConnenection)
            OleAfficheDataset = New DataSet
            OleAdaptater.Fill(OleAfficheDataset)
            Oledatable = OleAfficheDataset.Tables(0)
            DataDispo.RowCount = Oledatable.Rows.Count
            For i = 0 To Oledatable.Rows.Count - 1
                DataDispo.Rows(i).Cells("ColDispos").Value = Oledatable.Rows(i).Item("ColDispo")
                DataDispo.Rows(i).Cells("LibelleDispos").Value = Oledatable.Rows(i).Item("Libelle")
                DataDispo.Rows(i).Cells("Fiche").Value = Oledatable.Rows(i).Item("Fichier")
            Next i
        Catch ex As Exception

        End Try
    End Sub
    Private Sub CrerUnFormatColParametre()
        Dim j As Integer
        Dim OleAdaptaterFormat As OleDbDataAdapter
        Dim OleFormatDataset As DataSet
        Dim OledatableFormat As DataTable
        Dim OleCommandSave As OleDbCommand
        SaveFileXml.Filter = "Fichier Xml (*.Xml)|*.Xml"
        SaveFileXml.Title = "Enregistrer le Format d'integration"
        SaveFileXml.InitialDirectory = PathsFileFormatiers
        SaveFileXml.FileName = Cmb_Format.Text
        Try
            If SaveFileXml.ShowDialog = Windows.Forms.DialogResult.OK Then
                If SaveFileXml.FileName <> "" Then
                    OleAdaptaterFormat = New OleDbDataAdapter("select * from WIT_FORMAT where (NomFormat='" & Trim(System.IO.Path.GetFileName(SaveFileXml.FileName)) & "')", OleConnenection)
                    OleFormatDataset = New DataSet
                    OleAdaptaterFormat.Fill(OleFormatDataset)
                    OledatableFormat = OleFormatDataset.Tables(0)
                    If OledatableFormat.Rows.Count = 0 Then
                        SaveFileXml.CheckPathExists = True
                        Dim Filexml As New XmlTextWriter(SaveFileXml.FileName, Nothing)
                        Filexml.WriteStartDocument()
                        Filexml.WriteStartElement("FORMAT_IMPORT_TRSFERT")
                        For j = 0 To DataSelect.RowCount - 1
                            Filexml.WriteStartElement("ColUse")
                            Filexml.WriteString(CStr(DataSelect.Rows(j).Cells("Selection").Value))
                            Filexml.WriteEndElement()

                            Filexml.WriteStartElement("Libelle")
                            Filexml.WriteString(CStr(DataSelect.Rows(j).Cells("Libelles").Value))
                            Filexml.WriteEndElement()
                            If IsNumeric(DataSelect.Rows(j).Cells("Position").Value) Then
                                Filexml.WriteStartElement("iPosLeft")
                                Filexml.WriteString(CInt(DataSelect.Rows(j).Cells("Position").Value))
                                Filexml.WriteEndElement()
                            Else
                                Filexml.WriteStartElement("iPosLeft")
                                Filexml.WriteString(CInt("0"))
                                Filexml.WriteEndElement()
                            End If
                            If DataSelect.Rows(j).Cells("Infos").Value = True Then
                                Filexml.WriteStartElement("Info")
                                Filexml.WriteString("oui")
                                Filexml.WriteEndElement()
                            Else
                                Filexml.WriteStartElement("Info")
                                Filexml.WriteString("non")
                                Filexml.WriteEndElement()
                            End If

                            Filexml.WriteStartElement("Fichier")
                            Filexml.WriteString(DataSelect.Rows(j).Cells("Fichier").Value)
                            Filexml.WriteEndElement()

                            If DataSelect.Rows(j).Cells("Piece").Value = True Then
                                Filexml.WriteStartElement("Piece")
                                Filexml.WriteString("oui")
                                Filexml.WriteEndElement()
                            Else
                                Filexml.WriteStartElement("Piece")
                                Filexml.WriteString("non")
                                Filexml.WriteEndElement()
                            End If

                            If DataSelect.Rows(j).Cells("Article").Value = True Then
                                Filexml.WriteStartElement("Article")
                                Filexml.WriteString("oui")
                                Filexml.WriteEndElement()
                            Else
                                Filexml.WriteStartElement("Article")
                                Filexml.WriteString("non")
                                Filexml.WriteEndElement()
                            End If

                            If DataSelect.Rows(j).Cells("Defauts").Value = True Then
                                Filexml.WriteStartElement("Defaut")
                                Filexml.WriteString("1")
                                Filexml.WriteEndElement()
                            Else
                                Filexml.WriteStartElement("Defaut")
                                Filexml.WriteString("0")
                                Filexml.WriteEndElement()
                            End If
                            Filexml.WriteStartElement("ValeurDefaut")
                            Filexml.WriteString(DataSelect.Rows(j).Cells("ValeurDefauts").Value)
                            Filexml.WriteEndElement()

                            Filexml.WriteStartElement("Decalage")
                            Filexml.WriteString(NumUpDown.Value)
                            Filexml.WriteEndElement()

                            Filexml.WriteStartElement("TypeFormat")
                            Filexml.WriteString(Renvoietypeformat(Trim(Txtype.Text)))
                            Filexml.WriteEndElement()

                            Filexml.WriteStartElement("MODE_FORMAT")
                            Filexml.WriteString(Trim(CbMod.Text))
                            Filexml.WriteEndElement()

                            Filexml.WriteStartElement("DateFormat")
                            Filexml.WriteString(Trim(Cb_Date.Text))
                            Filexml.WriteEndElement()

                            If Ckauto.Checked = True Then
                                Filexml.WriteStartElement("PieceAuto")
                                Filexml.WriteString("oui")
                                Filexml.WriteEndElement()
                            Else
                                Filexml.WriteStartElement("PieceAuto")
                                Filexml.WriteString("non")
                                Filexml.WriteEndElement()
                            End If

                            If CkPunitaire.Checked = True Then
                                Filexml.WriteStartElement("PUDefaut")
                                Filexml.WriteString("oui")
                                Filexml.WriteEndElement()
                            Else
                                Filexml.WriteStartElement("PUDefaut")
                                Filexml.WriteString("non")
                                Filexml.WriteEndElement()
                            End If

                            Filexml.WriteStartElement("IMPORT")
                            Filexml.WriteString(Trim(CbSaisie.Text))
                            Filexml.WriteEndElement()
                        Next j
                        Filexml.WriteEndElement()
                        Filexml.Close()
                        OleCommandSave = New OleDbCommand("Insert Into WIT_FORMAT (Chemin,NomFormat,Type) VALUES ('" & SaveFileXml.FileName & "','" & Trim(System.IO.Path.GetFileName(SaveFileXml.FileName)) & "','" & Renvoietypeformat(Trim(Txtype.Text)) & "')")
                        OleCommandSave.Connection = OleConnenection
                        OleCommandSave.ExecuteNonQuery()
                    Else
                        If OledatableFormat.Rows.Count <> 0 And SaveFileXml.FileName = OledatableFormat.Rows(0).Item("Chemin") Then
                            SaveFileXml.CheckPathExists = True
                            File.Delete(SaveFileXml.FileName)
                            Dim Filexml As New XmlTextWriter(SaveFileXml.FileName, Nothing)
                            Filexml.WriteStartDocument()
                            Filexml.WriteStartElement("FORMAT_IMPORT_TRSFERT")
                            For j = 0 To DataSelect.RowCount - 1

                                Filexml.WriteStartElement("ColUse")
                                Filexml.WriteString(CStr(DataSelect.Rows(j).Cells("Selection").Value))
                                Filexml.WriteEndElement()

                                Filexml.WriteStartElement("Libelle")
                                Filexml.WriteString(CStr(DataSelect.Rows(j).Cells("Libelles").Value))
                                Filexml.WriteEndElement()
                                If IsNumeric(DataSelect.Rows(j).Cells("Position").Value) Then
                                    Filexml.WriteStartElement("iPosLeft")
                                    Filexml.WriteString(CInt(DataSelect.Rows(j).Cells("Position").Value))
                                    Filexml.WriteEndElement()
                                Else
                                    Filexml.WriteStartElement("iPosLeft")
                                    Filexml.WriteString(CInt("0"))
                                    Filexml.WriteEndElement()
                                End If
                                If DataSelect.Rows(j).Cells("Infos").Value = True Then
                                    Filexml.WriteStartElement("Info")
                                    Filexml.WriteString("oui")
                                    Filexml.WriteEndElement()
                                Else
                                    Filexml.WriteStartElement("Info")
                                    Filexml.WriteString("non")
                                    Filexml.WriteEndElement()
                                End If

                                Filexml.WriteStartElement("Fichier")
                                Filexml.WriteString(DataSelect.Rows(j).Cells("Fichier").Value)
                                Filexml.WriteEndElement()

                                If DataSelect.Rows(j).Cells("Piece").Value = True Then
                                    Filexml.WriteStartElement("Piece")
                                    Filexml.WriteString("oui")
                                    Filexml.WriteEndElement()
                                Else
                                    Filexml.WriteStartElement("Piece")
                                    Filexml.WriteString("non")
                                    Filexml.WriteEndElement()
                                End If

                                If DataSelect.Rows(j).Cells("Article").Value = True Then
                                    Filexml.WriteStartElement("Article")
                                    Filexml.WriteString("oui")
                                    Filexml.WriteEndElement()
                                Else
                                    Filexml.WriteStartElement("Article")
                                    Filexml.WriteString("non")
                                    Filexml.WriteEndElement()
                                End If

                                If DataSelect.Rows(j).Cells("Defauts").Value = True Then
                                    Filexml.WriteStartElement("Defaut")
                                    Filexml.WriteString("1")
                                    Filexml.WriteEndElement()
                                Else
                                    Filexml.WriteStartElement("Defaut")
                                    Filexml.WriteString("0")
                                    Filexml.WriteEndElement()
                                End If

                                Filexml.WriteStartElement("ValeurDefaut")
                                Filexml.WriteString(DataSelect.Rows(j).Cells("ValeurDefauts").Value)
                                Filexml.WriteEndElement()


                                Filexml.WriteStartElement("Decalage")
                                Filexml.WriteString(NumUpDown.Value)
                                Filexml.WriteEndElement()

                                Filexml.WriteStartElement("TypeFormat")
                                Filexml.WriteString(Renvoietypeformat(Trim(Txtype.Text)))
                                Filexml.WriteEndElement()

                                Filexml.WriteStartElement("MODE_FORMAT")
                                Filexml.WriteString(Trim(CbMod.Text))
                                Filexml.WriteEndElement()

                                Filexml.WriteStartElement("DateFormat")
                                Filexml.WriteString(Trim(Cb_Date.Text))
                                Filexml.WriteEndElement()

                                If Ckauto.Checked = True Then
                                    Filexml.WriteStartElement("PieceAuto")
                                    Filexml.WriteString("oui")
                                    Filexml.WriteEndElement()
                                Else
                                    Filexml.WriteStartElement("PieceAuto")
                                    Filexml.WriteString("non")
                                    Filexml.WriteEndElement()
                                End If

                                If CkPunitaire.Checked = True Then
                                    Filexml.WriteStartElement("PUDefaut")
                                    Filexml.WriteString("oui")
                                    Filexml.WriteEndElement()
                                Else
                                    Filexml.WriteStartElement("PUDefaut")
                                    Filexml.WriteString("non")
                                    Filexml.WriteEndElement()
                                End If

                                Filexml.WriteStartElement("IMPORT")
                                Filexml.WriteString(Trim(CbSaisie.Text))
                                Filexml.WriteEndElement()
                            Next j
                            Filexml.WriteEndElement()
                            Filexml.Close()
                            OleCommandSave = New OleDbCommand("UPDATE  WIT_FORMAT SET Type='" & Renvoietypeformat(Trim(Txtype.Text)) & "' where (NomFormat='" & Trim(System.IO.Path.GetFileName(SaveFileXml.FileName)) & "')")
                            OleCommandSave.Connection = OleConnenection
                            OleCommandSave.ExecuteNonQuery()
                            OleCommandSave = New OleDbCommand("UPDATE  WIT_SCHEMA SET Type='" & Renvoietypeformat(Trim(Txtype.Text)) & "',Mode='" & Trim(CbMod.Text) & "' where (NomFormat='" & Trim(System.IO.Path.GetFileName(SaveFileXml.FileName)) & "')")
                            OleCommandSave.Connection = OleConnenection
                            OleCommandSave.ExecuteNonQuery()

                        Else
                            MsgBox("Le Fichier " & System.IO.Path.GetFileName(SaveFileXml.FileName) & " existe déja! Duplication Impossible", MsgBoxStyle.Information, "Format d'integration")
                        End If
                    End If
                Else
                    MsgBox("Saisir un Nom de Fichier", MsgBoxStyle.Information, "Format d'integration")
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub AfficheFormat()
        Dim OleAdaptaterAfficheFormat As OleDbDataAdapter
        Dim OleAfficheFormatDataset As DataSet
        Dim OledatableAfficheFormat As DataTable
        Dim i As Integer
        Try
            Cmb_Format.Items.Clear()
            OleAdaptaterAfficheFormat = New OleDbDataAdapter("select * from WIT_FORMAT WHERE Type <> 'Longueur Fixe'", OleConnenection)
            OleAfficheFormatDataset = New DataSet
            OleAdaptaterAfficheFormat.Fill(OleAfficheFormatDataset)
            OledatableAfficheFormat = OleAfficheFormatDataset.Tables(0)
            For i = 0 To OledatableAfficheFormat.Rows.Count - 1
                If Trim(OledatableAfficheFormat.Rows(i).Item("Chemin")) <> "" Then
                    Cmb_Format.Text = OledatableAfficheFormat.Rows(i).Item("NomFormat")
                End If
                If Trim(OledatableAfficheFormat.Rows(i).Item("Chemin")) <> "" Then
                    Cmb_Format.Items.Add(OledatableAfficheFormat.Rows(i).Item("NomFormat"))
                End If
            Next i
        Catch ex As Exception

        End Try
    End Sub
    Private Sub DeleteFormat()
        Dim DelFormat As String
        Dim i As Integer
        Dim Delschema As String
        Dim OleAdaptaterDeleteFormat As OleDbDataAdapter
        Dim OleDeleteFormatDataset As DataSet
        Dim OledatableDeleteFormat As DataTable
        Dim OleComDeleteFor As OleDbCommand
        Try
            If File.Exists(Txt_Chemin.Text) = True Then
                DelFormat = "Delete * from WIT_FORMAT where Chemin='" & Txt_Chemin.Text & "' And Type='" & Renvoietypeformat(Trim(Txtype.Text)) & "'"
                OleComDeleteFor = New OleDbCommand(DelFormat)
                OleComDeleteFor.Connection = OleConnenection
                OleComDeleteFor.ExecuteNonQuery()
                File.Delete(Txt_Chemin.Text)
                OleAdaptaterDeleteFormat = New OleDbDataAdapter("select * from WIT_SCHEMA where CheminFormat='" & Txt_Chemin.Text & "' And Type='" & Renvoietypeformat(Trim(Txtype.Text)) & "'", OleConnenection)
                OleDeleteFormatDataset = New DataSet
                OleAdaptaterDeleteFormat.Fill(OleDeleteFormatDataset)
                OledatableDeleteFormat = OleDeleteFormatDataset.Tables(0)
                If OledatableDeleteFormat.Rows.Count <> 0 Then
                    For i = 0 To OledatableDeleteFormat.Rows.Count - 1
                        Delschema = "delete from WIT_SCHEMA where CheminFormat='" & Txt_Chemin.Text & "' and CheminFilexport='" & OledatableDeleteFormat.Rows(i).Item("CheminFilexport") & "' And Type='" & Renvoietypeformat(Trim(Txtype.Text)) & "'"
                        OleComDeleteFor = New OleDbCommand(Delschema)
                        OleComDeleteFor.Connection = OleConnenection
                        OleComDeleteFor.ExecuteNonQuery()
                    Next i
                End If
                AfficheFormat()
                Cmb_Format.Text = ""
                Txt_Chemin.Text = ""
                AffichFormatModifiable()
            Else
                If MsgBox("Chemin du Fichier inexistant!", MsgBoxStyle.OkCancel, "Format d'integration") = MsgBoxResult.Ok Then
                    DelFormat = "Delete * from WIT_FORMAT where Chemin='" & Txt_Chemin.Text & "' And Type='" & Renvoietypeformat(Trim(Txtype.Text)) & "'"
                    OleComDeleteFor = New OleDbCommand(DelFormat)
                    OleComDeleteFor.Connection = OleConnenection
                    OleComDeleteFor.ExecuteNonQuery()
                    OleAdaptaterDeleteFormat = New OleDbDataAdapter("select * from WIT_SCHEMA where CheminFormat='" & Txt_Chemin.Text & "' And Type='" & Renvoietypeformat(Trim(Txtype.Text)) & "'", OleConnenection)
                    OleDeleteFormatDataset = New DataSet
                    OleAdaptaterDeleteFormat.Fill(OleDeleteFormatDataset)
                    OledatableDeleteFormat = OleDeleteFormatDataset.Tables(0)
                    If OledatableDeleteFormat.Rows.Count <> 0 Then
                        For i = 0 To OledatableDeleteFormat.Rows.Count - 1
                            Delschema = "delete from WIT_SCHEMA where CheminFormat='" & Txt_Chemin.Text & "' and CheminFilexport='" & OledatableDeleteFormat.Rows(i).Item("CheminFilexport") & "' And Type='" & Renvoietypeformat(Trim(Txtype.Text)) & "'"
                            OleComDeleteFor = New OleDbCommand(Delschema)
                            OleComDeleteFor.Connection = OleConnenection
                            OleComDeleteFor.ExecuteNonQuery()
                        Next i
                    End If
                    AfficheFormat()
                    Cmb_Format.Text = ""
                    Txt_Chemin.Text = ""
                    AffichFormatModifiable()
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub AffichFormatModifiable()
        Dim NomColonne As String
        Dim NomEntete As String
        Dim typedeformat As Object = Nothing
        Dim typeImport As Object = Nothing
        Dim ModeFormat As Object = Nothing
        Dim DateFormat As Object = Nothing
        Dim PosLeft As Integer
        Dim poslongueur, PieceModifier, ValeurDefaut, Defaut, SageFichier As String
        Dim Decal, PieceAuto, Punitaire, LigneArticle As Object
        Dim i As Integer
        DataSelect.Rows.Clear()
        Try
            If Trim(Txt_Chemin.Text) <> "" Then
                If File.Exists(Txt_Chemin.Text) = True Then
                    Dim FileXml As New XmlTextReader(Trim(Txt_Chemin.Text))
                    i = 0
                    While (FileXml.Read())
                        If FileXml.LocalName = "ColUse" Then
                            NomColonne = FileXml.ReadString

                            FileXml.Read()
                            NomEntete = FileXml.ReadString

                            FileXml.Read()
                            PosLeft = FileXml.ReadString

                            FileXml.Read()
                            poslongueur = FileXml.ReadString

                            FileXml.Read()
                            SageFichier = FileXml.ReadString

                            FileXml.Read()
                            PieceModifier = FileXml.ReadString

                            FileXml.Read()
                            LigneArticle = FileXml.ReadString

                            FileXml.Read()
                            Defaut = FileXml.ReadString

                            FileXml.Read()
                            ValeurDefaut = FileXml.ReadString

                            FileXml.Read()
                            Decal = FileXml.ReadString

                            FileXml.Read()
                            typedeformat = FileXml.ReadString

                            FileXml.Read()
                            ModeFormat = FileXml.ReadString

                            FileXml.Read()
                            DateFormat = FileXml.ReadString

                            FileXml.Read()
                            PieceAuto = FileXml.ReadString

                            FileXml.Read()
                            Punitaire = FileXml.ReadString

                            FileXml.Read()
                            typeImport = FileXml.ReadString

                            If (NomColonne <> "" And NomEntete <> "") Then
                                DataSelect.RowCount = i + 1
                                DataSelect.Rows(i).Cells("Libelles").Value = NomEntete
                                DataSelect.Rows(i).Cells("Selection").Value = NomColonne
                                DataSelect.Rows(i).Cells("Fichier").Value = SageFichier

                                If LigneArticle = "oui" Then
                                    DataSelect.Rows(i).Cells("Article").Value = True
                                Else
                                    DataSelect.Rows(i).Cells("Article").Value = False
                                End If

                                If PieceModifier = "oui" Then
                                    DataSelect.Rows(i).Cells("Piece").Value = True
                                Else
                                    DataSelect.Rows(i).Cells("Piece").Value = False
                                End If
                                If Defaut = "0" Then
                                    DataSelect.Rows(i).Cells("Position").Value = PosLeft
                                    If poslongueur = "oui" Then
                                        DataSelect.Rows(i).Cells("Infos").Value = True
                                        DataSelect.Rows(i).Cells("Defauts").Value = False
                                        DataSelect.Rows(i).Cells("ValeurDefauts").ReadOnly = True
                                    Else
                                        DataSelect.Rows(i).Cells("Infos").Value = False
                                        DataSelect.Rows(i).Cells("Defauts").Value = False
                                        DataSelect.Rows(i).Cells("ValeurDefauts").ReadOnly = True
                                    End If
                                Else
                                    If poslongueur = "oui" Then
                                        DataSelect.Rows(i).Cells("Infos").Value = True
                                    Else
                                        DataSelect.Rows(i).Cells("Infos").Value = False
                                    End If
                                    DataSelect.Rows(i).Cells("Defauts").Value = True
                                    DataSelect.Rows(i).Cells("Position").ReadOnly = True
                                    DataSelect.Rows(i).Cells("ValeurDefauts").Value = ValeurDefaut
                                End If
                                If PieceModifier = "oui" Then
                                    DataSelect.Rows(i).Cells("Piece").Value = True
                                Else
                                    DataSelect.Rows(i).Cells("Piece").Value = False
                                End If
                                If PieceAuto = "oui" Then
                                    Ckauto.Checked = True
                                Else
                                    Ckauto.Checked = False
                                End If
                                If Punitaire = "oui" Then
                                    CkPunitaire.Checked = True
                                Else
                                    CkPunitaire.Checked = False
                                End If
                                NumUpDown.Value = Decal
                                i = i + 1
                            End If
                        End If
                    End While
                    Me.Text = "Formats d'integration des documents de Transfert " & Afficheauuser(typedeformat)
                    Txtype.Text = Afficheauuser(typedeformat)
                    CbMod.Text = ModeFormat
                    Cb_Date.Text = DateFormat
                    CbSaisie.Text = typeImport
                    FileXml.Close()
                Else
                    MsgBox("Nom du Format inexistant!", MsgBoxStyle.Information, "Format d'integration")
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub MonterLigne()
        Try
            If Index >= 1 Then
                DataSelect.Rows.Insert(Index - 1, DataSelect.Rows(Index).Cells("Selection").Value, DataSelect.Rows(Index).Cells("Position").Value, DataSelect.Rows(Index).Cells("Infos").Value, DataSelect.Rows(Index).Cells("Fichier").Value, DataSelect.Rows(Index).Cells("Piece").Value, DataSelect.Rows(Index).Cells("Article").Value, DataSelect.Rows(Index).Cells("Defauts").Value, DataSelect.Rows(Index).Cells("ValeurDefauts").Value, DataSelect.Rows(Index).Cells("Libelles").Value)
                DataSelect.Rows.RemoveAt(Index + 1)
                DataSelect.Rows(Index - 1).Selected = True
                If Index >= 1 Then
                    Index = Index - 1
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub DescendreLigne()
        Try
            If Index < DataSelect.RowCount - 1 Then
                DataSelect.Rows.Insert(Index, DataSelect.Rows(Index + 1).Cells("Selection").Value, DataSelect.Rows(Index + 1).Cells("Position").Value, DataSelect.Rows(Index + 1).Cells("Infos").Value, DataSelect.Rows(Index + 1).Cells("Fichier").Value, DataSelect.Rows(Index + 1).Cells("Piece").Value, DataSelect.Rows(Index + 1).Cells("Article").Value, DataSelect.Rows(Index + 1).Cells("Defauts").Value, DataSelect.Rows(Index + 1).Cells("ValeurDefauts").Value, DataSelect.Rows(Index + 1).Cells("Libelles").Value)
                DataSelect.Rows.RemoveAt(Index + 2)
                DataSelect.Rows(Index + 1).Selected = True
                If Index < DataSelect.RowCount - 1 Then
                    Index = Index + 1
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub DeleteColDispo()
        Dim i As Integer
        Dim Colbool As Boolean = False
        Dim OleAdaptaterDeleteDispo As OleDbDataAdapter
        Dim OleDeleteDispoDataset As DataSet
        Dim OledatableDeleteDispo As DataTable
        Try
            For i = 0 To DataSelect.RowCount - 1
                If DataDispo.Rows(CellCol).Cells("ColDispos").Value = DataSelect.Rows(i).Cells("Selection").Value And DataDispo.Rows(CellCol).Cells("Fiche").Value = DataSelect.Rows(i).Cells("Fichier").Value Then
                    Colbool = True
                Else
                End If
            Next i
            If Colbool = False Then
                OleAdaptaterDeleteDispo = New OleDbDataAdapter("select * from WIT_COL where ColDispo='" & DataDispo.Rows(CellCol).Cells("ColDispos").Value & "'", OleConnenection)
                OleDeleteDispoDataset = New DataSet
                OleAdaptaterDeleteDispo.Fill(OleDeleteDispoDataset)
                OledatableDeleteDispo = OleDeleteDispoDataset.Tables(0)
                If OledatableDeleteDispo.Rows.Count <> 0 Then
                    If OledatableDeleteDispo.Rows(0).Item("Libre") = True Then
                        DataSelect.Rows.Add(DataDispo.Rows(CellCol).Cells("ColDispos").Value, Nothing, True, DataDispo.Rows(CellCol).Cells("Fiche").Value, Nothing, Nothing, Nothing, Nothing, DataDispo.Rows(CellCol).Cells("LibelleDispos").Value)
                    Else
                        If OledatableDeleteDispo.Rows(0).Item("InfoLigne") = True Then
                            DataSelect.Rows.Add(DataDispo.Rows(CellCol).Cells("ColDispos").Value, Nothing, True, DataDispo.Rows(CellCol).Cells("Fiche").Value, Nothing, Nothing, Nothing, Nothing, DataDispo.Rows(CellCol).Cells("LibelleDispos").Value)
                        Else
                            DataSelect.Rows.Add(DataDispo.Rows(CellCol).Cells("ColDispos").Value, Nothing, False, DataDispo.Rows(CellCol).Cells("Fiche").Value, Nothing, Nothing, Nothing, Nothing, DataDispo.Rows(CellCol).Cells("LibelleDispos").Value)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub DeleteColSelect()
        Dim first As Integer
        Dim last As Integer
        Try
            first = DataSelect.Rows.GetFirstRow(DataGridViewElementStates.Displayed)
            last = DataSelect.Rows.GetLastRow(DataGridViewElementStates.Displayed)
            If last >= 0 Then
                If last - first >= 0 Then
                    DataSelect.Rows.RemoveAt(DataSelect.CurrentRow.Index)
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub FormatintegrationTransfert_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If Connected() = True Then
                Ckauto.Checked = True
                CbFichier.Text = "Document Entête"
                AfficheColDispo(CbFichier.Text)
                AfficheFormat()
            End If
            Me.WindowState = FormWindowState.Maximized
        Catch ex As Exception

        End Try
    End Sub
    Private Sub Cmb_Format_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmb_Format.SelectedIndexChanged
        Dim OleAdaptaterCmb As OleDbDataAdapter
        Dim OleCmbDataset As DataSet
        Dim OledatableCmb As DataTable
        Try
            OleAdaptaterCmb = New OleDbDataAdapter("select * from WIT_FORMAT where NomFormat='" & Cmb_Format.Items.Item(Cmb_Format.SelectedIndex) & "'", OleConnenection)
            OleCmbDataset = New DataSet
            OleAdaptaterCmb.Fill(OleCmbDataset)
            OledatableCmb = OleCmbDataset.Tables(0)
            If OledatableCmb.Rows.Count <> 0 Then
                Txt_Chemin.Text = OledatableCmb.Rows(0).Item("Chemin")
            End If
            AffichFormatModifiable()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub BT_DelDispo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_DelDispo.Click
        Try
            DeleteColDispo()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub BT_DelSelt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_DelSelt.Click
        Try
            DeleteColSelect()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub BT_UP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_UP.Click
        Try
            MonterLigne()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub BT_Down_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_Down.Click
        Try
            DescendreLigne()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub BT_SaveFormat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_SaveFormat.Click
        Try
            Dim ExisteCoche As Boolean = False
            If Trim(CbMod.Text) = "Modification" Then
                Dim i As Integer
                For i = 0 To DataSelect.RowCount - 1
                    If DataSelect.Rows(i).Cells("Piece").Value = True Then
                        ExisteCoche = True
                        Exit For
                    End If
                Next i
                If ExisteCoche = True Then
                    CrerUnFormatColParametre()
                    AfficheFormat()
                Else
                    MsgBox("Aucun identifiant de N°Piece n'a été Coché", MsgBoxStyle.Information, "Enregistrer Format")
                End If
            Else
                If Trim(CbMod.Text) = "Création" Then
                    Dim i As Integer
                    For i = 0 To DataSelect.RowCount - 1
                        If DataSelect.Rows(i).Cells("Piece").Value = True Then
                            ExisteCoche = True
                            Exit For
                        End If
                    Next i
                    If ExisteCoche = True Then
                        CrerUnFormatColParametre()
                        AfficheFormat()
                    Else
                        MsgBox("Aucun identifiant de N°Piece n'a été Coché", MsgBoxStyle.Information, "Enregistrer Format")
                    End If

                End If
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Function ExisteRechercheArticle() As Boolean
        Dim i As Integer
        ExisteRechercheArticle = True
        For i = 0 To DataSelect.RowCount - 1
            If DataSelect.Rows(i).Cells("Article").Value = True Then
                If Trim(DataSelect.Rows(i).Cells("Fichier").Value) = "F_DOCLIGNE" Then
                    ExisteRechercheArticle = True
                    Exit For
                Else
                    ExisteRechercheArticle = False
                    Exit For
                End If
            End If
        Next i
    End Function
    Private Sub BT_DelForm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_DelForm.Click
        Try
            DeleteFormat()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub Cmb_Format_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Cmb_Format.TextChanged
        Dim OleAdaptaterCmb As OleDbDataAdapter
        Dim OleCmbDataset As DataSet
        Dim OledatableCmb As DataTable
        Try
            OleAdaptaterCmb = New OleDbDataAdapter("select * from WIT_FORMAT where (NomFormat='" & Trim(Cmb_Format.Text) & "')", OleConnenection)
            OleCmbDataset = New DataSet
            OleAdaptaterCmb.Fill(OleCmbDataset)
            OledatableCmb = OleCmbDataset.Tables(0)
            If OledatableCmb.Rows.Count <> 0 Then
                Txt_Chemin.Text = OledatableCmb.Rows(0).Item("Chemin")
            End If
            AffichFormatModifiable()
        Catch ex As Exception

        End Try
    End Sub
    Private Sub BT_New_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BT_New.Click
        Ckauto.Checked = True
        DataSelect.Rows.Clear()
        Txt_Chemin.Text = ""
        Cmb_Format.Text = ""
        NumUpDown.Value = 0
    End Sub
    Private Function Verificationformat(ByRef fichierformat As String) As Boolean
        Dim ExisteFormat As Boolean = False
        xdoc = New XmlDocument
        xdoc.Load(Trim(fichierformat))
        racine = xdoc.DocumentElement
        nodelist = racine.ChildNodes
        Dim i As Integer
        For i = 0 To nodelist.Count - 1
            If Trim(nodelist.ItemOf(i).Name) = "TypeFormat" Then
                If Renvoietypeformat(Trim(Txtype.Text)) = nodelist.ItemOf(i).InnerText Then
                    ExisteFormat = True
                    Exit For
                Else
                    MsgBox("Ce fichier ne correspond pas au format sélectionné!", MsgBoxStyle.Information, "Selection Fichier Format")
                    Exit For
                End If
            End If
        Next i
        Verificationformat = ExisteFormat
    End Function
    Private Sub DataSelect_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataSelect.CellClick
        Try
            If e.RowIndex >= 0 Then
                Index = e.RowIndex
                DataSelect.UpdateCellValue(e.ColumnIndex, e.RowIndex)
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub DataSelect_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataSelect.CellContentClick
        If e.RowIndex >= 0 Then
            If DataSelect.Columns(e.ColumnIndex).Name = "Piece" Then
                Dim i As Integer
                For i = 0 To DataSelect.RowCount - 1
                    DataSelect.Rows(i).Cells("Piece").Value = False
                Next i

            End If
            If DataSelect.Columns(e.ColumnIndex).Name = "Article" Then
                Dim i As Integer
                For i = 0 To DataSelect.RowCount - 1
                    DataSelect.Rows(i).Cells("Article").Value = False
                Next i
                DataSelect.UpdateCellValue(e.ColumnIndex, e.RowIndex)
                DataSelect.EndEdit()
                If DataSelect.Rows(e.RowIndex).Cells("Article").Value = True And Trim(DataSelect.Rows(e.RowIndex).Cells("Fichier").Value) = "F_DOCLIGNE" Then
                    RechercheCriteredocument.TxtFichier.Text = DataSelect.Rows(e.RowIndex).Cells("Fichier").Value
                    RechercheCriteredocument.TxtLibelle.Text = DataSelect.Rows(e.RowIndex).Cells("Selection").Value
                    RechercheCriteredocument.TxtSage.Text = DataSelect.Rows(e.RowIndex).Cells("Libelles").Value
                    RechercheCriteredocument.ShowDialog()
                Else
                    DataSelect.Rows(e.RowIndex).Cells("Article").Value = False
                End If
            End If
        End If

    End Sub
    Private Sub DataDispo_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataDispo.CellClick
        Try
            If e.RowIndex >= 0 Then
                CellCol = e.RowIndex
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub DataDispo_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataDispo.CellEndEdit
        Try
            If e.RowIndex >= 0 Then
                If DataSelect.Columns(e.ColumnIndex).Name = "Defauts" Then
                    If DataSelect.Rows(e.RowIndex).Cells("Defauts").Value = True Then
                        DataSelect.Rows(e.RowIndex).Cells("Position").Value = 0
                        DataSelect.Rows(e.RowIndex).Cells("Position").Value = 0
                        DataSelect.Rows(e.RowIndex).Cells("Position").ReadOnly = True
                        DataSelect.Rows(e.RowIndex).Cells("Position").ReadOnly = True
                    Else
                        DataSelect.Rows(e.RowIndex).Cells("Position").ReadOnly = False
                        DataSelect.Rows(e.RowIndex).Cells("Position").ReadOnly = False
                    End If
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub DataSelect_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataSelect.CellContentDoubleClick
        If e.RowIndex >= 0 Then
            If DataSelect.Columns(e.ColumnIndex).Name = "Piece" Then
                Dim i As Integer
                For i = 0 To DataSelect.RowCount - 1
                    DataSelect.Rows(i).Cells("Piece").Value = False
                Next i

            End If
            If DataSelect.Columns(e.ColumnIndex).Name = "Article" Then
                Dim i As Integer
                For i = 0 To DataSelect.RowCount - 1
                    DataSelect.Rows(i).Cells("Article").Value = False
                Next i
                DataSelect.UpdateCellValue(e.ColumnIndex, e.RowIndex)
                DataSelect.EndEdit()
                If DataSelect.Rows(e.RowIndex).Cells("Article").Value = True And Trim(DataSelect.Rows(e.RowIndex).Cells("Fichier").Value) = "F_DOCLIGNE" Then
                    RechercheCriteredocument.TxtFichier.Text = DataSelect.Rows(e.RowIndex).Cells("Fichier").Value
                    RechercheCriteredocument.TxtLibelle.Text = DataSelect.Rows(e.RowIndex).Cells("Selection").Value
                    RechercheCriteredocument.TxtSage.Text = DataSelect.Rows(e.RowIndex).Cells("Libelles").Value
                    RechercheCriteredocument.ShowDialog()
                Else
                    DataSelect.Rows(e.RowIndex).Cells("Article").Value = False
                End If
            End If
        End If
    End Sub

    Private Sub DataSelect_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataSelect.CellDoubleClick
        Try
            If e.RowIndex >= 0 Then
                Index = e.RowIndex
                DataSelect.UpdateCellValue(e.ColumnIndex, e.RowIndex)
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub DataSelect_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataSelect.CellEndEdit
        Try
            If e.RowIndex >= 0 Then
                If DataSelect.Columns(e.ColumnIndex).Name = "Defauts" Then
                    If DataSelect.Rows(e.RowIndex).Cells("Defauts").Value = True Then
                        DataSelect.Rows(e.RowIndex).Cells("Position").Value = 0
                        DataSelect.Rows(e.RowIndex).Cells("Position").ReadOnly = True
                        DataSelect.Rows(e.RowIndex).Cells("ValeurDefauts").ReadOnly = False
                    Else
                        DataSelect.Rows(e.RowIndex).Cells("Position").ReadOnly = False
                        DataSelect.Rows(e.RowIndex).Cells("ValeurDefauts").ReadOnly = True
                        DataSelect.Rows(e.RowIndex).Cells("ValeurDefauts").Value = ""
                    End If
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub CbFichier_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CbFichier.SelectedIndexChanged
        AfficheColDispo(CbFichier.SelectedItem)
    End Sub
End Class