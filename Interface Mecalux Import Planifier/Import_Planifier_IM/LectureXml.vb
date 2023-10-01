Imports System.Xml
Imports System.IO
Imports System
Imports System.Data.OleDb
Public Class LectureXml
    Dim FileName As Object
    Public xdoc As XmlDocument
    Public racine As XmlElement
    Public nodelist As XmlNodeList
    Public nodelist2 As XmlNodeList
    Public Error_Article As StreamWriter


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        OpenFileDialog1.Filter = "Fichier Access (*.xml)|*.xml"
        OpenFileDialog1.FileName = Nothing
        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            Textxml.Text = OpenFileDialog1.FileName
            FileName = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If Trim(TextFour.Text) <> "" Then
            Dim OleAdaptaterEnreg As OleDbDataAdapter
            Dim OleEnregDataset As DataSet
            Dim OledatableEnreg As DataTable
            OleAdaptaterEnreg = New OleDbDataAdapter("select * From FOURNISSEUR where  Fournisseur='" & Trim(TextFour.Text) & "'", OleConnenection)
            OleEnregDataset = New DataSet
            OleAdaptaterEnreg.Fill(OleEnregDataset)
            OledatableEnreg = OleEnregDataset.Tables(0)
            If OledatableEnreg.Rows.Count <> 0 Then
                If File.Exists(Trim(FileName)) Then
                    ReadFile(Trim(FileName), Trim(OledatableEnreg.Rows(0).Item("Dossier_Jour")))
                Else
                    MsgBox("Le Chemin du Fichier est Invalide", MsgBoxStyle.Information, "Integration des ARTICLES")
                End If
            Else
                MsgBox("Le Fournisseur est Inexistant,Créez le fournisseur", MsgBoxStyle.Information, "Integration des ARTICLES")
            End If
        Else
            MsgBox("Saisir Un Code Fournisseur Existant", MsgBoxStyle.Information, "Integration des ARTICLES")
        End If
    End Sub
    Public Sub ReadFile(ByVal nomfich As String, ByRef Filejour As String)
        Try
            Dim i, j As Integer
            Dim Libelle, CodeEan, Fournisseur As Object
            CodeEan = Nothing
            Libelle = Nothing
            Fournisseur = Nothing
            If Directory.Exists(Filejour) = True Then
                Error_Article = File.AppendText(Filejour & "Article" & Strings.Right(DateAndTime.Year(Now), 2) & "" & Format(DateAndTime.Month(Now), "00") & "" & Format(DateAndTime.Day(Now), "00") & "[" & "" & Format(DateAndTime.Hour(Now), "00") & "-" & Format(DateAndTime.Minute(Now), "00") & "-" & Format(DateAndTime.Second(Now), "00") & "]")
            Else
                Error_Article = File.AppendText("C:\Article" & Strings.Right(DateAndTime.Year(Now), 2) & "" & Format(DateAndTime.Month(Now), "00") & "" & Format(DateAndTime.Day(Now), "00") & "[" & "" & Format(DateAndTime.Hour(Now), "00") & "-" & Format(DateAndTime.Minute(Now), "00") & "-" & Format(DateAndTime.Second(Now), "00") & "]")
            End If
            Error_Article.WriteLine("Traitement du Fichier " & nomfich)
            xdoc = New XmlDocument
            xdoc.Load(nomfich)
            racine = xdoc.DocumentElement
            nodelist = racine.ChildNodes
            ProgressBar1.Value = ProgressBar1.Minimum
            ProgressBar1.Maximum = nodelist.Count + 1
            For i = 0 To nodelist.Count - 1
                nodelist2 = racine.ChildNodes(i).ChildNodes
                For j = 0 To nodelist2.Count - 1
                    If nodelist2.ItemOf(j).Name = "CH01101" Then
                        Libelle = nodelist2.ItemOf(j).InnerText
                    End If
                    If nodelist2.ItemOf(j).Name = "CH01102" Then
                        CodeEan = nodelist2.ItemOf(j).InnerText
                    End If
                    If nodelist2.ItemOf(j).Name = "CH01121" Then
                        Fournisseur = nodelist2.ItemOf(j).InnerText
                    End If

                Next j
                If Trim(TextFour.Text) <> "" And Trim(CodeEan) <> "" Then
                    EnregistrerLeSchema(CodeEan, Trim(Trim(Fournisseur) & "" & Trim(Libelle)), Trim(TextFour.Text))
                    CodeEan = Nothing
                    Libelle = Nothing
                    Fournisseur = Nothing
                End If
                ProgressBar1.Value = ProgressBar1.Value + 1
            Next i
            Error_Article.Close()
        Catch ex As Exception

        End Try

    End Sub
    Private Sub EnregistrerLeSchema(ByRef Codeartfour As Object, ByRef Codeartdistri As Object, ByRef Codefour As Object)
        Dim OleAdaptaterEnreg As OleDbDataAdapter
        Dim OleEnregDataset As DataSet
        Dim OledatableEnreg As DataTable
        Dim OleCommandEnreg As OleDbCommand
        Dim Insertion As String
        If Trim(Codeartfour) <> "" And Trim(Codefour) <> "" Then
            OleAdaptaterEnreg = New OleDbDataAdapter("select * From ARTICLE WHERE  Fournisseur='" & Trim(Codefour) & "'and Code_Article_Fo='" & Trim(Codeartfour) & "'", OleConnenection)
            OleEnregDataset = New DataSet
            OleAdaptaterEnreg.Fill(OleEnregDataset)
            OledatableEnreg = OleEnregDataset.Tables(0)
            If OledatableEnreg.Rows.Count <> 0 Then
                Error_Article.WriteLine("L'article Existe Déja " & Trim(Codeartdistri) & " Code ART Fournisseur " & Codeartfour)
            Else
                If Trim(Codeartfour) <> "" And Trim(Codefour) <> "" Then
                    Insertion = "Insert Into ARTICLE (Fournisseur,Code_Article_Fo,Code_Article_Dis) VALUES ('" & Trim(Codefour) & "','" & Trim(Codeartfour) & "','" & Trim(Codeartdistri) & "')"
                    OleCommandEnreg = New OleDbCommand(Insertion)
                    OleCommandEnreg.Connection = OleConnenection
                    OleCommandEnreg.ExecuteNonQuery()
                    Error_Article.WriteLine("Article Créé " & Trim(Codeartdistri) & " Code ART Fournisseur " & Codeartfour)
                End If
            End If
        End If
    End Sub

    Private Sub LectureXml_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TextFour.Text = ""
        Textxml.Text = ""
        Dim BaseBool As Boolean
        BaseBool = Connected()
    End Sub
End Class