' ============================================================
' FrmClassement.vb  -  Résultat final par ordre de mérite (moyenne générale décroissante,
' ex aequo gérés), accès au bulletin et export PDF.
' ============================================================
Imports System.Drawing
Imports System.Windows.Forms

Public Class FrmClassement
    Inherits Form

    Private ReadOnly cboClasse As New ComboBox()
    Private ReadOnly grille As DataGridView = Ui.Grille()
    Private ReadOnly lblStats As New Label()
    Private _rc As ResultatClasse
    Private _enChargement As Boolean

    Public Sub New()
        Text = "Classement et bulletins"
        StartPosition = FormStartPosition.CenterParent
        ClientSize = New Size(1000, 580)
        MinimumSize = New Size(860, 460)
        Font = New Font("Segoe UI", 10)

        Dim haut As New Panel With {.Dock = DockStyle.Top, .Height = 100}
        Dim lblClasse = Ui.Etiquette("Classe :")
        lblClasse.Location = New Point(12, 18)
        cboClasse.Location = New Point(75, 14)
        cboClasse.Width = 300
        cboClasse.DropDownStyle = ComboBoxStyle.DropDownList
        Dim btnActualiser = Ui.Bouton("Actualiser", 120)
        btnActualiser.Location = New Point(390, 10)
        Dim btnVoir = Ui.Bouton("Voir le bulletin", 150)
        btnVoir.Location = New Point(12, 56)
        Dim btnPdf = Ui.Bouton("Exporter ce bulletin (PDF)", 230)
        btnPdf.Location = New Point(172, 56)
        Dim btnPdfTous = Ui.Bouton("Exporter toute la classe (PDF)", 250)
        btnPdfTous.Location = New Point(412, 56)
        haut.Controls.AddRange(New Control() {lblClasse, cboClasse, btnActualiser, btnVoir, btnPdf, btnPdfTous})

        Dim bas As New Panel With {.Dock = DockStyle.Bottom, .Height = 40}
        lblStats.AutoSize = True
        lblStats.Location = New Point(12, 10)
        bas.Controls.Add(lblStats)

        Ui.Empiler(Me, grille, haut, bas)

        AddHandler cboClasse.SelectedIndexChanged, AddressOf SurClasseChange
        AddHandler btnActualiser.Click, AddressOf SurActualiser
        AddHandler btnVoir.Click, AddressOf SurVoirBulletin
        AddHandler grille.CellDoubleClick, AddressOf SurDoubleClic
        AddHandler btnPdf.Click, AddressOf SurExporterUn
        AddHandler btnPdfTous.Click, AddressOf SurExporterTous
        AddHandler Me.Load, AddressOf SurChargement
    End Sub

    Private Sub SurChargement(sender As Object, e As EventArgs)
        _enChargement = True
        cboClasse.Items.Clear()
        For Each c In ClasseDepot.Lister()
            cboClasse.Items.Add(c)
        Next
        If cboClasse.Items.Count > 0 Then cboClasse.SelectedIndex = 0
        _enChargement = False
        Actualiser()
    End Sub

    Private Function ClasseCourante() As Classe
        Return TryCast(cboClasse.SelectedItem, Classe)
    End Function

    Private Function ResultatSelectionne() As ResultatEtudiant
        If grille.SelectedRows.Count = 0 Then Return Nothing
        Return TryCast(grille.SelectedRows(0).Tag, ResultatEtudiant)
    End Function

    Private Sub SurClasseChange(sender As Object, e As EventArgs)
        If _enChargement Then Return
        Actualiser()
    End Sub

    Private Sub SurActualiser(sender As Object, e As EventArgs)
        Actualiser()
    End Sub

    ''' <summary>Recalcule le classement de la classe choisie et reconstruit la grille.</summary>
    Private Sub Actualiser()
        grille.Columns.Clear()
        grille.Rows.Clear()
        lblStats.Text = ""
        _rc = Nothing

        Dim c = ClasseCourante()
        If c Is Nothing Then Return
        _rc = Classement.Calculer(c.Id)

        grille.Columns.Add("Rang", "Rang")
        grille.Columns.Add("Nom", "Nom et prénoms")
        For Each m In _rc.Matieres
            grille.Columns.Add("M" & m.Id, m.Nom & " (coef " & Outils.FmtNote(m.Coefficient) & ")")
        Next
        grille.Columns.Add("MoyGen", "Moyenne générale")
        grille.Columns.Add("Mention", "Mention")
        grille.Columns(0).FillWeight = 55
        grille.Columns(1).FillWeight = 190
        For i = 2 To grille.Columns.Count - 1
            grille.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Next

        For Each r In _rc.Resultats
            Dim valeurs As New List(Of Object) From {r.RangTexte, r.Etudiant.NomComplet}
            For Each l In r.Lignes
                valeurs.Add(Outils.Fmt(l.Moyenne))
            Next
            valeurs.Add(Outils.Fmt(r.MoyenneGenerale))
            valeurs.Add(r.Mention)

            Dim idx = grille.Rows.Add(valeurs.ToArray())
            grille.Rows(idx).Tag = r

            For j = 0 To r.Lignes.Count - 1
                Dim moy = r.Lignes(j).Moyenne
                If moy.HasValue AndAlso Not Calculateur.EstValidee(moy) Then
                    grille.Rows(idx).Cells(2 + j).Style.ForeColor = Color.Firebrick
                End If
            Next
            Dim colMoyGen = 2 + r.Lignes.Count
            grille.Rows(idx).Cells(colMoyGen).Style.Font = New Font(grille.Font, FontStyle.Bold)
            If r.MoyenneGenerale.HasValue AndAlso r.MoyenneGenerale.Value < Calculateur.SeuilValidation Then
                grille.Rows(idx).Cells(colMoyGen).Style.ForeColor = Color.Firebrick
            End If
        Next

        lblStats.Text = "Effectif : " & _rc.Resultats.Count &
                        "     Moyenne de la classe : " & Outils.Fmt(_rc.MoyenneClasse) &
                        "     Plus forte : " & Outils.Fmt(_rc.MoyenneMax) &
                        "     Plus faible : " & Outils.Fmt(_rc.MoyenneMin)
        If _rc.Matieres.Count = 0 Then
            lblStats.Text = "Cette classe n'a encore aucune matière (menu « Classes et matières de classe »)."
        End If
    End Sub

    ' ---------- Bulletin ----------

    Private Sub SurVoirBulletin(sender As Object, e As EventArgs)
        OuvrirBulletin()
    End Sub

    Private Sub SurDoubleClic(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex >= 0 Then OuvrirBulletin()
    End Sub

    Private Sub OuvrirBulletin()
        Dim r = ResultatSelectionne()
        If r Is Nothing OrElse _rc Is Nothing Then
            Ui.Info("Sélectionne d'abord un étudiant dans le classement.")
            Return
        End If
        Using f As New FrmBulletin(_rc, r)
            f.ShowDialog(Me)
        End Using
    End Sub

    ' ---------- Export PDF ----------

    Private Sub SurExporterUn(sender As Object, e As EventArgs)
        Dim r = ResultatSelectionne()
        If r Is Nothing OrElse _rc Is Nothing Then
            Ui.Info("Sélectionne d'abord un étudiant dans le classement.")
            Return
        End If
        BulletinPdf.ExporterAvecDialogue(Me, _rc, New ResultatEtudiant() {r}, "Bulletin_" & r.Etudiant.NomComplet)
    End Sub

    Private Sub SurExporterTous(sender As Object, e As EventArgs)
        If _rc Is Nothing OrElse _rc.Resultats.Count = 0 Then
            Ui.Info("Aucun étudiant à exporter pour cette classe.")
            Return
        End If
        BulletinPdf.ExporterAvecDialogue(Me, _rc, _rc.Resultats, "Bulletins_" & _rc.Classe.Nom)
    End Sub

End Class
