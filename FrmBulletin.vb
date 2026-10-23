' ============================================================
' FrmBulletin.vb  -  Bulletin d'un étudiant affiché dans un formulaire, avec export PDF
' ============================================================
Imports System.Drawing
Imports System.Windows.Forms

Public Class FrmBulletin
    Inherits Form

    Private ReadOnly _rc As ResultatClasse
    Private ReadOnly _r As ResultatEtudiant

    Public Sub New(rc As ResultatClasse, r As ResultatEtudiant)
        _rc = rc
        _r = r

        Text = "Bulletin - " & r.Etudiant.NomComplet
        StartPosition = FormStartPosition.CenterParent
        ClientSize = New Size(800, 680)
        MinimumSize = New Size(760, 600)
        Font = New Font("Segoe UI", 10)

        ' ---------- En-tête ----------
        Dim entete As New TableLayoutPanel With {.Dock = DockStyle.Top, .AutoSize = True, .ColumnCount = 1}
        entete.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
        AjouterLigne(entete, "UATM GASA FORMATION", 18, True, 42)
        AjouterLigne(entete, "BULLETIN DE NOTES", 13, True, 30)
        If rc.Classe.AnneeScolaire <> "" Then
            AjouterLigne(entete, "Année scolaire : " & rc.Classe.AnneeScolaire, 10, False, 24)
        End If
        AjouterLigne(entete, "Classe : " & rc.Classe.Nom, 10, False, 24)
        AjouterLigne(entete, r.Etudiant.NomComplet, 13, True, 38)

        ' ---------- Tableau des matières ----------
        Dim grille = Ui.Grille()
        grille.Columns.Add("Matiere", "Matière")
        grille.Columns.Add("Coef", "Coef")
        grille.Columns.Add("I1", "Interro 1")
        grille.Columns.Add("I2", "Interro 2")
        grille.Columns.Add("Dev", "Devoir")
        grille.Columns.Add("Moy", "Moyenne")
        grille.Columns.Add("Statut", "Statut")
        grille.Columns(0).FillWeight = 200
        For i = 1 To 6
            grille.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Next
        For Each l In r.Lignes
            Dim idx = grille.Rows.Add(l.Nom, Outils.FmtNote(l.Coefficient), Outils.FmtNote(l.Interro1),
                                      Outils.FmtNote(l.Interro2), Outils.FmtNote(l.Devoir),
                                      Outils.Fmt(l.Moyenne), l.Statut)
            Dim couleur As Color
            If Not l.Moyenne.HasValue Then
                couleur = Color.Gray
            ElseIf Calculateur.EstValidee(l.Moyenne) Then
                couleur = Color.DarkGreen
            Else
                couleur = Color.Firebrick
            End If
            grille.Rows(idx).Cells(5).Style.ForeColor = couleur
            grille.Rows(idx).Cells(6).Style.ForeColor = couleur
        Next
        grille.ClearSelection()

        ' ---------- Récapitulatif ----------
        Dim bas As New Panel With {.Dock = DockStyle.Bottom, .Height = 200}
        Dim police = New Font("Segoe UI", 12, FontStyle.Bold)

        Dim lblMoy = Ui.Etiquette("Moyenne générale : " & Outils.Fmt(r.MoyenneGenerale) & " / 20")
        lblMoy.Font = police
        lblMoy.Location = New Point(20, 12)
        Dim lblMention = Ui.Etiquette("Mention : " & r.Mention)
        lblMention.Font = police
        lblMention.Location = New Point(20, 44)
        Dim lblRang = Ui.Etiquette("Rang : " & r.RangTexte & " sur " & r.Effectif)
        lblRang.Font = police
        lblRang.Location = New Point(20, 76)
        Dim lblStats = Ui.Etiquette("Moyenne de la classe : " & Outils.Fmt(rc.MoyenneClasse) &
                                    "     Plus forte : " & Outils.Fmt(rc.MoyenneMax) &
                                    "     Plus faible : " & Outils.Fmt(rc.MoyenneMin))
        lblStats.Location = New Point(20, 112)

        Dim btnPdf = Ui.Bouton("Exporter en PDF", 170)
        btnPdf.Location = New Point(20, 152)
        Dim btnFermer = Ui.Bouton("Fermer", 110)
        btnFermer.Location = New Point(200, 152)

        bas.Controls.AddRange(New Control() {lblMoy, lblMention, lblRang, lblStats, btnPdf, btnFermer})

        Ui.Empiler(Me, grille, entete, bas)

        AddHandler btnPdf.Click, AddressOf SurExporter
        AddHandler btnFermer.Click, Sub(s, e) Close()
    End Sub

    Private Shared Sub AjouterLigne(tlp As TableLayoutPanel, texte As String, taille As Single, gras As Boolean, hauteur As Integer)
        Dim l As New Label With {
            .Text = texte, .AutoSize = False, .Dock = DockStyle.Fill, .Height = hauteur,
            .TextAlign = ContentAlignment.MiddleCenter,
            .Font = New Font("Segoe UI", taille, If(gras, FontStyle.Bold, FontStyle.Regular))}
        tlp.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        tlp.Controls.Add(l)
    End Sub

    Private Sub SurExporter(sender As Object, e As EventArgs)
        BulletinPdf.ExporterAvecDialogue(Me, _rc, New ResultatEtudiant() {_r}, "Bulletin_" & _r.Etudiant.NomComplet)
    End Sub

End Class
