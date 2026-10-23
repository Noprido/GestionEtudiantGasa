' ============================================================
' BulletinPdf.vb  -  Export des bulletins en PDF (une page par étudiant)
' Package NuGet requis : PDFsharp-GDI (version 6.x)
' ============================================================
Imports System.IO
Imports System.Windows.Forms
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf

Public Module BulletinPdf

    Private Const LargeurPage As Double = 595.0
    Private Const Marge As Double = 40.0

    ''' <summary>Demande où enregistrer, crée le PDF, propose de l'ouvrir.</summary>
    Public Sub ExporterAvecDialogue(proprietaire As IWin32Window, rc As ResultatClasse,
                                    resultats As IEnumerable(Of ResultatEtudiant), nomFichier As String)
        Dim liste = resultats.ToList()
        If liste.Count = 0 Then
            Ui.Info("Aucun étudiant à exporter.")
            Return
        End If

        Dim chemin As String
        Using dlg As New SaveFileDialog With {
            .Filter = "Fichier PDF (*.pdf)|*.pdf",
            .FileName = NettoyerNom(nomFichier) & ".pdf",
            .Title = "Enregistrer en PDF"}
            If dlg.ShowDialog(proprietaire) <> DialogResult.OK Then Return
            chemin = dlg.FileName
        End Using

        Try
            Exporter(rc, liste, chemin)
        Catch ex As Exception
            Ui.Erreur("Impossible de créer le PDF :" & vbCrLf & ex.Message)
            Return
        End Try

        If Ui.Confirmer("Le PDF a été créé. Voulez-vous l'ouvrir ?") Then Ui.OuvrirFichier(chemin)
    End Sub

    Public Sub Exporter(rc As ResultatClasse, resultats As IEnumerable(Of ResultatEtudiant), chemin As String)
        PdfSharp.PdfSharpLicense.SetAgplLicense()
        Dim doc As New PdfDocument()
        doc.Info.Title = "Bulletins - " & rc.Classe.Nom
        For Each r In resultats
            Dim page = doc.AddPage()
            Using gfx = XGraphics.FromPdfPage(page)
                DessinerBulletin(gfx, rc, r)
            End Using
        Next
        doc.Save(chemin)
    End Sub

    Public Function NettoyerNom(nom As String) As String
        For Each c In Path.GetInvalidFileNameChars()
            nom = nom.Replace(c, "_"c)
        Next
        Return nom
    End Function

    ' ---------------------------------------------------------------

    Private Sub DessinerBulletin(gfx As XGraphics, rc As ResultatClasse, r As ResultatEtudiant)
        Dim fTitre As New XFont("Arial", 18, XFontStyleEx.Bold)
        Dim fSousTitre As New XFont("Arial", 13, XFontStyleEx.Bold)
        Dim fNormal As New XFont("Arial", 10, XFontStyleEx.Regular)
        Dim fGras As New XFont("Arial", 10, XFontStyleEx.Bold)
        Dim fGrand As New XFont("Arial", 12, XFontStyleEx.Bold)
        Dim fPetit As New XFont("Arial", 8, XFontStyleEx.Regular)

        Dim largeur As Double = LargeurPage - 2 * Marge
        Dim y As Double = Marge

        ' --- En-tête ---
        Texte(gfx, "UATM GASA FORMATION", fTitre, Marge, y, largeur, 26, XStringFormats.Center)
        y += 30
        Texte(gfx, "BULLETIN DE NOTES", fSousTitre, Marge, y, largeur, 20, XStringFormats.Center)
        y += 24
        If rc.Classe.AnneeScolaire <> "" Then
            Texte(gfx, "Année scolaire : " & rc.Classe.AnneeScolaire, fNormal, Marge, y, largeur, 16, XStringFormats.Center)
            y += 18
        End If
        gfx.DrawLine(XPens.Black, Marge, y + 4, Marge + largeur, y + 4)
        y += 16

        ' --- Identité ---
        Texte(gfx, "Nom et prénoms : " & r.Etudiant.NomComplet, fGrand, Marge, y, largeur, 18, XStringFormats.TopLeft)
        y += 22
        Texte(gfx, "Classe : " & rc.Classe.Nom & "      Effectif : " & r.Effectif, fNormal, Marge, y, largeur, 16, XStringFormats.TopLeft)
        y += 30

        ' --- Tableau des matières ---
        Dim largeurs As Double() = {150, 35, 50, 50, 50, 60, 120}
        Const hauteurLigne As Double = 20

        DessinerLigne(gfx, Marge, y, hauteurLigne, largeurs,
                      New String() {"Matière", "Coef", "Interro 1", "Interro 2", "Devoir", "Moyenne", "Statut"},
                      fGras, XBrushes.LightGray)
        y += hauteurLigne

        For Each l In r.Lignes
            DessinerLigne(gfx, Marge, y, hauteurLigne, largeurs,
                          New String() {l.Nom, Outils.FmtNote(l.Coefficient), Outils.FmtNote(l.Interro1),
                                        Outils.FmtNote(l.Interro2), Outils.FmtNote(l.Devoir),
                                        Outils.Fmt(l.Moyenne), l.Statut},
                          fNormal, Nothing)
            y += hauteurLigne
        Next
        y += 24

        ' --- Récapitulatif ---
        Texte(gfx, "Moyenne générale : " & Outils.Fmt(r.MoyenneGenerale) & " / 20", fGrand, Marge, y, largeur, 18, XStringFormats.TopLeft)
        y += 24
        Texte(gfx, "Mention : " & r.Mention, fGrand, Marge, y, largeur, 18, XStringFormats.TopLeft)
        y += 24
        Texte(gfx, "Rang : " & r.RangTexte & " sur " & r.Effectif, fGrand, Marge, y, largeur, 18, XStringFormats.TopLeft)
        y += 32

        Texte(gfx, "Moyenne de la classe : " & Outils.Fmt(rc.MoyenneClasse) &
                   "     Plus forte moyenne : " & Outils.Fmt(rc.MoyenneMax) &
                   "     Plus faible moyenne : " & Outils.Fmt(rc.MoyenneMin),
              fNormal, Marge, y, largeur, 16, XStringFormats.TopLeft)
        y += 50

        ' --- Pied de page ---
        Texte(gfx, "Édité le " & Date.Today.ToString("dd/MM/yyyy"), fNormal, Marge, y, 200, 16, XStringFormats.TopLeft)
        Texte(gfx, "La Direction", fGras, Marge + largeur - 160, y, 160, 16, XStringFormats.TopCenter)
        Texte(gfx, "Une matière est validée à partir d'une moyenne de 10/20.", fPetit, Marge, 800, largeur, 12, XStringFormats.TopLeft)
    End Sub

    Private Sub Texte(gfx As XGraphics, s As String, police As XFont,
                      x As Double, y As Double, largeur As Double, hauteur As Double, format As XStringFormat)
        gfx.DrawString(s, police, XBrushes.Black, New XRect(x, y, largeur, hauteur), format)
    End Sub

    Private Sub DessinerLigne(gfx As XGraphics, x0 As Double, y As Double, hauteur As Double,
                              largeurs As Double(), valeurs As String(), police As XFont, fond As XBrush)
        Dim x As Double = x0
        For i = 0 To largeurs.Length - 1
            Dim cadre As New XRect(x, y, largeurs(i), hauteur)
            If fond IsNot Nothing Then
                gfx.DrawRectangle(XPens.Black, fond, cadre)
            Else
                gfx.DrawRectangle(XPens.Black, cadre)
            End If
            Dim decalage As Double = If(i = 0, 4, 0)
            Dim zone As New XRect(x + decalage, y, largeurs(i) - decalage, hauteur)
            Dim alignement = If(i = 0, XStringFormats.CenterLeft, XStringFormats.Center)
            gfx.DrawString(valeurs(i), police, XBrushes.Black, zone, alignement)
            x += largeurs(i)
        Next
    End Sub

End Module
