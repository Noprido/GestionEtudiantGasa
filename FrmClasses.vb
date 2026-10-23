' ============================================================
' FrmClasses.vb  -  Gestion des classes et des matières de chaque classe
' Gauche : liste des classes. Droite : matières de la classe sélectionnée
' (ajout via liste déroulante alimentée par le catalogue).
' ============================================================
Imports System.Drawing
Imports System.Windows.Forms

Public Class FrmClasses
    Inherits Form

    Private ReadOnly grilleClasses As DataGridView = Ui.Grille()
    Private ReadOnly grilleMatieres As DataGridView = Ui.Grille()
    Private ReadOnly txtNom As New TextBox()
    Private ReadOnly txtAnnee As New TextBox()
    Private ReadOnly cboMatiere As New ComboBox()
    Private _enChargement As Boolean

    Public Sub New()
        Text = "Classes et matières de classe"
        StartPosition = FormStartPosition.CenterParent
        ClientSize = New Size(1000, 560)
        MinimumSize = New Size(900, 500)
        Font = New Font("Segoe UI", 10)

        grilleClasses.Columns.Add("Nom", "Classe")
        grilleClasses.Columns.Add("Annee", "Année scolaire")
        grilleMatieres.Columns.Add("Nom", "Matière")
        grilleMatieres.Columns.Add("Coef", "Coefficient")

        ' ---------- Côté gauche : les classes ----------
        Dim panGauche As New Panel With {.Dock = DockStyle.Left, .Width = 470}
        Dim titreG As New Label With {
            .Text = "Classes", .Dock = DockStyle.Top, .Height = 30,
            .Font = New Font("Segoe UI", 11, FontStyle.Bold), .TextAlign = ContentAlignment.MiddleLeft}
        Dim basG As New Panel With {.Dock = DockStyle.Bottom, .Height = 130}

        Dim lblNom = Ui.Etiquette("Nom de la classe :")
        lblNom.Location = New Point(10, 10)
        txtNom.Location = New Point(10, 34)
        txtNom.Width = 230
        Dim lblAnnee = Ui.Etiquette("Année scolaire :")
        lblAnnee.Location = New Point(255, 10)
        txtAnnee.Location = New Point(255, 34)
        txtAnnee.Width = 150
        txtAnnee.Text = AnneeParDefaut()

        Dim btnAjouter = Ui.Bouton("Ajouter", 110)
        btnAjouter.Location = New Point(10, 80)
        Dim btnModifier = Ui.Bouton("Modifier", 110)
        btnModifier.Location = New Point(130, 80)
        Dim btnSupprimer = Ui.Bouton("Supprimer", 110)
        btnSupprimer.Location = New Point(250, 80)

        basG.Controls.AddRange(New Control() {lblNom, txtNom, lblAnnee, txtAnnee, btnAjouter, btnModifier, btnSupprimer})
        Ui.Empiler(panGauche, grilleClasses, titreG, basG)

        ' ---------- Côté droit : matières de la classe ----------
        Dim panDroit As New Panel With {.Dock = DockStyle.Fill}
        Dim titreD As New Label With {
            .Text = "   Matières de la classe sélectionnée", .Dock = DockStyle.Top, .Height = 30,
            .Font = New Font("Segoe UI", 11, FontStyle.Bold), .TextAlign = ContentAlignment.MiddleLeft}
        Dim basD As New Panel With {.Dock = DockStyle.Bottom, .Height = 170}

        Dim lblMat = Ui.Etiquette("Matière du catalogue à ajouter :")
        lblMat.Location = New Point(12, 10)
        cboMatiere.Location = New Point(12, 34)
        cboMatiere.Width = 290
        cboMatiere.DropDownStyle = ComboBoxStyle.DropDownList
        Dim btnAjouterMat = Ui.Bouton("Ajouter à la classe", 170)
        btnAjouterMat.Location = New Point(312, 30)
        Dim btnRetirerMat = Ui.Bouton("Retirer la matière sélectionnée", 250)
        btnRetirerMat.Location = New Point(12, 76)
        Dim btnCatalogue = Ui.Bouton("Gérer le catalogue", 170)
        btnCatalogue.Location = New Point(272, 76)
        Dim btnEtudiants = Ui.Bouton("Étudiants de la classe", 250)
        btnEtudiants.Location = New Point(12, 122)

        basD.Controls.AddRange(New Control() {lblMat, cboMatiere, btnAjouterMat, btnRetirerMat, btnCatalogue, btnEtudiants})
        Ui.Empiler(panDroit, grilleMatieres, titreD, basD)

        Ui.Empiler(Me, panDroit, panGauche)

        AddHandler grilleClasses.SelectionChanged, AddressOf SurSelectionClasse
        AddHandler btnAjouter.Click, AddressOf SurAjouterClasse
        AddHandler btnModifier.Click, AddressOf SurModifierClasse
        AddHandler btnSupprimer.Click, AddressOf SurSupprimerClasse
        AddHandler btnAjouterMat.Click, AddressOf SurAjouterMatiere
        AddHandler btnRetirerMat.Click, AddressOf SurRetirerMatiere
        AddHandler btnCatalogue.Click, AddressOf SurCatalogue
        AddHandler btnEtudiants.Click, AddressOf SurEtudiants
        AddHandler Me.Load, AddressOf SurChargement
    End Sub

    Private Shared Function AnneeParDefaut() As String
        Dim d = Date.Today
        Dim debut = If(d.Month >= 9, d.Year, d.Year - 1)
        Return debut & "-" & (debut + 1)
    End Function

    Private Sub SurChargement(sender As Object, e As EventArgs)
        ChargerClasses()
    End Sub

    Private Function ClasseSelectionnee() As Classe
        If grilleClasses.SelectedRows.Count = 0 Then Return Nothing
        Return TryCast(grilleClasses.SelectedRows(0).Tag, Classe)
    End Function

    Private Sub ChargerClasses(Optional idASelectionner As Integer = 0)
        _enChargement = True
        grilleClasses.Rows.Clear()
        Dim aSelectionner As DataGridViewRow = Nothing
        For Each c In ClasseDepot.Lister()
            Dim i = grilleClasses.Rows.Add(c.Nom, c.AnneeScolaire)
            grilleClasses.Rows(i).Tag = c
            If c.Id = idASelectionner Then aSelectionner = grilleClasses.Rows(i)
        Next
        If aSelectionner Is Nothing AndAlso grilleClasses.Rows.Count > 0 Then aSelectionner = grilleClasses.Rows(0)
        If aSelectionner IsNot Nothing Then
            grilleClasses.ClearSelection()
            aSelectionner.Selected = True
            grilleClasses.CurrentCell = aSelectionner.Cells(0)
        End If
        _enChargement = False
        AfficherSelectionClasse()
    End Sub

    Private Sub AfficherSelectionClasse()
        Dim c = ClasseSelectionnee()
        If c IsNot Nothing Then
            txtNom.Text = c.Nom
            txtAnnee.Text = c.AnneeScolaire
        End If
        ChargerMatieresClasse()
    End Sub

    Private Sub ChargerMatieresClasse()
        grilleMatieres.Rows.Clear()
        cboMatiere.Items.Clear()
        Dim c = ClasseSelectionnee()
        If c Is Nothing Then Return
        For Each m In MatiereDepot.ListerPourClasse(c.Id)
            Dim i = grilleMatieres.Rows.Add(m.Nom, Outils.FmtNote(m.Coefficient))
            grilleMatieres.Rows(i).Tag = m
        Next
        For Each m In MatiereDepot.ListerHorsClasse(c.Id)
            cboMatiere.Items.Add(m)
        Next
        If cboMatiere.Items.Count > 0 Then cboMatiere.SelectedIndex = 0
    End Sub

    Private Sub SurSelectionClasse(sender As Object, e As EventArgs)
        If _enChargement Then Return
        AfficherSelectionClasse()
    End Sub

    ' ---------- Classes ----------

    Private Sub SurAjouterClasse(sender As Object, e As EventArgs)
        Dim nom = txtNom.Text.Trim()
        If nom = "" Then
            Ui.Erreur("Saisis le nom de la classe.")
            Return
        End If
        If ClasseDepot.NomExiste(nom) Then
            Ui.Erreur("Une classe porte déjà ce nom.")
            Return
        End If
        Dim id = ClasseDepot.Ajouter(nom, txtAnnee.Text.Trim())
        ChargerClasses(id)
    End Sub

    Private Sub SurModifierClasse(sender As Object, e As EventArgs)
        Dim c = ClasseSelectionnee()
        If c Is Nothing Then
            Ui.Info("Sélectionne d'abord une classe dans la liste.")
            Return
        End If
        Dim nom = txtNom.Text.Trim()
        If nom = "" Then
            Ui.Erreur("Saisis le nom de la classe.")
            Return
        End If
        If ClasseDepot.NomExiste(nom, c.Id) Then
            Ui.Erreur("Une autre classe porte déjà ce nom.")
            Return
        End If
        ClasseDepot.Modifier(c.Id, nom, txtAnnee.Text.Trim())
        ChargerClasses(c.Id)
    End Sub

    Private Sub SurSupprimerClasse(sender As Object, e As EventArgs)
        Dim c = ClasseSelectionnee()
        If c Is Nothing Then
            Ui.Info("Sélectionne d'abord une classe dans la liste.")
            Return
        End If
        If Not Ui.Confirmer($"Supprimer la classe « {c.Nom} » ?" & vbCrLf &
                            "Ses étudiants et toutes leurs notes seront aussi supprimés.") Then Return
        ClasseDepot.Supprimer(c.Id)
        txtNom.Clear()
        ChargerClasses()
    End Sub

    ' ---------- Matières de la classe ----------

    Private Sub SurAjouterMatiere(sender As Object, e As EventArgs)
        Dim c = ClasseSelectionnee()
        If c Is Nothing Then
            Ui.Info("Sélectionne d'abord une classe.")
            Return
        End If
        Dim m = TryCast(cboMatiere.SelectedItem, Matiere)
        If m Is Nothing Then
            Ui.Info("Aucune matière à ajouter. Crée-en dans le catalogue (bouton « Gérer le catalogue »).")
            Return
        End If
        ClasseDepot.AjouterMatiere(c.Id, m.Id)
        ChargerMatieresClasse()
    End Sub

    Private Sub SurRetirerMatiere(sender As Object, e As EventArgs)
        Dim c = ClasseSelectionnee()
        If c Is Nothing OrElse grilleMatieres.SelectedRows.Count = 0 Then
            Ui.Info("Sélectionne une classe puis une de ses matières.")
            Return
        End If
        Dim m = TryCast(grilleMatieres.SelectedRows(0).Tag, Matiere)
        If m Is Nothing Then Return
        If Not Ui.Confirmer($"Retirer « {m.Nom} » de la classe {c.Nom} ?" & vbCrLf &
                            "Les notes saisies dans cette matière pour cette classe seront supprimées.") Then Return
        ClasseDepot.RetirerMatiere(c.Id, m.Id)
        ChargerMatieresClasse()
    End Sub

    Private Sub SurCatalogue(sender As Object, e As EventArgs)
        Using f As New FrmMatieres()
            f.ShowDialog(Me)
        End Using
        ChargerMatieresClasse()
    End Sub

    Private Sub SurEtudiants(sender As Object, e As EventArgs)
        Dim c = ClasseSelectionnee()
        If c Is Nothing Then
            Ui.Info("Sélectionne d'abord une classe.")
            Return
        End If
        Using f As New FrmEtudiants(c.Id)
            f.ShowDialog(Me)
        End Using
    End Sub

End Class
