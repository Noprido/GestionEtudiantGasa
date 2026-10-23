' ============================================================
' FrmMatieres.vb  -  Catalogue des matières (nom + coefficient)
' Permet aussi d'ajouter la matière sélectionnée à une classe.
' ============================================================
Imports System.Drawing
Imports System.Windows.Forms

Public Class FrmMatieres
    Inherits Form

    Private ReadOnly grille As DataGridView = Ui.Grille()
    Private ReadOnly txtNom As New TextBox()
    Private ReadOnly numCoef As New NumericUpDown()
    Private ReadOnly cboClasse As New ComboBox()

    Public Sub New()
        Text = "Catalogue des matières"
        StartPosition = FormStartPosition.CenterParent
        ClientSize = New Size(760, 540)
        MinimumSize = New Size(700, 480)
        Font = New Font("Segoe UI", 10)

        grille.Columns.Add("Nom", "Matière")
        grille.Columns.Add("Coefficient", "Coefficient")

        Dim bas As New Panel With {.Dock = DockStyle.Bottom, .Height = 190}

        Dim lblNom = Ui.Etiquette("Nom de la matière :")
        lblNom.Location = New Point(12, 12)
        txtNom.Location = New Point(12, 36)
        txtNom.Width = 320

        Dim lblCoef = Ui.Etiquette("Coefficient :")
        lblCoef.Location = New Point(350, 12)
        numCoef.Location = New Point(350, 36)
        numCoef.Width = 100
        numCoef.DecimalPlaces = 2
        numCoef.Increment = 0.5D
        numCoef.Minimum = 0.01D
        numCoef.Maximum = 100D
        numCoef.Value = 1D

        Dim btnAjouter = Ui.Bouton("Ajouter", 110)
        btnAjouter.Location = New Point(12, 78)
        Dim btnModifier = Ui.Bouton("Modifier", 110)
        btnModifier.Location = New Point(132, 78)
        Dim btnSupprimer = Ui.Bouton("Supprimer", 110)
        btnSupprimer.Location = New Point(252, 78)

        Dim lblClasse = Ui.Etiquette("Ajouter la matière sélectionnée à une classe :")
        lblClasse.Location = New Point(12, 128)
        cboClasse.Location = New Point(12, 152)
        cboClasse.Width = 320
        cboClasse.DropDownStyle = ComboBoxStyle.DropDownList
        Dim btnVersClasse = Ui.Bouton("Ajouter à la classe", 170)
        btnVersClasse.Location = New Point(345, 148)

        bas.Controls.AddRange(New Control() {lblNom, txtNom, lblCoef, numCoef,
                                             btnAjouter, btnModifier, btnSupprimer,
                                             lblClasse, cboClasse, btnVersClasse})
        Ui.Empiler(Me, grille, bas)

        AddHandler grille.SelectionChanged, AddressOf SurSelection
        AddHandler btnAjouter.Click, AddressOf SurAjouter
        AddHandler btnModifier.Click, AddressOf SurModifier
        AddHandler btnSupprimer.Click, AddressOf SurSupprimer
        AddHandler btnVersClasse.Click, AddressOf SurAjouterAClasse
        AddHandler Me.Load, AddressOf SurChargement
    End Sub

    Private Sub SurChargement(sender As Object, e As EventArgs)
        cboClasse.Items.Clear()
        For Each c In ClasseDepot.Lister()
            cboClasse.Items.Add(c)
        Next
        If cboClasse.Items.Count > 0 Then cboClasse.SelectedIndex = 0
        Charger()
    End Sub

    Private Function MatiereSelectionnee() As Matiere
        If grille.SelectedRows.Count = 0 Then Return Nothing
        Return TryCast(grille.SelectedRows(0).Tag, Matiere)
    End Function

    Private Sub Charger(Optional idASelectionner As Integer = 0)
        grille.Rows.Clear()
        Dim aSelectionner As DataGridViewRow = Nothing
        For Each m In MatiereDepot.Lister()
            Dim i = grille.Rows.Add(m.Nom, Outils.FmtNote(m.Coefficient))
            grille.Rows(i).Tag = m
            If m.Id = idASelectionner Then aSelectionner = grille.Rows(i)
        Next
        If aSelectionner IsNot Nothing Then
            grille.ClearSelection()
            aSelectionner.Selected = True
            grille.CurrentCell = aSelectionner.Cells(0)
        End If
    End Sub

    Private Sub SurSelection(sender As Object, e As EventArgs)
        Dim m = MatiereSelectionnee()
        If m Is Nothing Then Return
        txtNom.Text = m.Nom
        numCoef.Value = CDec(Math.Min(100, Math.Max(0.01, m.Coefficient)))
    End Sub

    Private Sub SurAjouter(sender As Object, e As EventArgs)
        Dim nom = txtNom.Text.Trim()
        If nom = "" Then
            Ui.Erreur("Saisis le nom de la matière.")
            Return
        End If
        If MatiereDepot.NomExiste(nom) Then
            Ui.Erreur("Cette matière existe déjà dans le catalogue.")
            Return
        End If
        Dim id = MatiereDepot.Ajouter(nom, CDbl(numCoef.Value))
        Charger(id)
    End Sub

    Private Sub SurModifier(sender As Object, e As EventArgs)
        Dim m = MatiereSelectionnee()
        If m Is Nothing Then
            Ui.Info("Sélectionne d'abord une matière dans la liste.")
            Return
        End If
        Dim nom = txtNom.Text.Trim()
        If nom = "" Then
            Ui.Erreur("Saisis le nom de la matière.")
            Return
        End If
        If MatiereDepot.NomExiste(nom, m.Id) Then
            Ui.Erreur("Une autre matière porte déjà ce nom.")
            Return
        End If
        MatiereDepot.Modifier(m.Id, nom, CDbl(numCoef.Value))
        Charger(m.Id)
    End Sub

    Private Sub SurSupprimer(sender As Object, e As EventArgs)
        Dim m = MatiereSelectionnee()
        If m Is Nothing Then
            Ui.Info("Sélectionne d'abord une matière dans la liste.")
            Return
        End If
        If MatiereDepot.EstUtilisee(m.Id) Then
            Ui.Erreur("Cette matière est utilisée dans au moins une classe." & vbCrLf &
                      "Retire-la d'abord de ces classes, puis réessaie.")
            Return
        End If
        If Not Ui.Confirmer($"Supprimer la matière « {m.Nom} » du catalogue ?") Then Return
        MatiereDepot.Supprimer(m.Id)
        txtNom.Clear()
        Charger()
    End Sub

    Private Sub SurAjouterAClasse(sender As Object, e As EventArgs)
        Dim m = MatiereSelectionnee()
        If m Is Nothing Then
            Ui.Info("Sélectionne d'abord une matière dans la liste.")
            Return
        End If
        Dim c = TryCast(cboClasse.SelectedItem, Classe)
        If c Is Nothing Then
            Ui.Info("Choisis une classe dans la liste déroulante (crée-en une d'abord si besoin).")
            Return
        End If
        If ClasseDepot.MatiereDejaPresente(c.Id, m.Id) Then
            Ui.Info($"« {m.Nom} » fait déjà partie de la classe {c.Nom}.")
            Return
        End If
        ClasseDepot.AjouterMatiere(c.Id, m.Id)
        Ui.Info($"« {m.Nom} » a été ajoutée à la classe {c.Nom}.")
    End Sub

End Class
