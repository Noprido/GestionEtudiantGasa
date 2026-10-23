' ============================================================
' FrmEtudiants.vb  -  Gestion des étudiants d'une classe
' ============================================================
Imports System.Drawing
Imports System.Windows.Forms

Public Class FrmEtudiants
    Inherits Form

    Private ReadOnly cboClasse As New ComboBox()
    Private ReadOnly grille As DataGridView = Ui.Grille()
    Private ReadOnly txtNom As New TextBox()
    Private ReadOnly txtPrenoms As New TextBox()
    Private ReadOnly lblNombre As New Label()
    Private ReadOnly _classeInitiale As Integer
    Private _enChargement As Boolean

    ''' <param name="classeId">Classe à présélectionner (0 = la première).</param>
    Public Sub New(Optional classeId As Integer = 0)
        _classeInitiale = classeId
        Text = "Étudiants"
        StartPosition = FormStartPosition.CenterParent
        ClientSize = New Size(760, 540)
        MinimumSize = New Size(700, 460)
        Font = New Font("Segoe UI", 10)

        grille.Columns.Add("Nom", "Nom")
        grille.Columns.Add("Prenoms", "Prénoms")

        Dim haut As New Panel With {.Dock = DockStyle.Top, .Height = 52}
        Dim lblClasse = Ui.Etiquette("Classe :")
        lblClasse.Location = New Point(12, 16)
        cboClasse.Location = New Point(75, 12)
        cboClasse.Width = 300
        cboClasse.DropDownStyle = ComboBoxStyle.DropDownList
        lblNombre.AutoSize = True
        lblNombre.Location = New Point(395, 16)
        haut.Controls.AddRange(New Control() {lblClasse, cboClasse, lblNombre})

        Dim bas As New Panel With {.Dock = DockStyle.Bottom, .Height = 130}
        Dim lblNom = Ui.Etiquette("Nom :")
        lblNom.Location = New Point(12, 10)
        txtNom.Location = New Point(12, 34)
        txtNom.Width = 250
        Dim lblPrenoms = Ui.Etiquette("Prénoms :")
        lblPrenoms.Location = New Point(280, 10)
        txtPrenoms.Location = New Point(280, 34)
        txtPrenoms.Width = 250
        Dim btnAjouter = Ui.Bouton("Ajouter", 110)
        btnAjouter.Location = New Point(12, 80)
        Dim btnModifier = Ui.Bouton("Modifier", 110)
        btnModifier.Location = New Point(132, 80)
        Dim btnSupprimer = Ui.Bouton("Supprimer", 110)
        btnSupprimer.Location = New Point(252, 80)
        bas.Controls.AddRange(New Control() {lblNom, txtNom, lblPrenoms, txtPrenoms, btnAjouter, btnModifier, btnSupprimer})

        Ui.Empiler(Me, grille, haut, bas)

        AddHandler cboClasse.SelectedIndexChanged, AddressOf SurClasseChange
        AddHandler grille.SelectionChanged, AddressOf SurSelection
        AddHandler btnAjouter.Click, AddressOf SurAjouter
        AddHandler btnModifier.Click, AddressOf SurModifier
        AddHandler btnSupprimer.Click, AddressOf SurSupprimer
        AddHandler Me.Load, AddressOf SurChargement
    End Sub

    Private Sub SurChargement(sender As Object, e As EventArgs)
        _enChargement = True
        cboClasse.Items.Clear()
        For Each c In ClasseDepot.Lister()
            Dim i = cboClasse.Items.Add(c)
            If c.Id = _classeInitiale Then cboClasse.SelectedIndex = i
        Next
        If cboClasse.SelectedIndex < 0 AndAlso cboClasse.Items.Count > 0 Then cboClasse.SelectedIndex = 0
        _enChargement = False
        ChargerEtudiants()
    End Sub

    Private Function ClasseCourante() As Classe
        Return TryCast(cboClasse.SelectedItem, Classe)
    End Function

    Private Function EtudiantSelectionne() As Etudiant
        If grille.SelectedRows.Count = 0 Then Return Nothing
        Return TryCast(grille.SelectedRows(0).Tag, Etudiant)
    End Function

    Private Sub ChargerEtudiants(Optional idASelectionner As Integer = 0)
        _enChargement = True
        grille.Rows.Clear()
        Dim c = ClasseCourante()
        If c IsNot Nothing Then
            Dim aSelectionner As DataGridViewRow = Nothing
            Dim liste = EtudiantDepot.ListerParClasse(c.Id)
            For Each et In liste
                Dim i = grille.Rows.Add(et.Nom, et.Prenoms)
                grille.Rows(i).Tag = et
                If et.Id = idASelectionner Then aSelectionner = grille.Rows(i)
            Next
            lblNombre.Text = liste.Count & " étudiant(s)"
            If aSelectionner IsNot Nothing Then
                grille.ClearSelection()
                aSelectionner.Selected = True
                grille.CurrentCell = aSelectionner.Cells(0)
            End If
        Else
            lblNombre.Text = ""
        End If
        _enChargement = False
    End Sub

    Private Sub SurClasseChange(sender As Object, e As EventArgs)
        If _enChargement Then Return
        txtNom.Clear()
        txtPrenoms.Clear()
        ChargerEtudiants()
    End Sub

    Private Sub SurSelection(sender As Object, e As EventArgs)
        If _enChargement Then Return
        Dim et = EtudiantSelectionne()
        If et Is Nothing Then Return
        txtNom.Text = et.Nom
        txtPrenoms.Text = et.Prenoms
    End Sub

    Private Sub SurAjouter(sender As Object, e As EventArgs)
        Dim c = ClasseCourante()
        If c Is Nothing Then
            Ui.Info("Crée d'abord une classe, puis choisis-la dans la liste.")
            Return
        End If
        Dim nom = txtNom.Text.Trim()
        If nom = "" Then
            Ui.Erreur("Saisis le nom de l'étudiant.")
            Return
        End If
        Dim id = EtudiantDepot.Ajouter(c.Id, nom, txtPrenoms.Text.Trim())
        txtNom.Clear()
        txtPrenoms.Clear()
        ChargerEtudiants(id)
        txtNom.Focus()
    End Sub

    Private Sub SurModifier(sender As Object, e As EventArgs)
        Dim et = EtudiantSelectionne()
        If et Is Nothing Then
            Ui.Info("Sélectionne d'abord un étudiant dans la liste.")
            Return
        End If
        Dim nom = txtNom.Text.Trim()
        If nom = "" Then
            Ui.Erreur("Saisis le nom de l'étudiant.")
            Return
        End If
        EtudiantDepot.Modifier(et.Id, nom, txtPrenoms.Text.Trim())
        ChargerEtudiants(et.Id)
    End Sub

    Private Sub SurSupprimer(sender As Object, e As EventArgs)
        Dim et = EtudiantSelectionne()
        If et Is Nothing Then
            Ui.Info("Sélectionne d'abord un étudiant dans la liste.")
            Return
        End If
        If Not Ui.Confirmer($"Supprimer {et.NomComplet} ?" & vbCrLf & "Toutes ses notes seront supprimées.") Then Return
        EtudiantDepot.Supprimer(et.Id)
        txtNom.Clear()
        txtPrenoms.Clear()
        ChargerEtudiants()
    End Sub

End Class
