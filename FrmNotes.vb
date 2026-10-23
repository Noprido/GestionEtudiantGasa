' ============================================================
' FrmNotes.vb  -  Saisie des notes : on choisit une classe et une matière,
' puis on saisit Interro 1, Interro 2 et Devoir de chaque étudiant.
' La moyenne et le statut (Validé / Non validé) se calculent en direct.
' Les champs peuvent rester vides (1 seule interro autorisée).
' ============================================================
Imports System.Drawing
Imports System.Windows.Forms

Public Class FrmNotes
    Inherits Form

    Private ReadOnly cboClasse As New ComboBox()
    Private ReadOnly cboMatiere As New ComboBox()
    Private ReadOnly grille As DataGridView = Ui.Grille()
    Private ReadOnly lblEtat As New Label()

    Private _classeIdCharge As Integer
    Private _matiereIdCharge As Integer
    Private _modifie As Boolean
    Private _enChargement As Boolean

    Public Sub New()
        Text = "Saisie des notes"
        StartPosition = FormStartPosition.CenterParent
        ClientSize = New Size(900, 580)
        MinimumSize = New Size(760, 460)
        Font = New Font("Segoe UI", 10)
        AutoValidate = AutoValidate.EnableAllowFocusChange

        ' Grille modifiable (sauf nom, moyenne et statut)
        grille.ReadOnly = False
        grille.SelectionMode = DataGridViewSelectionMode.CellSelect
        grille.Columns.Add("Etudiant", "Étudiant")
        grille.Columns.Add("Interro1", "Interro 1")
        grille.Columns.Add("Interro2", "Interro 2")
        grille.Columns.Add("Devoir", "Devoir")
        grille.Columns.Add("Moyenne", "Moyenne")
        grille.Columns.Add("Statut", "Statut")
        For Each i In {0, 4, 5}
            grille.Columns(i).ReadOnly = True
        Next
        For i = 1 To 5
            grille.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        Next
        grille.Columns(0).FillWeight = 220

        Dim haut As New Panel With {.Dock = DockStyle.Top, .Height = 56}
        Dim lblClasse = Ui.Etiquette("Classe :")
        lblClasse.Location = New Point(12, 18)
        cboClasse.Location = New Point(75, 14)
        cboClasse.Width = 260
        cboClasse.DropDownStyle = ComboBoxStyle.DropDownList
        Dim lblMatiere = Ui.Etiquette("Matière :")
        lblMatiere.Location = New Point(360, 18)
        cboMatiere.Location = New Point(430, 14)
        cboMatiere.Width = 280
        cboMatiere.DropDownStyle = ComboBoxStyle.DropDownList
        haut.Controls.AddRange(New Control() {lblClasse, cboClasse, lblMatiere, cboMatiere})

        Dim bas As New Panel With {.Dock = DockStyle.Bottom, .Height = 54}
        Dim btnEnregistrer = Ui.Bouton("Enregistrer", 150)
        btnEnregistrer.Location = New Point(12, 11)
        lblEtat.AutoSize = True
        lblEtat.Location = New Point(180, 18)
        bas.Controls.AddRange(New Control() {btnEnregistrer, lblEtat})

        Ui.Empiler(Me, grille, haut, bas)

        AddHandler cboClasse.SelectedIndexChanged, AddressOf SurClasseChange
        AddHandler cboMatiere.SelectedIndexChanged, AddressOf SurMatiereChange
        AddHandler grille.CellValidating, AddressOf SurValidation
        AddHandler grille.CellEndEdit, AddressOf SurFinEdition
        AddHandler btnEnregistrer.Click, AddressOf SurEnregistrer
        AddHandler Me.FormClosing, AddressOf SurFermeture
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
        ChargerMatieres()
    End Sub

    Private Function ClasseCourante() As Classe
        Return TryCast(cboClasse.SelectedItem, Classe)
    End Function

    Private Function MatiereCourante() As Matiere
        Return TryCast(cboMatiere.SelectedItem, Matiere)
    End Function

    ' ---------- Changement de classe / matière ----------

    Private Sub SurClasseChange(sender As Object, e As EventArgs)
        If _enChargement Then Return
        AvantChangement()
        ChargerMatieres()
    End Sub

    Private Sub SurMatiereChange(sender As Object, e As EventArgs)
        If _enChargement Then Return
        AvantChangement()
        ChargerNotes()
    End Sub

    ''' <summary>Propose d'enregistrer la saisie en cours avant de changer de classe ou de matière.</summary>
    Private Sub AvantChangement()
        If Not _modifie Then Return
        If Ui.Confirmer("Enregistrer les modifications de la saisie en cours ?") Then Enregistrer()
        _modifie = False
    End Sub

    Private Sub ChargerMatieres()
        _enChargement = True
        cboMatiere.Items.Clear()
        Dim c = ClasseCourante()
        If c IsNot Nothing Then
            For Each m In MatiereDepot.ListerPourClasse(c.Id)
                cboMatiere.Items.Add(m)
            Next
            If cboMatiere.Items.Count > 0 Then cboMatiere.SelectedIndex = 0
        End If
        _enChargement = False
        ChargerNotes()
    End Sub

    Private Sub ChargerNotes()
        grille.Rows.Clear()
        _classeIdCharge = 0
        _matiereIdCharge = 0
        _modifie = False
        lblEtat.Text = ""

        Dim c = ClasseCourante()
        Dim m = MatiereCourante()
        If c Is Nothing OrElse m Is Nothing Then Return

        Dim notes = NoteDepot.ListerParClasse(c.Id).
            Where(Function(n) n.MatiereId = m.Id).
            ToDictionary(Function(n) n.EtudiantId)

        _enChargement = True
        For Each et In EtudiantDepot.ListerParClasse(c.Id)
            Dim note As NoteMatiere = Nothing
            notes.TryGetValue(et.Id, note)
            Dim i = grille.Rows.Add(et.NomComplet,
                                    Outils.FmtNote(note?.Interro1),
                                    Outils.FmtNote(note?.Interro2),
                                    Outils.FmtNote(note?.Devoir), "", "")
            grille.Rows(i).Tag = et.Id
            Recalculer(grille.Rows(i))
        Next
        _enChargement = False

        _classeIdCharge = c.Id
        _matiereIdCharge = m.Id
        If grille.Rows.Count = 0 Then lblEtat.Text = "Aucun étudiant dans cette classe."
    End Sub

    ' ---------- Saisie ----------

    Private Function TexteCellule(ligne As DataGridViewRow, colonne As Integer) As String
        Return Convert.ToString(ligne.Cells(colonne).Value)
    End Function

    Private Sub SurValidation(sender As Object, e As DataGridViewCellValidatingEventArgs)
        If e.ColumnIndex < 1 OrElse e.ColumnIndex > 3 Then Return
        Dim ok As Boolean
        Outils.ParseNote(Convert.ToString(e.FormattedValue), ok)
        If Not ok Then
            e.Cancel = True
            Ui.Erreur("La note doit être un nombre compris entre 0 et 20.")
        End If
    End Sub

    Private Sub SurFinEdition(sender As Object, e As DataGridViewCellEventArgs)
        If _enChargement OrElse e.RowIndex < 0 Then Return
        Recalculer(grille.Rows(e.RowIndex))
        _modifie = True
        lblEtat.Text = "Modifications non enregistrées"
    End Sub

    Private Sub Recalculer(ligne As DataGridViewRow)
        Dim ok As Boolean
        Dim note As New NoteMatiere()
        note.Interro1 = Outils.ParseNote(TexteCellule(ligne, 1), ok)
        note.Interro2 = Outils.ParseNote(TexteCellule(ligne, 2), ok)
        note.Devoir = Outils.ParseNote(TexteCellule(ligne, 3), ok)

        Dim moyenne = Calculateur.MoyenneMatiere(note)
        ligne.Cells(4).Value = Outils.Fmt(moyenne)
        ligne.Cells(5).Value = Calculateur.StatutMatiere(moyenne)

        Dim couleur As Color
        If Not moyenne.HasValue Then
            couleur = Color.Gray
        ElseIf Calculateur.EstValidee(moyenne) Then
            couleur = Color.DarkGreen
        Else
            couleur = Color.Firebrick
        End If
        ligne.Cells(4).Style.ForeColor = couleur
        ligne.Cells(5).Style.ForeColor = couleur
    End Sub

    ' ---------- Enregistrement ----------

    Private Sub SurEnregistrer(sender As Object, e As EventArgs)
        If _classeIdCharge = 0 OrElse _matiereIdCharge = 0 Then
            Ui.Info("Choisis une classe et une matière.")
            Return
        End If
        If Enregistrer() Then lblEtat.Text = "Notes enregistrées à " & DateTime.Now.ToString("HH:mm:ss")
    End Sub

    Private Function Enregistrer() As Boolean
        If _classeIdCharge = 0 OrElse _matiereIdCharge = 0 Then Return True
        If Not grille.EndEdit() Then Return False

        Dim lot As New List(Of NoteMatiere)
        For Each ligne As DataGridViewRow In grille.Rows
            Dim ok1, ok2, ok3 As Boolean
            Dim note As New NoteMatiere With {.EtudiantId = CInt(ligne.Tag), .MatiereId = _matiereIdCharge}
            note.Interro1 = Outils.ParseNote(TexteCellule(ligne, 1), ok1)
            note.Interro2 = Outils.ParseNote(TexteCellule(ligne, 2), ok2)
            note.Devoir = Outils.ParseNote(TexteCellule(ligne, 3), ok3)
            If Not (ok1 AndAlso ok2 AndAlso ok3) Then
                Ui.Erreur("Note invalide pour " & TexteCellule(ligne, 0) & " (valeur entre 0 et 20).")
                Return False
            End If
            lot.Add(note)
        Next

        Try
            NoteDepot.EnregistrerLot(lot)
        Catch ex As Exception
            Ui.Erreur("Impossible d'enregistrer les notes :" & vbCrLf & ex.Message)
            Return False
        End Try
        _modifie = False
        Return True
    End Function

    Private Sub SurFermeture(sender As Object, e As FormClosingEventArgs)
        If Not _modifie Then Return
        Select Case MessageBox.Show("Enregistrer les modifications avant de fermer ?", "Saisie des notes",
                                    MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            Case DialogResult.Yes
                If Not Enregistrer() Then e.Cancel = True
            Case DialogResult.Cancel
                e.Cancel = True
        End Select
    End Sub

End Class
