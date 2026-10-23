' ============================================================
' FrmPrincipal.vb  -  Menu principal (formulaire de démarrage)
' Formulaire créé entièrement par code : pas besoin du concepteur visuel.
' ============================================================
Imports System.Drawing
Imports System.Windows.Forms

Public Class FrmPrincipal
    Inherits Form

    Public Sub New()
        Text = "Gestion des étudiants - UATM GASA FORMATION"
        StartPosition = FormStartPosition.CenterScreen
        ClientSize = New Size(480, 470)
        FormBorderStyle = FormBorderStyle.FixedSingle
        MaximizeBox = False
        Font = New Font("Segoe UI", 10)

        Dim entete As New Panel With {.Dock = DockStyle.Top, .Height = 110}
        Dim titre As New Label With {
            .Text = "UATM GASA FORMATION",
            .Font = New Font("Segoe UI", 18, FontStyle.Bold),
            .Dock = DockStyle.Top, .Height = 70, .TextAlign = ContentAlignment.MiddleCenter}
        Dim sousTitre As New Label With {
            .Text = "Gestion des notes et des bulletins",
            .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.TopCenter}
        Ui.Empiler(entete, sousTitre, titre)

        Dim menu As New FlowLayoutPanel With {
            .Dock = DockStyle.Fill, .FlowDirection = FlowDirection.TopDown,
            .WrapContents = False, .Padding = New Padding(90, 10, 0, 0)}

        AjouterEntree(menu, "Classes et matières de classe", Sub() Ouvrir(New FrmClasses()))
        AjouterEntree(menu, "Catalogue des matières", Sub() Ouvrir(New FrmMatieres()))
        AjouterEntree(menu, "Étudiants", Sub() Ouvrir(New FrmEtudiants()))
        AjouterEntree(menu, "Saisie des notes", Sub() Ouvrir(New FrmNotes()))
        AjouterEntree(menu, "Classement et bulletins", Sub() Ouvrir(New FrmClassement()))

        Ui.Empiler(Me, menu, entete)
        AddHandler Me.Load, AddressOf SurChargement
    End Sub

    Private Sub SurChargement(sender As Object, e As EventArgs)
        Try
            Db.Initialiser()
        Catch ex As Exception
            MessageBox.Show("Impossible d'initialiser la base de données :" & vbCrLf & ex.Message,
                            "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Close()
        End Try
    End Sub

    Private Sub AjouterEntree(parent As FlowLayoutPanel, texte As String, action As Action)
        Dim b = Ui.Bouton(texte, 300)
        b.Height = 44
        b.Margin = New Padding(0, 0, 0, 12)
        AddHandler b.Click, Sub(s, e) action()
        parent.Controls.Add(b)
    End Sub

    Private Sub Ouvrir(f As Form)
        Using f
            f.ShowDialog(Me)
        End Using
    End Sub

End Class
