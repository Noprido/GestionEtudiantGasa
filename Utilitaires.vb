' ============================================================
' Utilitaires.vb  -  Formatage des nombres et petits outils d'interface
' ============================================================
Imports System.Diagnostics
Imports System.Drawing
Imports System.Globalization
Imports System.Windows.Forms

Public Module Outils

    Private ReadOnly Fr As New CultureInfo("fr-FR")

    ''' <summary>Affiche une moyenne : "12,50", ou "-" si absente.</summary>
    Public Function Fmt(valeur As Double?) As String
        If Not valeur.HasValue Then Return "-"
        Return valeur.Value.ToString("0.00", Fr)
    End Function

    ''' <summary>Affiche une note ou un coefficient : "12,5", "10", ou "" si absent.</summary>
    Public Function FmtNote(valeur As Double?) As String
        If Not valeur.HasValue Then Return ""
        Return valeur.Value.ToString("0.##", Fr)
    End Function

    ''' <summary>
    ''' Convertit un texte saisi en note (virgule ou point acceptés).
    ''' Texte vide -> Nothing avec valide = True. Texte invalide ou hors 0-20 -> valide = False.
    ''' </summary>
    Public Function ParseNote(texte As String, ByRef valide As Boolean) As Double?
        valide = True
        If String.IsNullOrWhiteSpace(texte) Then Return Nothing
        Dim v As Double
        If Double.TryParse(texte.Trim().Replace(",", "."), NumberStyles.Float, CultureInfo.InvariantCulture, v) _
           AndAlso Calculateur.NoteValide(v) Then
            Return v
        End If
        valide = False
        Return Nothing
    End Function

End Module

Public Module Ui

    Public Function Bouton(texte As String, largeur As Integer) As Button
        Return New Button With {.Text = texte, .Width = largeur, .Height = 32}
    End Function

    Public Function Etiquette(texte As String) As Label
        Return New Label With {.Text = texte, .AutoSize = True}
    End Function

    ''' <summary>Grille en lecture seule, sans ligne vide, sélection de ligne entière.</summary>
    Public Function Grille() As DataGridView
        Return New DataGridView With {
            .ReadOnly = True,
            .AllowUserToAddRows = False,
            .AllowUserToDeleteRows = False,
            .AllowUserToResizeRows = False,
            .RowHeadersVisible = False,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .MultiSelect = False,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            .BackgroundColor = Color.White,
            .BorderStyle = BorderStyle.FixedSingle,
            .Dock = DockStyle.Fill
        }
    End Function

    ''' <summary>
    ''' Ajoute les contrôles collés aux bords (Top/Bottom/Left) puis le contrôle central (Fill).
    ''' Le contrôle Fill doit être amené à l'avant pour que le découpage de l'espace soit correct.
    ''' </summary>
    Public Sub Empiler(parent As Control, remplissage As Control, ParamArray bords As Control())
        For Each b In bords
            parent.Controls.Add(b)
        Next
        parent.Controls.Add(remplissage)
        remplissage.BringToFront()
    End Sub

    Public Sub Info(message As String)
        MessageBox.Show(message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Public Sub Erreur(message As String)
        MessageBox.Show(message, "Attention", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    End Sub

    Public Function Confirmer(message As String) As Boolean
        Return MessageBox.Show(message, "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes
    End Function

    Public Sub OuvrirFichier(chemin As String)
        Try
            Process.Start(New ProcessStartInfo(chemin) With {.UseShellExecute = True})
        Catch ex As Exception
            Erreur("Impossible d'ouvrir le fichier : " & ex.Message)
        End Try
    End Sub

End Module
