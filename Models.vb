' ============================================================
' Models.vb  -  Modèles de données (une classe par table SQLite)
' ============================================================

''' <summary>Matière du catalogue central (nom + coefficient).</summary>
Public Class Matiere
    Public Property Id As Integer
    Public Property Nom As String = ""
    Public Property Coefficient As Double

    ' Pratique pour l'affichage dans une ComboBox / ListBox
    Public Overrides Function ToString() As String
        Return $"{Nom} (coef {Coefficient})"
    End Function
End Class

''' <summary>Classe (ex. : "Licence 1 Informatique").</summary>
Public Class Classe
    Public Property Id As Integer
    Public Property Nom As String = ""
    Public Property AnneeScolaire As String = ""

    Public Overrides Function ToString() As String
        Return Nom
    End Function
End Class

''' <summary>Étudiant rattaché à une classe.</summary>
Public Class Etudiant
    Public Property Id As Integer
    Public Property ClasseId As Integer
    Public Property Nom As String = ""
    Public Property Prenoms As String = ""

    Public ReadOnly Property NomComplet As String
        Get
            Return $"{Nom} {Prenoms}".Trim()
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return NomComplet
    End Function
End Class

''' <summary>Notes d'un étudiant dans une matière (null = pas encore saisie).</summary>
Public Class NoteMatiere
    Public Property Id As Integer
    Public Property EtudiantId As Integer
    Public Property MatiereId As Integer
    Public Property Interro1 As Double?
    Public Property Interro2 As Double?
    Public Property Devoir As Double?
End Class
