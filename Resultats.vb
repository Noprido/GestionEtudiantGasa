' ============================================================
' Resultats.vb  -  Résultats par étudiant, classement par ordre de mérite (ex aequo gérés)
' ============================================================

''' <summary>Une matière sur le bulletin d'un étudiant.</summary>
Public Class LigneMatiere
    Public Property MatiereId As Integer
    Public Property Nom As String = ""
    Public Property Coefficient As Double
    Public Property Interro1 As Double?
    Public Property Interro2 As Double?
    Public Property Devoir As Double?
    Public Property Moyenne As Double?

    Public ReadOnly Property Statut As String
        Get
            Return Calculateur.StatutMatiere(Moyenne)
        End Get
    End Property
End Class

''' <summary>Résultat complet d'un étudiant.</summary>
Public Class ResultatEtudiant
    Public Property Etudiant As Etudiant = New Etudiant()
    Public Property Lignes As New List(Of LigneMatiere)
    Public Property MoyenneGenerale As Double?
    Public Property Mention As String = "-"
    ''' <summary>Nothing = non classé (notes manquantes).</summary>
    Public Property Rang As Integer?
    Public Property ExAequo As Boolean
    Public Property Effectif As Integer

    Public ReadOnly Property RangTexte As String
        Get
            If Not Rang.HasValue Then Return "-"
            Dim suffixe = If(Rang.Value = 1, "er", "e")
            Return Rang.Value.ToString() & suffixe & If(ExAequo, " ex aequo", "")
        End Get
    End Property
End Class

''' <summary>Résultats d'une classe entière, triés par ordre de mérite.</summary>
Public Class ResultatClasse
    Public Property Classe As Classe = New Classe()
    Public Property Matieres As New List(Of Matiere)
    Public Property Resultats As New List(Of ResultatEtudiant)
    Public Property MoyenneClasse As Double?
    Public Property MoyenneMax As Double?
    Public Property MoyenneMin As Double?
End Class

Public Module Classement

    ''' <summary>
    ''' Calcule les moyennes, mentions et rangs de tous les étudiants d'une classe.
    ''' Classement : moyenne générale décroissante ; ex aequo = même rang (1, 2, 2, 4...).
    ''' Les étudiants avec notes manquantes sont placés à la fin, sans rang.
    ''' </summary>
    Public Function Calculer(classeId As Integer) As ResultatClasse
        Dim rc As New ResultatClasse()
        rc.Classe = If(ClasseDepot.Obtenir(classeId), New Classe())
        rc.Matieres = MatiereDepot.ListerPourClasse(classeId)

        Dim etudiants = EtudiantDepot.ListerParClasse(classeId)
        Dim notes = NoteDepot.ListerParClasse(classeId)
        Dim parCle As New Dictionary(Of String, NoteMatiere)
        For Each n In notes
            parCle(n.EtudiantId & "|" & n.MatiereId) = n
        Next

        Dim tous As New List(Of ResultatEtudiant)
        For Each e In etudiants
            Dim res As New ResultatEtudiant With {.Etudiant = e, .Effectif = etudiants.Count}
            For Each m In rc.Matieres
                Dim note As NoteMatiere = Nothing
                parCle.TryGetValue(e.Id & "|" & m.Id, note)
                If note Is Nothing Then note = New NoteMatiere()
                Dim ligne As New LigneMatiere With {
                    .MatiereId = m.Id, .Nom = m.Nom, .Coefficient = m.Coefficient,
                    .Interro1 = note.Interro1, .Interro2 = note.Interro2, .Devoir = note.Devoir}
                ligne.Moyenne = Calculateur.MoyenneMatiere(note)
                res.Lignes.Add(ligne)
            Next
            res.MoyenneGenerale = Calculateur.MoyenneGenerale(res.Lignes.Select(Function(l) (l.Moyenne, l.Coefficient)))
            res.Mention = Calculateur.Mention(res.MoyenneGenerale)
            tous.Add(res)
        Next

        Dim avecMoyenne = tous.Where(Function(r) r.MoyenneGenerale.HasValue).
            OrderByDescending(Function(r) r.MoyenneGenerale.Value).
            ThenBy(Function(r) r.Etudiant.Nom).
            ThenBy(Function(r) r.Etudiant.Prenoms).ToList()
        Dim sansMoyenne = tous.Where(Function(r) Not r.MoyenneGenerale.HasValue).
            OrderBy(Function(r) r.Etudiant.Nom).
            ThenBy(Function(r) r.Etudiant.Prenoms).ToList()

        For Each r In avecMoyenne
            Dim valeur = r.MoyenneGenerale.Value
            Dim meilleurs = avecMoyenne.Where(Function(x) x.MoyenneGenerale.Value > valeur + 0.000001).Count()
            Dim identiques = avecMoyenne.Where(Function(x) Math.Abs(x.MoyenneGenerale.Value - valeur) < 0.000001).Count()
            r.Rang = meilleurs + 1
            r.ExAequo = identiques > 1
        Next

        If avecMoyenne.Count > 0 Then
            rc.MoyenneClasse = Calculateur.Arrondir(avecMoyenne.Average(Function(x) x.MoyenneGenerale.Value))
            rc.MoyenneMax = avecMoyenne.Max(Function(x) x.MoyenneGenerale.Value)
            rc.MoyenneMin = avecMoyenne.Min(Function(x) x.MoyenneGenerale.Value)
        End If

        rc.Resultats = avecMoyenne.Concat(sansMoyenne).ToList()
        Return rc
    End Function

End Module
