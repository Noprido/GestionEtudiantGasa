' ============================================================
' Calculateur.vb  -  Toute la logique de calcul (aucun accès base de données)
' ============================================================

Public Module Calculateur

    ''' <summary>Moyenne minimale pour valider une matière.</summary>
    Public Const SeuilValidation As Double = 10.0

    ''' <summary>Une note facultative est valide si vide ou comprise entre 0 et 20.</summary>
    Public Function NoteValide(note As Double?) As Boolean
        Return Not note.HasValue OrElse (note.Value >= 0 AndAlso note.Value <= 20)
    End Function

    ''' <summary>Arrondi à 2 décimales (0,005 arrondi vers le haut).</summary>
    Public Function Arrondir(valeur As Double) As Double
        Return Math.Round(valeur, 2, MidpointRounding.AwayFromZero)
    End Function

    ''' <summary>
    ''' Moyenne des interros : 2 notes -> leur moyenne ;
    ''' 1 seule note -> utilisée telle quelle ; aucune -> Nothing.
    ''' </summary>
    Public Function MoyenneInterros(interro1 As Double?, interro2 As Double?) As Double?
        If interro1.HasValue AndAlso interro2.HasValue Then
            Return (interro1.Value + interro2.Value) / 2
        End If
        If interro1.HasValue Then Return interro1.Value
        If interro2.HasValue Then Return interro2.Value
        Return Nothing
    End Function

    ''' <summary>
    ''' Moyenne d'une matière = (moyenne des interros + devoir) / 2.
    ''' Nothing si l'interro ou le devoir manque (matière incomplète).
    ''' </summary>
    Public Function MoyenneMatiere(note As NoteMatiere) As Double?
        Dim moyInterros = MoyenneInterros(note.Interro1, note.Interro2)
        If Not moyInterros.HasValue OrElse Not note.Devoir.HasValue Then Return Nothing
        Return Arrondir((moyInterros.Value + note.Devoir.Value) / 2)
    End Function

    ''' <summary>Vrai si la moyenne de la matière atteint le seuil (10/20).</summary>
    Public Function EstValidee(moyenneMatiere As Double?) As Boolean
        Return moyenneMatiere.HasValue AndAlso moyenneMatiere.Value >= SeuilValidation
    End Function

    ''' <summary>Texte affiché : "Validé", "Non validé" ou "Incomplet".</summary>
    Public Function StatutMatiere(moyenneMatiere As Double?) As String
        If Not moyenneMatiere.HasValue Then Return "Incomplet"
        Return If(EstValidee(moyenneMatiere), "Validé", "Non validé")
    End Function

    ''' <summary>
    ''' Moyenne générale = somme(moyenne x coefficient) / somme(coefficients).
    ''' Nothing si une matière est incomplète ou s'il n'y a aucune matière.
    ''' </summary>
    Public Function MoyenneGenerale(lignes As IEnumerable(Of (Moyenne As Double?, Coefficient As Double))) As Double?
        Dim somme = 0.0
        Dim totalCoef = 0.0
        For Each ligne In lignes
            If Not ligne.Moyenne.HasValue Then Return Nothing
            somme += ligne.Moyenne.Value * ligne.Coefficient
            totalCoef += ligne.Coefficient
        Next
        If totalCoef <= 0 Then Return Nothing
        Return Arrondir(somme / totalCoef)
    End Function

    ''' <summary>Mention générale selon la moyenne générale.</summary>
    Public Function Mention(moyenneGenerale As Double?) As String
        If Not moyenneGenerale.HasValue Then Return "-"
        Dim m = moyenneGenerale.Value
        If m >= 18 Then Return "Excellent"
        If m >= 16 Then Return "Très bien"
        If m >= 14 Then Return "Bien"
        If m >= 12 Then Return "Assez bien"
        If m >= 10 Then Return "Passable"
        Return "Insuffisant"
    End Function

End Module
