' ============================================================
' Depots.vb  -  Accès aux données (requêtes SQL)
' Contient : MatiereDepot, ClasseDepot, EtudiantDepot, NoteDepot
' ============================================================
Imports Microsoft.Data.Sqlite

''' <summary>Construction des commandes SQL avec paramètres : Commande(cn, sql, "$nom", valeur, ...)</summary>
Friend Module Sql

    Friend Function Commande(cn As SqliteConnection, texte As String, ParamArray params As Object()) As SqliteCommand
        Dim cmd = cn.CreateCommand()
        cmd.CommandText = texte
        For i = 0 To params.Length - 2 Step 2
            Dim valeur As Object = params(i + 1)
            If valeur Is Nothing Then valeur = DBNull.Value
            cmd.Parameters.AddWithValue(CStr(params(i)), valeur)
        Next
        Return cmd
    End Function

    Friend Function CommandeTx(cn As SqliteConnection, tx As SqliteTransaction, texte As String, ParamArray params As Object()) As SqliteCommand
        Dim cmd = Commande(cn, texte, params)
        cmd.Transaction = tx
        Return cmd
    End Function

End Module

' ------------------------------------------------------------
Public Module MatiereDepot

    Private Function LireMatieres(cmd As SqliteCommand) As List(Of Matiere)
        Dim liste As New List(Of Matiere)
        Using r = cmd.ExecuteReader()
            While r.Read()
                liste.Add(New Matiere With {
                    .Id = r.GetInt32(0), .Nom = r.GetString(1), .Coefficient = r.GetDouble(2)})
            End While
        End Using
        Return liste
    End Function

    ''' <summary>Toutes les matières du catalogue.</summary>
    Public Function Lister() As List(Of Matiere)
        Using cn = Db.Ouvrir()
            Return LireMatieres(Commande(cn, "SELECT Id, Nom, Coefficient FROM Matiere ORDER BY Id"))
        End Using
    End Function

    ''' <summary>Matières rattachées à une classe.</summary>
    Public Function ListerPourClasse(classeId As Integer) As List(Of Matiere)
        Using cn = Db.Ouvrir()
            Return LireMatieres(Commande(cn,
                "SELECT m.Id, m.Nom, m.Coefficient FROM Matiere m " &
                "JOIN ClasseMatiere cm ON cm.MatiereId = m.Id " &
                "WHERE cm.ClasseId = $c ORDER BY m.Id", "$c", classeId))
        End Using
    End Function

    ''' <summary>Matières du catalogue pas encore rattachées à la classe.</summary>
    Public Function ListerHorsClasse(classeId As Integer) As List(Of Matiere)
        Using cn = Db.Ouvrir()
            Return LireMatieres(Commande(cn,
                "SELECT Id, Nom, Coefficient FROM Matiere " &
                "WHERE Id NOT IN (SELECT MatiereId FROM ClasseMatiere WHERE ClasseId = $c) " &
                "ORDER BY Nom", "$c", classeId))
        End Using
    End Function

    Public Function NomExiste(nom As String, Optional exceptId As Integer = 0) As Boolean
        Using cn = Db.Ouvrir()
            Return Convert.ToInt64(Commande(cn,
                "SELECT COUNT(*) FROM Matiere WHERE Nom = $n AND Id <> $id",
                "$n", nom, "$id", exceptId).ExecuteScalar()) > 0
        End Using
    End Function

    Public Function Ajouter(nom As String, coefficient As Double) As Integer
        Using cn = Db.Ouvrir()
            Commande(cn, "INSERT INTO Matiere (Nom, Coefficient) VALUES ($n, $c)",
                     "$n", nom, "$c", coefficient).ExecuteNonQuery()
            Return Convert.ToInt32(Commande(cn, "SELECT last_insert_rowid()").ExecuteScalar())
        End Using
    End Function

    Public Sub Modifier(id As Integer, nom As String, coefficient As Double)
        Using cn = Db.Ouvrir()
            Commande(cn, "UPDATE Matiere SET Nom = $n, Coefficient = $c WHERE Id = $id",
                     "$n", nom, "$c", coefficient, "$id", id).ExecuteNonQuery()
        End Using
    End Sub

    ''' <summary>Vrai si la matière est utilisée par au moins une classe.</summary>
    Public Function EstUtilisee(id As Integer) As Boolean
        Using cn = Db.Ouvrir()
            Return Convert.ToInt64(Commande(cn,
                "SELECT COUNT(*) FROM ClasseMatiere WHERE MatiereId = $id", "$id", id).ExecuteScalar()) > 0
        End Using
    End Function

    Public Sub Supprimer(id As Integer)
        Using cn = Db.Ouvrir()
            Commande(cn, "DELETE FROM Matiere WHERE Id = $id", "$id", id).ExecuteNonQuery()
        End Using
    End Sub

End Module

' ------------------------------------------------------------
Public Module ClasseDepot

    Private Function LireClasse(r As SqliteDataReader) As Classe
        Return New Classe With {.Id = r.GetInt32(0), .Nom = r.GetString(1), .AnneeScolaire = r.GetString(2)}
    End Function

    Public Function Lister() As List(Of Classe)
        Dim liste As New List(Of Classe)
        Using cn = Db.Ouvrir()
            Using r = Commande(cn, "SELECT Id, Nom, AnneeScolaire FROM Classe ORDER BY Nom").ExecuteReader()
                While r.Read()
                    liste.Add(LireClasse(r))
                End While
            End Using
        End Using
        Return liste
    End Function

    Public Function Obtenir(id As Integer) As Classe
        Using cn = Db.Ouvrir()
            Using r = Commande(cn, "SELECT Id, Nom, AnneeScolaire FROM Classe WHERE Id = $id", "$id", id).ExecuteReader()
                If r.Read() Then Return LireClasse(r)
            End Using
        End Using
        Return Nothing
    End Function

    Public Function NomExiste(nom As String, Optional exceptId As Integer = 0) As Boolean
        Using cn = Db.Ouvrir()
            Return Convert.ToInt64(Commande(cn,
                "SELECT COUNT(*) FROM Classe WHERE Nom = $n AND Id <> $id",
                "$n", nom, "$id", exceptId).ExecuteScalar()) > 0
        End Using
    End Function

    Public Function Ajouter(nom As String, anneeScolaire As String) As Integer
        Using cn = Db.Ouvrir()
            Commande(cn, "INSERT INTO Classe (Nom, AnneeScolaire) VALUES ($n, $a)",
                     "$n", nom, "$a", anneeScolaire).ExecuteNonQuery()
            Return Convert.ToInt32(Commande(cn, "SELECT last_insert_rowid()").ExecuteScalar())
        End Using
    End Function

    Public Sub Modifier(id As Integer, nom As String, anneeScolaire As String)
        Using cn = Db.Ouvrir()
            Commande(cn, "UPDATE Classe SET Nom = $n, AnneeScolaire = $a WHERE Id = $id",
                     "$n", nom, "$a", anneeScolaire, "$id", id).ExecuteNonQuery()
        End Using
    End Sub

    ''' <summary>Supprime la classe, ses étudiants, leurs notes et ses liens avec les matières.</summary>
    Public Sub Supprimer(id As Integer)
        Using cn = Db.Ouvrir()
            Commande(cn, "DELETE FROM Classe WHERE Id = $id", "$id", id).ExecuteNonQuery()
        End Using
    End Sub

    Public Function MatiereDejaPresente(classeId As Integer, matiereId As Integer) As Boolean
        Using cn = Db.Ouvrir()
            Return Convert.ToInt64(Commande(cn,
                "SELECT COUNT(*) FROM ClasseMatiere WHERE ClasseId = $c AND MatiereId = $m",
                "$c", classeId, "$m", matiereId).ExecuteScalar()) > 0
        End Using
    End Function

    Public Sub AjouterMatiere(classeId As Integer, matiereId As Integer)
        Using cn = Db.Ouvrir()
            Commande(cn, "INSERT OR IGNORE INTO ClasseMatiere (ClasseId, MatiereId) VALUES ($c, $m)",
                     "$c", classeId, "$m", matiereId).ExecuteNonQuery()
        End Using
    End Sub

    ''' <summary>Retire la matière de la classe et supprime les notes correspondantes des étudiants de cette classe.</summary>
    Public Sub RetirerMatiere(classeId As Integer, matiereId As Integer)
        Using cn = Db.Ouvrir()
            Using tx = cn.BeginTransaction()
                CommandeTx(cn, tx,
                    "DELETE FROM Note WHERE MatiereId = $m " &
                    "AND EtudiantId IN (SELECT Id FROM Etudiant WHERE ClasseId = $c)",
                    "$m", matiereId, "$c", classeId).ExecuteNonQuery()
                CommandeTx(cn, tx,
                    "DELETE FROM ClasseMatiere WHERE ClasseId = $c AND MatiereId = $m",
                    "$c", classeId, "$m", matiereId).ExecuteNonQuery()
                tx.Commit()
            End Using
        End Using
    End Sub

End Module

' ------------------------------------------------------------
Public Module EtudiantDepot

    Public Function ListerParClasse(classeId As Integer) As List(Of Etudiant)
        Dim liste As New List(Of Etudiant)
        Using cn = Db.Ouvrir()
            Using r = Commande(cn,
                "SELECT Id, ClasseId, Nom, Prenoms FROM Etudiant WHERE ClasseId = $c ORDER BY Nom, Prenoms",
                "$c", classeId).ExecuteReader()
                While r.Read()
                    liste.Add(New Etudiant With {
                        .Id = r.GetInt32(0), .ClasseId = r.GetInt32(1),
                        .Nom = r.GetString(2), .Prenoms = r.GetString(3)})
                End While
            End Using
        End Using
        Return liste
    End Function

    Public Function Ajouter(classeId As Integer, nom As String, prenoms As String) As Integer
        Using cn = Db.Ouvrir()
            Commande(cn, "INSERT INTO Etudiant (ClasseId, Nom, Prenoms) VALUES ($c, $n, $p)",
                     "$c", classeId, "$n", nom, "$p", prenoms).ExecuteNonQuery()
            Return Convert.ToInt32(Commande(cn, "SELECT last_insert_rowid()").ExecuteScalar())
        End Using
    End Function

    Public Sub Modifier(id As Integer, nom As String, prenoms As String)
        Using cn = Db.Ouvrir()
            Commande(cn, "UPDATE Etudiant SET Nom = $n, Prenoms = $p WHERE Id = $id",
                     "$n", nom, "$p", prenoms, "$id", id).ExecuteNonQuery()
        End Using
    End Sub

    ''' <summary>Supprime l'étudiant et toutes ses notes.</summary>
    Public Sub Supprimer(id As Integer)
        Using cn = Db.Ouvrir()
            Commande(cn, "DELETE FROM Etudiant WHERE Id = $id", "$id", id).ExecuteNonQuery()
        End Using
    End Sub

End Module

' ------------------------------------------------------------
Public Module NoteDepot

    ''' <summary>Toutes les notes de tous les étudiants d'une classe.</summary>
    Public Function ListerParClasse(classeId As Integer) As List(Of NoteMatiere)
        Dim liste As New List(Of NoteMatiere)
        Using cn = Db.Ouvrir()
            Using r = Commande(cn,
                "SELECT n.Id, n.EtudiantId, n.MatiereId, n.Interro1, n.Interro2, n.Devoir " &
                "FROM Note n JOIN Etudiant e ON e.Id = n.EtudiantId WHERE e.ClasseId = $c",
                "$c", classeId).ExecuteReader()
                While r.Read()
                    liste.Add(New NoteMatiere With {
                        .Id = r.GetInt32(0), .EtudiantId = r.GetInt32(1), .MatiereId = r.GetInt32(2),
                        .Interro1 = Db.LireDouble(r, 3), .Interro2 = Db.LireDouble(r, 4), .Devoir = Db.LireDouble(r, 5)})
                End While
            End Using
        End Using
        Return liste
    End Function

    ''' <summary>Enregistre (insère ou met à jour) un lot de notes en une seule transaction.</summary>
    Public Sub EnregistrerLot(notes As IEnumerable(Of NoteMatiere))
        Using cn = Db.Ouvrir()
            Using tx = cn.BeginTransaction()
                For Each n In notes
                    CommandeTx(cn, tx,
                        "INSERT INTO Note (EtudiantId, MatiereId, Interro1, Interro2, Devoir) " &
                        "VALUES ($e, $m, $i1, $i2, $d) " &
                        "ON CONFLICT(EtudiantId, MatiereId) DO UPDATE SET " &
                        "Interro1 = excluded.Interro1, Interro2 = excluded.Interro2, Devoir = excluded.Devoir",
                        "$e", n.EtudiantId, "$m", n.MatiereId,
                        "$i1", Db.VersDb(n.Interro1), "$i2", Db.VersDb(n.Interro2), "$d", Db.VersDb(n.Devoir)).ExecuteNonQuery()
                Next
                tx.Commit()
            End Using
        End Using
    End Sub

End Module
