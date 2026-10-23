' ============================================================
' Db.vb  -  Accès SQLite : connexion, création du schéma, données par défaut
' Package NuGet requis : Microsoft.Data.Sqlite
' ============================================================
Imports System.IO
Imports Microsoft.Data.Sqlite

Public Module Db

    ' Fichier : %LocalAppData%\UATM_Bulletins\notes.db
    Private ReadOnly CheminBase As String = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "UATM_Bulletins", "notes.db")

    Private Const Schema As String =
        "CREATE TABLE IF NOT EXISTS Matiere (" &
        "  Id INTEGER PRIMARY KEY AUTOINCREMENT," &
        "  Nom TEXT NOT NULL UNIQUE COLLATE NOCASE," &
        "  Coefficient REAL NOT NULL CHECK (Coefficient > 0));" &
        "CREATE TABLE IF NOT EXISTS Classe (" &
        "  Id INTEGER PRIMARY KEY AUTOINCREMENT," &
        "  Nom TEXT NOT NULL UNIQUE COLLATE NOCASE," &
        "  AnneeScolaire TEXT NOT NULL DEFAULT '');" &
        "CREATE TABLE IF NOT EXISTS ClasseMatiere (" &
        "  ClasseId INTEGER NOT NULL REFERENCES Classe(Id) ON DELETE CASCADE," &
        "  MatiereId INTEGER NOT NULL REFERENCES Matiere(Id) ON DELETE RESTRICT," &
        "  PRIMARY KEY (ClasseId, MatiereId));" &
        "CREATE TABLE IF NOT EXISTS Etudiant (" &
        "  Id INTEGER PRIMARY KEY AUTOINCREMENT," &
        "  ClasseId INTEGER NOT NULL REFERENCES Classe(Id) ON DELETE CASCADE," &
        "  Nom TEXT NOT NULL," &
        "  Prenoms TEXT NOT NULL DEFAULT '');" &
        "CREATE TABLE IF NOT EXISTS Note (" &
        "  Id INTEGER PRIMARY KEY AUTOINCREMENT," &
        "  EtudiantId INTEGER NOT NULL REFERENCES Etudiant(Id) ON DELETE CASCADE," &
        "  MatiereId INTEGER NOT NULL REFERENCES Matiere(Id) ON DELETE RESTRICT," &
        "  Interro1 REAL CHECK (Interro1 IS NULL OR (Interro1 >= 0 AND Interro1 <= 20))," &
        "  Interro2 REAL CHECK (Interro2 IS NULL OR (Interro2 >= 0 AND Interro2 <= 20))," &
        "  Devoir REAL CHECK (Devoir IS NULL OR (Devoir >= 0 AND Devoir <= 20))," &
        "  UNIQUE (EtudiantId, MatiereId));"

    ''' <summary>Ouvre une connexion (clés étrangères activées).</summary>
    Public Function Ouvrir() As SqliteConnection
        Dim cn As New SqliteConnection($"Data Source={CheminBase}")
        cn.Open()
        Using cmd = cn.CreateCommand()
            cmd.CommandText = "PRAGMA foreign_keys = ON;"
            cmd.ExecuteNonQuery()
        End Using
        Return cn
    End Function

    ''' <summary>À appeler une fois au démarrage de l'application.</summary>
    Public Sub Initialiser()
        Directory.CreateDirectory(Path.GetDirectoryName(CheminBase))

        Using cn = Ouvrir()
            Using cmd = cn.CreateCommand()
                cmd.CommandText = Schema
                cmd.ExecuteNonQuery()
            End Using

            ' Matières de base, insérées seulement si le catalogue est vide
            Using cmd = cn.CreateCommand()
                cmd.CommandText = "SELECT COUNT(*) FROM Matiere"
                If Convert.ToInt64(cmd.ExecuteScalar()) = 0 Then
                    Dim defauts As (Nom As String, Coef As Double)() = {
                        ("Tronc commun", 3.0),
                        ("Spécialité", 4.0),
                        ("EPP", 3.0)
                    }
                    For Each d In defauts
                        Using ins = cn.CreateCommand()
                            ins.CommandText = "INSERT INTO Matiere (Nom, Coefficient) VALUES ($nom, $coef)"
                            ins.Parameters.AddWithValue("$nom", d.Nom)
                            ins.Parameters.AddWithValue("$coef", d.Coef)
                            ins.ExecuteNonQuery()
                        End Using
                    Next
                End If
            End Using
        End Using
    End Sub

    ' ---------- Utilitaires pour les notes facultatives (NULL) ----------

    ''' <summary>Valeur à passer en paramètre SQL (NULL si pas de note).</summary>
    Public Function VersDb(valeur As Double?) As Object
        If valeur.HasValue Then Return valeur.Value
        Return DBNull.Value
    End Function

    ''' <summary>Lit une colonne REAL pouvant être NULL.</summary>
    Public Function LireDouble(lecteur As SqliteDataReader, index As Integer) As Double?
        If lecteur.IsDBNull(index) Then Return Nothing
        Return lecteur.GetDouble(index)
    End Function

End Module
