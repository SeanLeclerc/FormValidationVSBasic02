Public Class Form1
    Dim index As String
    Dim mitAjour = False
    Public enregistrement As String
    Public Structure Membres
        Dim NoMembre As String
        Dim TypeMembre As String
        Dim LangCode As String
        Dim MembreNom As String
        Dim MembrePrenom As String
        Dim MembreAdresse As String
        Dim MembreVille As String
        Dim ProvCode As String
        Dim MembreCodePostal As String
        Dim MembreNoTel As String
        Dim MembreEMail As String
    End Structure
    Public Structure Provinces
        Dim ProvCode As String
        Dim ProvDesc As String
    End Structure
    Public Structure Langues
        Dim LangCode As String
        Dim LangDesc As String
    End Structure
    Public Structure TypesMembres
        Dim TypeMembre As String
        Dim TypeMembreDesc As String
    End Structure

    Public FicheMembre As New SortedList(Of String, Membres)
    Public FicheProvince As New SortedList(Of String, Provinces)
    Public FicheLangue As New SortedList(Of String, Langues)
    Public FicheTypeMembre As New SortedList(Of String, TypesMembres)

    Public RecMembre As New Membres
    Public RecProvince As New Provinces
    Public RecTypeMembre As New TypesMembres
    Public RecLangue As New Langues

    Private Sub Button8_Click_1(sender As System.Object, e As System.EventArgs) Handles Button8.Click
        MsgBox("Onglet membre: Mettre à jour la fiche du membre affiché dans la collection.")
        ' Ajouter ici le code pour mettre à jour le membre dans la collection
        ' N'oublier pas de marquer la collection comme modifiée 
        '             afin de la sauvegarder lors de la fermeture de l'application.

        mitAjour = True
        Modifier()

    End Sub
    Private Sub AjouterDansFichierSequentiel()

        CreerChaineDelimitee()
        Dim fileExists As Boolean
        fileExists = My.Computer.FileSystem.FileExists("..\..\..\BB_Membres.txt")
        MsgBox("Fichier non existant. Il sera créé.")
            My.Computer.FileSystem.WriteAllText("..\..\..\BB_Membres.txt",
                                                                    enregistrement & vbCrLf, False)



    End Sub

    Public Sub CreerChaineDelimitee()
        For Each element In FicheMembre.Values
            enregistrement += element.NoMembre & "|" &
                element.TypeMembre & "|" &
                element.LangCode & "|" &
                element.MembreNom & "|" &
                element.MembrePrenom & "|" &
                element.MembreAdresse & "|" &
                element.MembreVille & "|" &
                element.ProvCode & "|" &
                element.MembreCodePostal & "|" &
                element.MembreNoTel & "|" &
                element.MembreEMail & vbCrLf
        Next

    End Sub

    Private Sub Button2_Click_1(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        MsgBox("Onglet membre: Afficher la première fiche de la collection des membres.")

        For Each element In FicheMembre.Values
            If element.NoMembre = "M000001" Then
                txtMembre.Text = element.NoMembre
                txtNom.Text = element.MembreNom
                txtPrenom.Text = element.MembrePrenom
                txtAdresse.Text = element.MembreAdresse
                txtVille.Text = element.MembreVille
                comboProvince.Text = element.ProvCode
                txtCodePostal.Text = element.MembreCodePostal
                txtTel.Text = element.MembreNoTel
                txtCourriel.Text = element.MembreEMail
                comboLangue.Text = element.LangCode
                comboType.Text = element.TypeMembre
            End If
        Next
        AfficherLangue()
        AfficherProvince()
        AfficherMembre()

        'Ajouter ici le code pour afficher le premier membre de la collection
        '(faites attention, vous êtes peut-être dèjà sur le premier membre)
        'N'oubliez pas de rafraîchir le nom de la province, la description de la langue et du type de membre

    End Sub

    Private Sub Button3_Click_1(sender As System.Object, e As System.EventArgs) Handles Button3.Click


        If txtMembre.Text = "M000001" = False Then
            MsgBox("Onglet membre: Afficher la fiche précédente de la collection des membres.")
            ConstruitCleRetour()
            Prochaine()
            AfficherProvince()
            AfficherLangue()
            AfficherMembre()
        Else
            MsgBox("Vous ne pouvez pas.")
        End If
        'Ajouter ici le code pour afficher le  membre précédent
        '(faites attention, vous êtes peut-être dèjà sur le premier membre)
        'N'oubliez pas de rafraîchir le nom de la province, la description de la langue et du type de membre

    End Sub


    Private Sub Button5_Click_1(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        MsgBox("Onglet membre: Afficher la dernière fiche de la collection des membres.")

        For Each element In FicheMembre.Values
            If element.NoMembre = "M000104" Then
                txtMembre.Text = element.NoMembre
                txtNom.Text = element.MembreNom
                txtPrenom.Text = element.MembrePrenom
                txtAdresse.Text = element.MembreAdresse
                txtVille.Text = element.MembreVille
                comboProvince.Text = element.ProvCode
                txtCodePostal.Text = element.MembreCodePostal
                txtTel.Text = element.MembreNoTel
                txtCourriel.Text = element.MembreEMail
                comboLangue.Text = element.LangCode
                comboType.Text = element.TypeMembre
            End If
        Next
        AfficherProvince()
        AfficherLangue()
        'Ajouter ici le code pour afficher le dernier membre 
        '(faites attention, vous êtes peut-être sur le dernier membre)
        'N'oubliez pas de rafraîchir le nom de la province, la description de la langue et du type de membre

    End Sub

    Private Sub ChargerCollections_Click(sender As System.Object, e As System.EventArgs) Handles ChargerCollections.Click
        MsgBox("Chargement des données")

        'Ajouter ici le code nécessaire pour charger les donnéers des 4 fichiers de données diponibles 
        '               dans les collections respectives.
        'Vous devrez également dans cette étape charger les collections ITEMS des 3 COMBOXBOX du formulaires 
        '               à partir des données des collections qui viennet d'être chargées.
        'Et finalement vous devez afficher l'information du premier membre.
        'N'oubliez pas de rafraîchir le nom de la province, la description de la langue et du type de membre
        loadMembres()
        loadLangue()
        loadProvince()
        TypeMembre()

        For Each element In FicheMembre.Values
            If element.NoMembre = "M000001" Then
                txtMembre.Text = element.NoMembre
                txtNom.Text = element.MembreNom
                txtPrenom.Text = element.MembrePrenom
                txtAdresse.Text = element.MembreAdresse
                txtVille.Text = element.MembreVille
                comboProvince.Text = element.ProvCode
                txtCodePostal.Text = element.MembreCodePostal
                txtTel.Text = element.MembreNoTel
                txtCourriel.Text = element.MembreEMail
                comboLangue.Text = element.LangCode
                comboType.Text = element.TypeMembre
            End If
        Next

        'Afficher les Desc
        AfficherLangue()
        AfficherProvince()
        AfficherMembre()

    End Sub
    Public Sub TypeMembre()
        'Type de membre
        Dim lineNo As Integer = 0
        Dim fileReader As System.IO.StreamReader
        Dim stringReader As String
        Dim fileExists As Boolean
        fileExists = My.Computer.FileSystem.FileExists("..\..\..\BB_TypesMembres.txt")
        If fileExists Then
            fileReader = My.Computer.FileSystem.OpenTextFileReader("..\..\..\BB_TypesMembres.txt")
            While Not fileReader.EndOfStream
                stringReader = fileReader.ReadLine()
                RemettreFichierTypeMembre(stringReader)
                lineNo += 1
            End While
            fileReader.Close()
            Me.comboType.Items.Clear()
            For Each element In FicheTypeMembre.Keys
                Me.comboType.Items.Add(element)
            Next element
        Else
            MsgBox("Le fichier n'existe pas !")
        End If
    End Sub
    Public Sub loadProvince()
        'Province
        Dim lineNo As Integer = 0
        Dim fileReader As System.IO.StreamReader
        Dim stringReader As String
        Dim fileExists As Boolean
        fileExists = My.Computer.FileSystem.FileExists("..\..\..\BB_Provinces.txt")
        If fileExists Then
            fileReader = My.Computer.FileSystem.OpenTextFileReader("..\..\..\BB_Provinces.txt")
            While Not fileReader.EndOfStream
                stringReader = fileReader.ReadLine()
                RemettreFichierProvince(stringReader)
                lineNo += 1
            End While
            fileReader.Close()
            Me.comboProvince.Items.Clear()
            For Each element In FicheProvince.Keys
                Me.comboProvince.Items.Add(element)
            Next element
        Else
            MsgBox("Le fichier n'existe pas !")
        End If
    End Sub

    Public Sub loadLangue()
        'Langue
        Dim lineNo As Integer = 0
        Dim fileReader As System.IO.StreamReader
        Dim stringReader As String
        Dim fileExists As Boolean
        fileExists = My.Computer.FileSystem.FileExists("..\..\..\BB_Langues.txt")
        If fileExists Then
            fileReader = My.Computer.FileSystem.OpenTextFileReader("..\..\..\BB_Langues.txt")
            While Not fileReader.EndOfStream
                stringReader = fileReader.ReadLine()
                RemettreFichierLangue(stringReader)
                lineNo += 1
            End While
            fileReader.Close()
            Me.comboLangue.Items.Clear()
            For Each element In FicheLangue.Keys
                Me.comboLangue.Items.Add(element)
            Next element
        Else
            MsgBox("Le fichier n'existe pas !")
        End If
    End Sub
    Public Sub loadMembres()

        Dim lineNo As Integer = 0
        Dim fileReader As System.IO.StreamReader
        Dim stringReader As String
        Dim fileExists As Boolean
        fileExists = My.Computer.FileSystem.FileExists("..\..\..\BB_Membres.txt")
        If fileExists Then
            fileReader = My.Computer.FileSystem.OpenTextFileReader("..\..\..\BB_Membres.txt")
            While Not fileReader.EndOfStream
                stringReader = fileReader.ReadLine()
                RemettreFichierDansComboBox(stringReader)
                lineNo += 1
            End While
            fileReader.Close()
            'Me.comboLangue.Items.Clear()
            'For Each element In FicheMembre.Keys
            '    Me.comboLangue.Items.Add(element)
            'Next element
        Else
            MsgBox("Le fichier n'existe pas !")
        End If
    End Sub


    Public Sub RemettreFichierDansComboBox(ByRef stringReader As String)
        Dim file As New ArrayList
        Dim key As String
        If CountCharacter(stringReader, "|") = 10 Then
            file = LireLigne(stringReader, 11)
            key = file(0)
            RecMembre.NoMembre = file(0)
            RecMembre.TypeMembre = file(1)
            RecMembre.LangCode = file(2)
            RecMembre.MembreNom = file(3)
            RecMembre.MembrePrenom = file(4)
            RecMembre.MembreAdresse = file(5)
            RecMembre.MembreVille = file(6)
            RecMembre.ProvCode = file(7)
            RecMembre.MembreCodePostal = file(8)
            RecMembre.MembreNoTel = file(9)
            RecMembre.MembreEMail = file(10)




            If FicheMembre.ContainsKey(key) Then
                MsgBox("ERREUR : La clé ==> " & CStr(key) & " <== existe déjà.")
            Else
                FicheMembre.Add(key, RecMembre)
            End If
        Else
            MsgBox("ERREUR : Une entrée invalide à été ignorée.")
        End If


    End Sub

    Public Function CountCharacter(ByVal value As String, ByVal chaine As Char) As Integer
        Dim count As Integer = 0
        For Each caract As Char In value
            If caract = chaine Then
                count += 1
            End If
        Next
        Return count
    End Function

    Public Function LireLigne(ByRef ligne As String, nbChamps As Integer) As ArrayList
        Dim affiche As New ArrayList
        Dim indice As Integer = 1
        Dim count As Integer = 0
        While count <> nbChamps
            indice = InStr(ligne, "|")
            If indice <> 0 Then
                affiche.Add(ligne.Substring(0, indice - 1))
                ligne = ligne.Remove(0, indice)
            Else
                affiche.Add(ligne.Substring(0))
            End If
            count += 1
        End While
        Return affiche
    End Function
    '-------------------------------------------------------------
    Public Sub RemettreFichierLangue(ByRef stringReader As String)
        Dim file As New ArrayList
        Dim Langue As Langues
        Dim key As String
        If CountCharacter(stringReader, "|") = 1 Then
            file = LireLigne(stringReader, 3)
            key = file(0)
            Langue.LangCode = file(0)
            Langue.LangDesc = file(1)

            If FicheLangue.ContainsKey(key) Then
                MsgBox("ERREUR : La clé ==> " & CStr(key) & " <== existe déjà.")
            Else
                FicheLangue.Add(key, Langue)
            End If
        Else
            MsgBox("ERREUR : Une entrée invalide à été ignorée.")
        End If

    End Sub
    '--------------------------------------------------------

    Public Sub RemettreFichierProvince(ByRef stringReader As String)
        Dim file As New ArrayList
        Dim Province As Provinces
        Dim key As String
        If CountCharacter(stringReader, "|") = 1 Then
            file = LireLigne(stringReader, 3)
            key = file(0)
            Province.ProvCode = file(0)
            Province.ProvDesc = file(1)

            If FicheProvince.ContainsKey(key) Then
                MsgBox("ERREUR : La clé ==> " & CStr(key) & " <== existe déjà.")
            Else
                FicheProvince.Add(key, Province)
            End If
        Else
            MsgBox("ERREUR : Une entrée invalide à été ignorée.")
        End If

    End Sub
    '--------------------------------------------------------
    '--------------------------------------------------------

    Public Sub RemettreFichierTypeMembre(ByRef stringReader As String)
        Dim file As New ArrayList
        Dim Type As TypesMembres
        Dim key As String
        If CountCharacter(stringReader, "|") = 1 Then
            file = LireLigne(stringReader, 3)
            key = file(0)
            Type.TypeMembre = file(0)
            Type.TypeMembreDesc = file(1)

            If FicheTypeMembre.ContainsKey(key) Then
                MsgBox("ERREUR : La clé ==> " & CStr(key) & " <== existe déjà.")
            Else
                FicheTypeMembre.Add(key, Type)
            End If
        Else
            MsgBox("ERREUR : Une entrée invalide à été ignorée.")
        End If

    End Sub
    '--------------------------------------------------------
    'Affiche la desc de la langue
    Public Sub AfficherLangue()
        If comboLangue.Text = "AL" Then
            txtLangue.Text = "Allemend"
        End If
        If comboLangue.Text = "CH" Then
            txtLangue.Text = "Chinois"
        End If
        If comboLangue.Text = "EN" Then
            txtLangue.Text = "Anglais"
        End If
        If comboLangue.Text = "FR" Then
            txtLangue.Text = "Francais"
        End If
        If comboLangue.Text = "IT" Then
            txtLangue.Text = "Italien"
        End If
        If comboLangue.Text = "PO" Then
            txtLangue.Text = "Portugais"
        End If
        If comboLangue.Text = "SP" Then
            txtLangue.Text = "Espagnol"
        End If

    End Sub
    '---------------------------------------------------------
    Public Sub AfficherProvince()
        If comboProvince.Text = "AB" Then
            txtProvince.Text = "Alberta"
        End If
        If comboProvince.Text = "BC" Then
            txtProvince.Text = "Colombie-Britannique"
        End If
        If comboProvince.Text = "MB" Then
            txtProvince.Text = "Manitoba"
        End If
        If comboProvince.Text = "NB" Then
            txtProvince.Text = "Nouveau-Brunswick"
        End If
        If comboProvince.Text = "NL" Then
            txtProvince.Text = "Terre-Neuve-et-Labrador"
        End If
        If comboProvince.Text = "NS" Then
            txtProvince.Text = "Nouvelle-Écosse"
        End If
        If comboProvince.Text = "NT" Then
            txtProvince.Text = "Territoires-du-Nord-Ouest"
        End If
        If comboProvince.Text = "NU" Then
            txtProvince.Text = "Nunavut"
        End If
        If comboProvince.Text = "ON" Then
            txtProvince.Text = "Ontario"
        End If
        If comboProvince.Text = "PE" Then
            txtProvince.Text = "Ile-du-Prince-Edouard"
        End If
        If comboProvince.Text = "QC" Then
            txtProvince.Text = "Québec"
        End If
        If comboProvince.Text = "SK" Then
            txtProvince.Text = "Saskatchewan"
        End If
        If comboProvince.Text = "YT" Then
            txtProvince.Text = "Yukon"
        End If
    End Sub
    Public Sub AfficherMembre()
        If comboType.Text = "MADO" Then
            txtType.Text = "Adolescents"
        End If
        If comboType.Text = "MADU" Then
            txtType.Text = "Adultes"
        End If
        If comboType.Text = "MAIN" Then
            txtType.Text = "Aines"
        End If
        If comboType.Text = "MENF" Then
            txtType.Text = "Enfants"
        End If
        If comboType.Text = "METU" Then
            txtType.Text = "Etudiant"
        End If
    End Sub



    Private Sub Sauvegarder_Click(sender As System.Object, e As System.EventArgs) Handles Sauvegarder.Click
        MsgBox("Sauvegarde de la collection des Membres")

        CreerChaineDelimitee()
        AjouterDansFichierSequentiel()
        'Ajouter ici le code nécessaire pour sauvegarder (Écrire par dessus le fichier existant)
        '               les donnéers de la collection des Membres, si celle-ci 
        '               a subi des changements (Mise-À-Jour).
        'Vous ne pouvez pas fermer l'application sans faire cette sauvegarde.

    End Sub

    Private Sub FermerApplication_Click(sender As System.Object, e As System.EventArgs) Handles FermerApplication.Click
        MsgBox("Fermeture de l'application")
        If mitAjour = True Then
            Close()
        End If

        'Vous ne pouvez pas fermer l'application si la collection des membres a été modifiée
        '        et que la collection n'a pas été sauvegardée.

    End Sub

    Public Sub Modifier()
        Dim cle As New Membres
        cle = CType(FicheMembre(index), Membres)
        cle.NoMembre = txtMembre.Text
        cle.TypeMembre = txtType.Text
        cle.LangCode = txtLangue.Text
        cle.MembreNom = txtNom.Text
        cle.MembrePrenom = txtPrenom.Text
        cle.MembreAdresse = txtAdresse.Text
        cle.MembreVille = txtVille.Text
        cle.ProvCode = comboProvince.Text
        cle.MembreCodePostal = txtCodePostal.Text
        cle.MembreNoTel = txtTel.Text
        cle.MembreEMail = txtCourriel.Text
        FicheMembre.Remove(index)
        FicheMembre.Add(index, cle)
    End Sub
    Public Sub Prochaine()
        Dim cle As New Membres
        cle = CType(FicheMembre(index), Membres)
        txtMembre.Text = cle.NoMembre
        txtType.Text = cle.TypeMembre
        txtLangue.Text = cle.LangCode
        txtNom.Text = cle.MembreNom
        txtPrenom.Text = cle.MembrePrenom
        txtAdresse.Text = cle.MembreAdresse
        txtVille.Text = cle.MembreVille
        comboProvince.Text = cle.ProvCode
        txtCodePostal.Text = cle.MembreCodePostal
        txtTel.Text = cle.MembreNoTel
        txtCourriel.Text = cle.MembreEMail
        comboLangue.Text = cle.LangCode
        comboType.Text = cle.TypeMembre
    End Sub
    Public Sub ConstruitCle()
        Dim og = CInt((txtMembre.Text).Substring(1, 6))
        og += 1
        If og < 10 Then
            index = ("M00000" + CStr(og))
        End If

        If og >= 10 And og < 100 Then
            index = ("M0000" + CStr(og))
        End If

        If og >= 100 Then
            index = ("M000" + CStr(og))
        End If

    End Sub

    Public Sub ConstruitCleRetour()
        Dim og = CInt((txtMembre.Text).Substring(1, 6))
        og -= 1
        If og < 10 Then
            index = ("M00000" + CStr(og))
        End If

        If og >= 10 And og < 100 Then
            index = ("M0000" + CStr(og))
        End If

        If og >= 100 Then
            index = ("M000" + CStr(og))
        End If

    End Sub
    Private Sub Button4_Click_1(sender As System.Object, e As System.EventArgs) Handles Button4.Click

        'Ajouter ici le code pour afficher le  membre suivant
        '(faites attention, vous êtes peut-être sur le dernier membre)
        'N'oubliez pas de rafraîchir le nom de la province, la description de la langue et du type de membre
        'If currentContact < FicheMembre.Count - 1 Then
        '    currentContact += 1
        'End If
        'Prochaine()
        If txtMembre.Text = "M000104" = False Then
            MsgBox("Onglet membre: Afficher la fiche suivante de la collection des membres.")
            ConstruitCle()
            Prochaine()
            AfficherMembre()
            AfficherProvince()
            AfficherLangue()
        Else
            MsgBox("Vous ne pouvez pas.")
        End If

    End Sub
End Class
