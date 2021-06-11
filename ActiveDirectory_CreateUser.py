"""
AUTEUR : FARUK DEMIRCI
DATE : 30/05/2021
TACHE : CE SCRIPT A TOUT D'ABORD POUR BUT DE CREER PLUSIEURS COMPTE D'UTILISATEUR VIA UN FICHIER .CSV DANS UN ANNUAIRE ACTIVE DIRECTORY.
        DEUXIEMEMENT CE SCRIPT PERMET DE CREER UN FICHIER PERSONNEL POUR CHAQUE UTILISATEUR.
        DERNIEREMENT CE SCRIPT DONNE LES DROITS NECESSAIRE POUR CHAQUE FICHIER UTILISATEUR.
"""

#IMPORTATION DES MODULES.
from pyad import *
import csv
import os
import win32security
import ntsecuritycon as con


#CONNECTION AU COMPTE ADMINISTRATEUR DU DOMAINE
pyad.set_defaults(ldap_server="Winserver2016.mondomaine.local", username="Administrateur", password="Tunahan78")

ou = pyad.adcontainer.ADContainer.from_dn("ou=Paris, dc=mondomaine, dc=local")

#LECTURE DU FICHIER CSV + DECLARATION DES VARIABLES.
with open('nouvelutilisateur.csv', 'r', encoding='utf8') as csv_file:
    csv_reader = csv.DictReader(csv_file)
    for row in csv_reader:
        employee_id = row["employee_id"]
        nom = row["nom"]
        prenom = row["prenom"]
        description = row["description"]

        #CREATION DES UTILISATEURS.
        try:
            pyad.aduser.ADUser.create(nom + " " + prenom, ou, password="Motdepasse2021", upn_suffix="@mondomaine.local",
            optional_attributes={
            'userPrincipalName': employee_id + '@mondomaine.local', 'givenName': prenom,
            'sn': nom, 'description': description, 'displayName': nom + " " + prenom,
            'samaccountname': employee_id, 'homeDirectory': "\\\\DATA\Personnel" + "\\" + employee_id,
            'homeDrive': "H:"})

            print("L'utilisateur " + nom + " " + prenom + " a bien été créé.")

        except:
            print("L'utilisateur " + nom + " " + prenom + " n'a pas été créé.")

        #CREATION DU FICHIER PERSONNEL POUR CHAQUE UTILISATEURS.
        try:
            directory = employee_id
            parent_dir = "C:\DATA\Personnel"
            path = os.path.join(parent_dir, directory)
            os.mkdir(path)
            print("Le dossier pour l'utilisateur " + nom + " " + prenom + " a bien été créé.")


        except:
            print("Le dossier pour l'utilisateur " + nom + " " + prenom + " n'a pas été créé.")
        #ATTRIBUTION DES DROITS D'UTILISATEUR.
        try:
                        FILENAME = (r"C:\\DATA\\Personnel\\") + employee_id

                        usery, domain, type = win32security.LookupAccountName("", employee_id + "@mondomaine.local")

                        sd = win32security.GetFileSecurity(FILENAME, win32security.DACL_SECURITY_INFORMATION)
                        dacl = sd.GetSecurityDescriptorDacl()

                        dacl.AddAccessAllowedAce(win32security.ACL_REVISION, con.FILE_ALL_ACCESS, usery)

                        sd.SetSecurityDescriptorDacl(1, dacl, 0)
                        win32security.SetFileSecurity(FILENAME, win32security.DACL_SECURITY_INFORMATION, sd)
                        print("L'utilisateur a le contrôle totale sur son répertoire personnel.")
                        print("")

        except:
                        print("L'utilisateur n'a pas le contrôle totale sur son répertoire personnel.")
                        print("")