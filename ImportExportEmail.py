"""Module gestion contact"""

import codecs
import os
from os import path as os_path
import sys
import re
from tkinter.filedialog import askopenfilename, asksaveasfilename
import subprocess

GROUPE_DEFAULT_NAME = "* My Contacts"

NAME_COLONNE_NAME = "Name"
NAME_COLONNE_GIVEN_NAME = "Given Name"
NAME_COLONNE_FAMILY_NAME = "Family Name"
NAME_COLONNE_GROUP = "Group Membership"
NAME_COLONNE_EMAIL1_TYPE = "E-mail 1 - Type"
NAME_COLONNE_EMAIL1_VALUE = "E-mail 1 - Value"
NAME_COLONNE_ORGANIZATION_NAME = "Organization 1 - Name"
NAME_COLONNE_ADRESS_FORMATTED = "Address 1 - Formatted"
NAME_COLONNE_ADRESS_STREET = "Address 1 - Street"
NB_COL_MIN_GOOGLE = 27

NAME_COLUMN_NOM_CSV = "Nom"
NAME_COLUMN_PRENOM_CSV = "Prénom"
NAME_COLUMN_ADRESSE_MAIL_CSV = "Adresse mail"
NAME_COLUMN_GROUPE_CSV = "Groupe"
NAME_COLUMN_TELEPHONE_CSV = "Téléphone"
NAME_COLUMN_ENTREPRISE_CSV = "Entreprise"
NAME_COLUMN_ADRESSE_CSV = "Adresse postale"
NAME_COLUMN_CONTACT_CSV = "Contact"

TYPE_FICHIER_GOOGLE = "GOOGLE"
TYPE_FICHIER_SAMSUNG = "SAMSUNG"
TYPE_FICHIER_CSV = "CSV"
LISTE_TYPE_FICHIER = [TYPE_FICHIER_GOOGLE, TYPE_FICHIER_SAMSUNG, TYPE_FICHIER_CSV]

FILE_SAVE_LOCAL = "ListingGestionEmail.csv"

L_COL_EXP_GOOGLE = [NAME_COLONNE_NAME
                    , NAME_COLONNE_GIVEN_NAME
                    , NAME_COLONNE_FAMILY_NAME
                    , NAME_COLONNE_GROUP]
L_CONTACT = []

TYPE_ARG_INPUT_FILE = 'INPUT_FILE'
L_ARG_INPUT_FILE = ['-i', '--InputFile']
TYPE_ARG_INPUT_TYPE = 'INPUT_TYPE'
L_ARG_INPUT_TYPE = ['-it', '--InputType']
TYPE_ARG_OUTPUT_FILE = 'OUTPUT_FILE'
L_ARG_OUTPUT_FILE = ['-o', '--OutputFile']
TYPE_ARG_OUTPUT_TYPE = 'OUTPUT_TYPE'
L_ARG_OUTPUT_TYPE = ['-ot', '--OutputType']
TYPE_ARG_HELP = 'HELP'
L_ARG_HELP = ['-h', '--help']
MAP_ARG_APPLI = {TYPE_ARG_INPUT_FILE: L_ARG_INPUT_FILE, TYPE_ARG_INPUT_TYPE: L_ARG_INPUT_TYPE,
TYPE_ARG_OUTPUT_FILE: L_ARG_OUTPUT_FILE, TYPE_ARG_OUTPUT_TYPE: L_ARG_OUTPUT_TYPE,
TYPE_ARG_HELP: L_ARG_HELP}

def menu_modification_groupe(contact):
    """Modification du(des) groupe(s) du contact"""
    num_affiche = 1
    clef_menu = {}
    for groupe in contact.groupe:
        print(str(num_affiche)+' : '+groupe)
        clef_menu[str(num_affiche)] = groupe
        num_affiche += 1
        if num_affiche % 10 == 0:
            print("")
            print(" 0) Retour modification.")
            print(" a) Ajouter un groupe.")
            commande = input("Saisir la commande... "+
                             "(Entrée pour passer au suivant).\n")
        if commande != "":
            break

def modifier_contact(nom):
    """Modification des informations du contact"""
    cls()
    print(" :: Modification contact :: ")
    for contact in L_CONTACT:
        if contact.nom+contact.prenom == nom:
            print("  NOM : "+contact.nom)
            print("  PRENOM : "+contact.prenom)
            print("  GROUPE : "+" ::: ".join(contact.groupe))
            print("  EMAIL  : "+ ", ".join(contact.email))
            print("  SOCIETE :  " + contact.entreprise)
            contact_selectionne = contact
            break
    commande = input("Modifier le contact ? (O)ui / (N)on\n")
    if commande.upper() == 'N':
        return
    valeur_saisie = input("Saisir le nouveau Nom : ("+contact_selectionne.nom+")\n")
    if valeur_saisie != "":
        contact_selectionne.nom = valeur_saisie
    valeur_saisie = input("Saisir le nouveau Prénom : ("+contact_selectionne.prenom+")\n")
    if valeur_saisie != "":
        contact_selectionne.prenom = valeur_saisie
    valeur_saisie = input("Saisir le nouveau E-mail : ("+", ".join(contact_selectionne.email)+")\n")
    if valeur_saisie != "":
        while valeur_saisie != "":
            contact_selectionne.email.append(valeur_saisie)
            valeur_saisie = input("Saisir le nouveau E-mail : (" +
                                  ", ".join(contact_selectionne.email) +
                                  "). Entrée pour finir.\n")
    valeur_saisie = input("Groupe du contact : "+",".join(contact_selectionne.groupe)+
                          "\nVoulez-vous le(s) modifier ? (O)ui / (N)on\n")
    if valeur_saisie.upper() == "O":
        menu_modification_groupe(contact_selectionne)

def menu_traiter_selection_nom(filtre_nom):
    """Traitement de la sélection d'un nom pour modification"""
    cls()
    commande = ""
    num_affiche = 1
    print(" :: Sélection contact à modifier ::")
    print(" Filtre nom : "+filtre_nom)
    clef_menu = {}
    for contact in L_CONTACT:
        if -1 != contact.nom.find(filtre_nom):
            print(str(num_affiche)+' : '+contact.nom+' '+contact.prenom)
            clef_menu[str(num_affiche)] = contact.nom+contact.prenom
            num_affiche += 1
            if num_affiche % 10 == 0:
                print("")
                print(" 0) Retour modification.")
                commande = input("Saisir la commande... "+
                                 "(Entrée pour passer au suivant).\n")
        if commande != "":
            break
    if commande == "":
        print("")
        print(" 0) Retour modification.")
        commande = input("Saisir la commande...\n")
    if commande.isalnum():
        if commande != "0":
            modifier_contact(clef_menu[commande])
            pause_menu()
    else:
        print("Aucun nom sélectionné.")
        pause_menu()

def menu_modifier_contact():
    """Modification d'un contact"""
    fin_modif_contact = 0
    filtre_nom = ""
    while fin_modif_contact == 0:
        cls()
        print(" :: Modification contact ::")
        print(" Filtre nom : "+filtre_nom)
        print("Liste commande :")
        print(" 1) Recherche par nom.")
        print(" 2) Selectionner contact.")
        print("")
        print(" 0) Retour menu principal.")
        commande = input("Saisir la commande...\n")

        if commande.isdigit():
            if commande == "0":
                fin_modif_contact = 1
            elif commande == "1":
                filtre_nom = input("Saisir le filtre nom :\n")
            elif commande == '2':
                menu_traiter_selection_nom(filtre_nom)
        else:
            print("Commande saisie incorrecte.")

def pause_menu():
    """Affichage pour faire une pause"""
    input("Taper Entrée pour continuer...")

def cls():
    """Clear screen"""
    os.system('cls')

def print_contact():
    """Affichage des contacts en mémoire"""
    strContact = ""
    for i in range(0,len(L_CONTACT)):
        strContact = L_CONTACT[i].export_contact().replace(r'\n', '\r\n')
        if "" == strContact:
            cls()
            i = 0
            strContact = L_CONTACT[i].export_contact().replace(r'\n', '\r\n')
        print(strContact)
    pause_menu()

def extract_info_google(file_name, google_colonnes):
    """Extraction des contacts à partir de l'export GOOGLE"""
    info_ligne = []
    num_ligne = 1
    if os.path.isfile(file_name):
        with codecs.open(file_name, "r", "utf16") as fichier:
            lignes = fichier.readlines()
            for ligne in lignes:
                if '\r' in ligne:
                    ligne = ligne.replace('\r', '')
                if '\n' in ligne:
                    ligne = ligne.replace('\n', '')
                if num_ligne == 1:
                    del google_colonnes[:]
                    for nom_colonne in ligne.split(","):
                        google_colonnes.append(nom_colonne)
                else:
                    info_ligne = ligne.split(",")
                    if len(info_ligne) > NB_COL_MIN_GOOGLE:
                        groupes = info_ligne[google_colonnes.index(NAME_COLONNE_GROUP)]
                        liste_groupe = []
                        while -1 != groupes.find(":"):
                            index = groupes.find(":")
                            groupe_trouve = groupes[:index-1]
                            liste_groupe.append(groupe_trouve)
                            groupes = groupes[index+4:]
                        # Récupération du dernier groupe de la liste
                        liste_groupe.append(groupes)
                        # Récupération des emails
                        liste_email = extract_email(google_colonnes, info_ligne)
                        # Récupération des telephones
                        liste_tel = extract_tel(google_colonnes, info_ligne)
                        index_col_family = google_colonnes.index(NAME_COLONNE_FAMILY_NAME)
                        index_col_given_name = google_colonnes.index(NAME_COLONNE_GIVEN_NAME)
                        infos_contact = [info_ligne[index_col_family], info_ligne[index_col_given_name]]
                        L_CONTACT.append(Contact(infos=infos_contact,
                                                email=liste_email,
                                                groupe=liste_groupe,
                                                num_tel=liste_tel))
                num_ligne += 1
        fichier.close()
    else:
        print("Erreur fichier <"+file_name+"> non présent.")
        pause_menu()

def extract_email(google_colonnes, info_ligne):
    """ Extraction des valeurs d'email dans les listes """
    regexp = re.compile(r'E-mail \d* - Value')
    num_col = 0
    liste_email = []
    for nom_colonne in google_colonnes:
        if num_col > len(info_ligne):
            break
        if regexp.search(nom_colonne) and info_ligne[num_col] != '':
            liste_email.append(info_ligne[num_col])
        num_col += 1
    return liste_email

def extract_tel(google_colonnes, info_ligne):
    """ Extraction des valeurs telephones dans les listes """
    regexp = re.compile(r'Phone \d* - Value')
    num_col = 0
    liste_tel = []
    for nom_colonne in google_colonnes:
        if num_col > len(info_ligne):
            break
        if regexp.search(nom_colonne) and info_ligne[num_col] != '':
            liste_tel.append(info_ligne[num_col])
        num_col += 1
    return liste_tel

def check_colonne_minimale_csv(liste_colonne_csv):
    return ((str.upper(NAME_COLUMN_GROUPE_CSV) in liste_colonne_csv) and (str.upper(NAME_COLUMN_ADRESSE_MAIL_CSV) in liste_colonne_csv)
    and (str.upper(NAME_COLUMN_NOM_CSV) in liste_colonne_csv))

def extract_info_csv(file_name):
    """Extraction des infos dans un fichier formater en CSV.
    Le séparateur utiliser est ';' ou ','.
    Colonnes obligatoires : "Nom", "Groupe" et "Adresse mail"
    Colonnes traitées facultatives : "Prénom", Téléphone", "Entreprise", "Adresse postale", "Contact"
    Colonne Groupe : Liste nom Groupe séparé par des ',' et encadré par des '"' si plusieurs
    Colonne Adresse mail : Liste adresse mail séparées par des ',' et encadrées par des '"' si plusieurs
    """
    info_ligne = []
    b_premiere_ligne = True
    liste_colonne = []
    emails = []
    
    with codecs.open(file_name, "r", "utf16") as fichier:
        lignes = fichier.readlines()
        for ligne in lignes:
            if not b_premiere_ligne:
                if '\r' in ligne:
                    ligne = ligne.replace('\r', '')
                if '\n' in ligne:
                    ligne = ligne.replace('\n', '')
                info_ligne = ligne.split(";")
                infos_contact = ["", "", "", ""]
                liste_groupe = []
                liste_email = []
                liste_telephone = []
                for i in range(0, len(liste_colonne)):
                    if liste_colonne[i] == NAME_COLUMN_NOM_CSV.upper():
                        infos_contact[0] = info_ligne[i].strip()
                    if liste_colonne[i] == NAME_COLUMN_PRENOM_CSV.upper():
                        infos_contact[1] = info_ligne[i].strip()
                    if liste_colonne[i] == NAME_COLUMN_CONTACT_CSV.upper():
                        dyn_contact = info_ligne[i].strip().replace('"', '').split(" ")
                        lg_champ = len(dyn_contact)
                        if lg_champ > 0:
                            infos_contact[0] = info_ligne[i].strip().split(" ")[0]
                        if lg_champ > 1:
                            infos_contact[1] = info_ligne[i].strip().split(" ")[1]
                    # Récupération des groupes
                    if liste_colonne[i] == NAME_COLUMN_GROUPE_CSV.upper():
                        groupes = info_ligne[i].split(',')
                        lg_groupe = len(groupes)
                        if lg_groupe > 0:
                            for groupe in groupes:
                                liste_groupe.append(groupe)
                    # Récupération des e-mail
                    if NAME_COLUMN_ADRESSE_MAIL_CSV.upper() in liste_colonne[i]:
                        if ',' in info_ligne[i]:
                            emails = info_ligne[i].strip().split(",")
                            for email in emails:
                                liste_email.append(email)
                        else: 
                            liste_email.append(info_ligne[i])
                    # Récupération des téléphones
                    if NAME_COLUMN_TELEPHONE_CSV.upper() in liste_colonne[i]:
                        liste_telephone.append(info_ligne[i])
                    # Récupération de l'entreprise
                    if NAME_COLUMN_ENTREPRISE_CSV.upper() in liste_colonne[i]:
                        infos_contact[2] = info_ligne[i]
                    # Récupération de l'adresse
                    if NAME_COLUMN_ADRESSE_CSV.upper() in liste_colonne[i]:
                        infos_contact[3] = info_ligne[i]
                L_CONTACT.append(Contact(infos=infos_contact,
                                         email=liste_email,
                                         groupe=liste_groupe,
                                         num_tel=liste_telephone))
            else:
                if '\r' in ligne:
                    ligne = ligne.replace('\r', '')
                if '\n' in ligne:
                    ligne = ligne.replace('\n', '')
                ligne = ligne.upper()
                if ';' in ligne:
                    liste_colonne = ligne.split(';')
                elif ',' in ligne:
                    liste_colonne = ligne.split(',')
                else:
                    print("Format de la liste des colonnes invalide. Séparateur ',' ou ';'.")
                    pause_menu()
                    return
                b_premiere_ligne = False
                if not check_colonne_minimale_csv(liste_colonne):
                    print("Manque colonne dans fichier csv : " + file_name + ". Colonnes présente : "+ ",".join(liste_colonne))
                    pause_menu()
                    return
        fichier.close()

def extract_info_ligne_samsung(info_ligne, liste_colonne_samsung):
    """Extraction des infos de contact à partir d'une ligne du fichier Kies"""
    reg_email = re.compile(r'^E-mail\d*(Type)')
    reg_telephone = re.compile(r'^Numéro de téléphone\d*\(Type\)')
    reg_nom = re.compile(r'^Afficher le nom$')
    liste_email = []
    liste_tel = []
    infos_contact = []
    nb_email = 0
    nb_telephone = 0
    for i in range(0, len(liste_colonne_samsung)):
        nom_colonne = liste_colonne_samsung[i]
        if reg_email.search(nom_colonne) and info_ligne[i+1].strip().replace('"', '') != "":
            liste_email.append(info_ligne[i+1].strip().replace('"', ''))
            nb_email += 1
        if reg_telephone.search(nom_colonne) and info_ligne[i+1].strip().replace('"', '').replace(' ', '') != "":
            liste_tel.append(info_ligne[i+1].strip().replace('"', '').replace(' ', ''))
            nb_telephone += 1
        if reg_nom.search(nom_colonne):
            infos_contact = info_ligne[i].strip().replace('"', '').split(' ')
    L_CONTACT.append(Contact(infos=infos_contact,
                             num_tel=liste_tel,
                             email=liste_email))

def extract_info_samsung(file_name):
    """Extraction des infos de contact à partir d'un fichier Kies"""
    info_ligne = []
    num_ligne = 1
    liste_colonne_samsung = []
    with codecs.open(file_name, mode='r') as fichier:
        lignes = fichier.readlines()
        for ligne in lignes:
            if '\r' in ligne:
                ligne = ligne.replace('\r', '')
            if '\n' in ligne:
                ligne = ligne.replace('\n', '')
            info_ligne = ligne.split(";")
            if num_ligne == 1:
                for nom_colonne in info_ligne:
                    nom_colonne = nom_colonne.replace('"', '')
                    liste_colonne_samsung.append(nom_colonne)
                print(str(liste_colonne_samsung))
            else:
                if extract_info_ligne_samsung(info_ligne, liste_colonne_samsung) is False:
                    print("Erreur import ligne <"+num_ligne+"> fichier <"+file_name+">")
                    return
            num_ligne += 1
        fichier.close()
    pause_menu()

def extract_info_contact(file_name, google_colonnes, type_fichier='GOOGLE'):
    """Extraction des informations à partir de type fichier"""
    erreur = False
    if TYPE_FICHIER_GOOGLE == type_fichier:
        extract_info_google(file_name, google_colonnes)
    elif TYPE_FICHIER_CSV == type_fichier:
        extract_info_csv(file_name)
    elif TYPE_FICHIER_SAMSUNG == type_fichier:
        extract_info_samsung(file_name)
    else:
        print("Extraction de type de fichier <"+type_fichier+"> non géré.")
        erreur = True
        pause_menu()
    if not erreur:
        google_colonnes = calcul_col_google(google_colonnes)
    return google_colonnes

def calcul_col_google(google_colonnes):
    """ Recalcul les colonne google si pas assez pour les données des contacts """
    nb_tel_max = 0
    for contact in L_CONTACT:
        if len(contact.num_tel) > nb_tel_max:
            nb_tel_max = len(contact.num_tel)
    nb_col_tel = 1
    num_col_tel_max = 0
    reg_phone = re.compile(r'Phone \d* - Value')
    for nom_col in google_colonnes:
        if reg_phone.search(nom_col):
            nb_col_tel += 1
            num_col_tel_max = google_colonnes.index(nom_col)
    # Ajout colonne tel si nécessaire
    while nb_tel_max > nb_col_tel:
        google_colonnes.insert(num_col_tel_max+1, "Phone {0} - Type".format(nb_col_tel))
        google_colonnes.insert(num_col_tel_max+2, "Phone {0} - Value".format(nb_col_tel))
        num_col_tel_max += 2
        nb_col_tel += 1
    return google_colonnes

def sauvegarde_contact_google(file_name, forceSave):
    """Sauvegarde des contacts dans le fichier de GOOGLE"""
    deleteFile = forceSave
    if file_name == "":
        input("Veuillez saisir un nom de fichier non vide.")
        return
    if os.path.exists(file_name) and not forceSave:
        reponse = input("Le fichier <"+file_name+
                        "> existe. Voulez-vous l'écraser ? (O)ui / (N)on.\n")
        if reponse.upper() == "O":
            deleteFile = True
        else:
            return
    if deleteFile and os.path.exists(file_name):
        os.remove(file_name)

    """ Parse une première fois pour ajouter colonne manquante si necessaire """
    for contact in L_CONTACT:
        ligne = contact.export_contact()

    fichier = codecs.open(file_name, 'w', 'utf16')
    fichier.write(",".join(L_COL_EXP_GOOGLE)+'\r\n')
    for contact in L_CONTACT:
        ligne = contact.export_contact()
        fichier.write(ligne+'\r\n')
    fichier.close()
    if forceSave:
        print("Export de '"+str(len(L_CONTACT))+"' dans le fichier '"+file_name+"' réussi.")

def menu_modifier_type_fichier(type_fichier_courant):
    """Menu pour modification du type de fichier"""
    sortir = 0
    commande = "-1"
    while sortir == 0:
        cls()
        print(" :: Modification type fichier ::")
        print(" Type fichier : " + type_fichier_courant)
        print("Liste commande :")
        print(" 1) "+TYPE_FICHIER_GOOGLE+".")
        print(" 2) "+TYPE_FICHIER_SAMSUNG+".")
        print(" 3) "+TYPE_FICHIER_CSV+".")
        print("")
        print(" 0) Annuler.")
        commande = input("Saisir votre commande...\n")

        if commande == "1":
            type_fichier_courant = TYPE_FICHIER_GOOGLE
            sortir = 1
        elif commande == "2":
            type_fichier_courant = TYPE_FICHIER_SAMSUNG
            sortir = 1
        elif commande == "3":
            type_fichier_courant = TYPE_FICHIER_CSV
            sortir = 1
        elif commande == "0":
            sortir = 1
        else:
            print("Saisie ("+str(commande)+") invalide.")
            pause_menu()

    return type_fichier_courant

def sauvegarde_contact(type_fichier_courant=TYPE_FICHIER_GOOGLE, file_name=FILE_SAVE_LOCAL, forceSave=False):
    """Sauvegarde des contacts dans un fichier
    type_fichier_courant : type de fichier courant
    file_name : chemin et nom de fichier de sortie"""
    cls()
    sortir = 0
    while sortir == 0:
        cls()
        if not forceSave:
            print(" :: Sauvegarde des contacts ::")
            print(" Fichier destination : " + file_name)
            print(" Type fichier : " + type_fichier_courant)
            print("Liste commande :")
            print(" 1) Saisir nom fichier.")
            print(" 2) Sauvegarder.")
            print("")
            print(" 0) sortir du programme.")
            commande_saisie = input("Saisir votre commande...\n")
        else:
            commande_saisie = "2"

        if commande_saisie == "0":
            sortir = 1
        elif commande_saisie == "1":
            file_name = asksaveasfilename(title="Fichier destination...",
                                          filetypes=[('csv files', '.csv')])
        elif commande_saisie == "2":
            if type_fichier_courant == TYPE_FICHIER_GOOGLE:
                sauvegarde_contact_google(file_name, forceSave)
            elif type_fichier_courant == TYPE_FICHIER_SAMSUNG:
                print("TODO Sauvegarde fichier Samsung")
                pause_menu()
            elif type_fichier_courant == TYPE_FICHIER_CSV:
                print("TODO Sauvegarde fichier CSV")
                pause_menu()
            else:
                print("Type export <"+type_fichier_courant+"> non géré.")
                pause_menu()
            if forceSave:
                sortir = 1
        else:
            print("Commande <"+commande_saisie+"> non gérée.")

def traiter_liste_tel(num_tel):
    """ Traitement des numéro de téléphone """
    index = 0
    for tel in num_tel:
        if tel != "" and len(tel) > 1 and tel[0] != "+":
            if tel[0] == "0":
                tel = tel[1:]
            tel = tel.replace('.', '').replace(' ', '')
            tel = "+33" + tel
            num_tel[index] = tel
        index += 1
    return num_tel

class Contact:
    """Classe les données d'un contact"""
    def __init__(self, infos=None, email=None, groupe=None, num_tel=None):
        """Constructeur de la classe"""
        if infos is None:
            infos = ["", ""]
        if groupe is None:
            groupe = [GROUPE_DEFAULT_NAME]
        if email is None:
            email = []
        if num_tel is None:
            num_tel = []
        if 0 == len(infos):
            print("Impossible de créer le contact. Aucune information disponible.")
        elif 1 == len(infos):
            infos.append("")
        self.nom = infos[0]
        self.prenom = infos[1]
        self.email = email
        self.groupe = groupe
        self.num_tel = traiter_liste_tel(num_tel)
        if len(infos) >= 3:
            self.entreprise = infos[2]
        else:
            self.entreprise = ""
        if len(infos) >= 4:
            self.adresse = infos[3]
        else:
            self.adresse = ""

    def __str__(self):
        """Affichage un peu plus joli de nos objets"""
        return ("Nom <"+self.nom+"> Prenom <"+self.prenom+
                "> Email <"+str(self.email)+"> Group <"+" ::: ".join(self.groupe)+
                "> Telephone <"+", ".join(self.num_tel)+"> Adresse <"+self.adresse+">.")

    def export_contact_google(self):
        """Retourne un contact formaté pour GOOGLE"""
        liste_info = []
        nb_phone = len(self.num_tel)
        idx_email_google = 0
        idx_phone_contact = 0
        is_ajoutcolonne = 0
        regexp_email = re.compile(r'E-mail \d* - Value')
        regexp_phone = re.compile(r'Phone \d* - Value')
        for i in range(0, len(L_COL_EXP_GOOGLE)):
            nom_colonne = L_COL_EXP_GOOGLE[i]
            liste_info.append("")
            liste_info[i] = ""
            if regexp_email.search(nom_colonne) and idx_email_google < len(self.email):
                if len(self.email) > 1:
                    liste_info[i] = "\""
                    liste_info[i] += ">, <".join(self.email)
                    liste_info[i] += "\""
                else:
                    liste_info[i] = self.email[0]
            if regexp_phone.search(nom_colonne) and idx_phone_contact < len(self.num_tel):
                liste_info[i] = self.num_tel[idx_phone_contact]
                idx_phone_contact += 1
        if nb_phone > idx_phone_contact:
            for i in range(0, nb_phone-idx_phone_contact):
                liste_info.append(self.num_tel[idx_phone_contact+i])
                L_COL_EXP_GOOGLE.append("Phone "+str(idx_phone_contact+i)+" - Type")
                L_COL_EXP_GOOGLE.append("Phone "+str(idx_phone_contact+i)+" - Value")
                is_ajoutcolonne = 1
        if is_ajoutcolonne == 1:
            return ""
        if "" == self.prenom:
            liste_info[L_COL_EXP_GOOGLE.index(NAME_COLONNE_NAME)] = self.nom
        elif "" == self.nom:
            liste_info[L_COL_EXP_GOOGLE.index(NAME_COLONNE_NAME)] = self.prenom
        else:
            liste_info[L_COL_EXP_GOOGLE.index(NAME_COLONNE_NAME)] = self.prenom + " " + self.nom
        liste_info[L_COL_EXP_GOOGLE.index(NAME_COLONNE_FAMILY_NAME)] = self.nom
        liste_info[L_COL_EXP_GOOGLE.index(NAME_COLONNE_GIVEN_NAME)] = self.prenom
        if GROUPE_DEFAULT_NAME not in self.groupe:
            self.groupe.append(GROUPE_DEFAULT_NAME)
        liste_info[L_COL_EXP_GOOGLE.index(NAME_COLONNE_GROUP)] = " ::: ".join(self.groupe)
        if NAME_COLONNE_ORGANIZATION_NAME not in L_COL_EXP_GOOGLE:
            L_COL_EXP_GOOGLE.append(NAME_COLONNE_ORGANIZATION_NAME)
        liste_info[L_COL_EXP_GOOGLE.index(NAME_COLONNE_ORGANIZATION_NAME)] = self.entreprise
        if NAME_COLONNE_ADRESS_FORMATTED not in L_COL_EXP_GOOGLE:
            L_COL_EXP_GOOGLE.append(NAME_COLONNE_ADRESS_FORMATTED)
        liste_info[L_COL_EXP_GOOGLE.index(NAME_COLONNE_ADRESS_FORMATTED)] = self.adresse
        if NAME_COLONNE_ADRESS_STREET not in L_COL_EXP_GOOGLE:
            L_COL_EXP_GOOGLE.append(NAME_COLONNE_ADRESS_STREET)
        liste_info[L_COL_EXP_GOOGLE.index(NAME_COLONNE_ADRESS_STREET)] = self.adresse
        return ",".join(liste_info)

    def export_contact(self, type_export=TYPE_FICHIER_GOOGLE):
        """Retourne le contact pour être exporté"""
        s_result = ""
        if type_export == TYPE_FICHIER_GOOGLE:
            s_result = self.export_contact_google()
        elif type_export == TYPE_FICHIER_SAMSUNG:
            print("TODO export samsung")
        elif type_export == TYPE_FICHIER_CSV:
            print("TODO export CSV")
        else:
            print("Type export <"+str(type_export)+"> non géré.")
        return s_result

def print_usage(erreur=""):
    print("\n GestionEmail \n\n  Permet de convertir un fichier email dans un autre format pour l'import dans Google ou pour les GSM Samsung.\n\n   Options disponibles :"+
          "\n\n\t-h, --help : Affichage de ce message"+
          "\n\n\t-i, --InputFile : Chemin et nom du fichier à transformer"+
          "\n\n\t-it, --InputType : Type de fichier à transformer dans " + str(LISTE_TYPE_FICHIER) +
          "\n\n\t-o, --OutputFile : Chemin et nom du fichier générer. Par défaut " + FILE_SAVE_LOCAL +
          "\n\n\t-ot, --OutputType : Type de fichier à générer dans " + str(LISTE_TYPE_FICHIER))
    if "" != erreur:
        print("\n\n  Erreur : '"+erreur+"'")

def affiche_menu_gestion_fichier(nom_fichier_courant, google_colonnes):
    sortir = 0
    type_fichier_courant = TYPE_FICHIER_GOOGLE
    while sortir == 0:
        cls()
        print(" :: Gestion des contacts ::")
        print(" Fichier en cours : " + c)
        print(" Type fichier : " + type_fichier_courant)
        print(" Nb contact présent : " + str(len(L_CONTACT)))
        print("Liste commande :")
        print(" 1) Saisir nom fichier.")
        print(" 2) Charger fichier.")
        print(" 3) Sauvegarder fichier contact.")
        print(" 4) Modifier type fichier.")
        print("")
        print(" 0) sortir du programme.")
        commande_saisie = input("Saisir votre commande...\n")

        if commande_saisie == "0":
            sortir = 1
        elif commande_saisie == "1":
            nom_fichier_courant = askopenfilename(title="Ouvrir votre document",
                                                    filetypes=[('csv files', '.csv'),
                                                                ('txt files', '.txt'),
                                                                ('all files', '.*')])
        elif commande_saisie == "2":
            google_colonnes = extract_info_contact(nom_fichier_courant, google_colonnes, type_fichier_courant)
        elif commande_saisie == "3":
            sauvegarde_contact()
        elif commande_saisie == "4":
            type_fichier_courant = menu_modifier_type_fichier(type_fichier_courant)

def affiche_menu_gestion_contact(nom_fichier_courant):
    sortir = 0
    type_fichier_courant = TYPE_FICHIER_GOOGLE
    while sortir == 0:
        cls()
        print(" :: Gestion des contacts ::")
        print(" Fichier en cours : " + nom_fichier_courant)
        print(" Type fichier : " + type_fichier_courant)
        print(" Nb contact présent : " + str(len(L_CONTACT)))
        print("Liste commande :")
        print(" 1) Saisir un contact (TODO).")
        print(" 2) Liste contact.")
        print(" 3) Modifier contact.")
        print(" 4) Vider les contacts.")
        print("")
        print(" 0) sortir du programme.")
        commande_saisie = input("Saisir votre commande...\n")

        if commande_saisie == "0":
            sortir = 1
        elif commande_saisie == "1":
            print("TODO")
        elif commande_saisie == "2":
            print_contact()
        elif commande_saisie == "3":
            menu_modifier_contact()
        elif commande_saisie == "4":
            L_CONTACT.clear()

def affiche_menu_principal(nom_fichier_courant, google_colonnes):
    sortir = 0
    type_fichier_courant = TYPE_FICHIER_GOOGLE
    while sortir == 0:
        cls()
        print(" :: Gestion des contacts ::")
        print(" Fichier en cours : " + nom_fichier_courant)
        print(" Type fichier : " + type_fichier_courant)
        print(" Nb contact présent : " + str(len(L_CONTACT)))
        print("Liste commande :")
        print(" 1) Gestion fichier.")
        if 0 != len(L_CONTACT):
            print(" 2) Gestion contact.")
        print("")
        print(" 0) sortir du programme.")
        commande_saisie = input("Saisir votre commande...\n")

        if commande_saisie == "0":
            sortir = 1
        elif commande_saisie == "1":
            affiche_menu_gestion_fichier(nom_fichier_courant, google_colonnes)
        elif commande_saisie == "2":
            affiche_menu_gestion_contact(nom_fichier_courant)
s
def traitemente_argument(argv, google_colonnes):
    """arg_erreur = 0"""
    map_argument = {}
    for type_arg in MAP_ARG_APPLI.keys():
        for option in MAP_ARG_APPLI[type_arg]:
            if option in argv and len(argv) > argv.index(option) +1:
                print(str(argv) + " nb arg " + str(len(argv)) + " idx " + str(argv.index(option)))
                map_argument[type_arg] = argv[argv.index(option)+1]
    if (TYPE_ARG_INPUT_TYPE in map_argument.keys() and TYPE_ARG_INPUT_FILE in map_argument.keys()
    and TYPE_ARG_OUTPUT_TYPE in map_argument.keys() ) :
        fileOutputName = ""
        if map_argument[TYPE_ARG_INPUT_TYPE] not in LISTE_TYPE_FICHIER:
            print_usage("Le type de fichier à importer '" + map_argument[TYPE_ARG_INPUT_TYPE] + "' n'est pas correct.")
            exit(-1)
        if map_argument[TYPE_ARG_OUTPUT_TYPE] not in LISTE_TYPE_FICHIER:
            print_usage("Le type de fichier à générer '" + map_argument[TYPE_ARG_OUTPUT_TYPE] + "' n'est pas correct.")
            exit(-1)
        if not os.path.isfile(map_argument[TYPE_ARG_INPUT_FILE]):
            print_usage("Le fichier '" + map_argument[TYPE_ARG_INPUT_FILE] + "' n'existe pas. Import impossible.")
            exit(-2)
        if not os.path.isdir(os.path.dirname(map_argument[TYPE_ARG_OUTPUT_FILE])):
            print("Le fichier '" + map_argument[TYPE_ARG_OUTPUT_FILE] + "' n'existe pas. Utilisation du fichier par défaut : "+FILE_SAVE_LOCAL+".")
            fileOutputName = os.path.dirname(sys.argv[0]) + "/" + FILE_SAVE_LOCAL
        else:
            fileOutputName = map_argument[TYPE_ARG_OUTPUT_FILE]
        google_colonnes = extract_info_contact(map_argument[TYPE_ARG_INPUT_FILE], google_colonnes, map_argument[TYPE_ARG_INPUT_TYPE])
        sauvegarde_contact(map_argument[TYPE_ARG_OUTPUT_TYPE], fileOutputName, True)
    elif TYPE_ARG_HELP:
        print_usage()
    else:
        print_usage("Erreur argument:" + str(argv) + " argumment présent : " + str(map_argument))

def main(argv, google_colonnes):
    """Gestion des contacts pour import dans Gmail"""
    if os.path.isfile(FILE_SAVE_LOCAL):
        # Récupération des contacts sauvegardés
        google_colonnes = extract_info_contact(FILE_SAVE_LOCAL, google_colonnes, TYPE_FICHIER_GOOGLE)
        nom_fichier_courant = FILE_SAVE_LOCAL
    else:
        # Récupération des noms de colonnes google
        localPath = os_path.abspath(os_path.split(__file__)[0])
        nom_fichier_courant = localPath + "/google.csv"
        google_colonnes = extract_info_contact(nom_fichier_courant, google_colonnes, TYPE_FICHIER_GOOGLE)
    if len(argv) < 1:
        affiche_menu_principal(nom_fichier_courant, google_colonnes)
    else: # Si argument à la commande
        traitemente_argument(argv, google_colonnes)

if __name__ == '__main__':
    main(sys.argv[1:], L_COL_EXP_GOOGLE)
