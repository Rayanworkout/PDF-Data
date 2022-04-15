import os
from datetime import date
import re
import fitz
import pandas as pd

data_dirigeant = {"Prénom": "", "Nom": "", "Prénom & Nom": "",
                  "Genre": "Mr ?", "Fonction": "Dirigeant", "Société": "", "Catégorie": "Entreprise", "Type": "",
                  "Groupe": "",
                  "Réseau": "Region Sud Comité",
                  "Comment : Apporteur / Sté Apportée ou  Comité / Apporteur / Métier": "",
                  "Téléphone mobile": "", "Adresse de messagerie": "", "Rue (bureau)": "", "Code postal (bureau)": "",
                  "Date contact": date.today().strftime('%d/%m/%Y'), "Date Comité": "", "CA 2020 K€": "",
                  "CA 2023 K€": "", "Effectif 20": ""}

data_partenaire = {"Prénom": "", "Nom": "", "Prénom & Nom": "",
                   "Genre": "Mr ?", "Fonction": "Partenaire", "Société": "", "Catégorie": "Entreprise", "Type": "",
                   "Groupe": "",
                   "Réseau": "Region Sud Comité", "Comment : Apporteur / Sté Apportée ou  Comité / Apporteur / Métier":
                   "Apporteur Region Sud Invest", "Téléphone mobile": "",
                   "Adresse de messagerie": "", "Rue (bureau)": "",
                   "Code postal (bureau)": "", "Date contact": date.today().strftime('%d/%m/%Y'), "Date Comité": "",
                   "CA 2020 K€": "", "CA 2023 K€": "", "Effectif 20": ""}


def get_content(pdf, file_text):
    # LOOP THROUGH ALL PAGES AND ADD CONTENT TO A VARIABLE
    for x in range(pdf.page_count):
        page = pdf.load_page(x).get_text()
        file_text += page
    return file_text.lower()


def get_name(pdf):  # GETTING COMPANY NAME BY PRETTYFYING FILENAME
    nom_societe = pdf.split('.pdf')[0]
    for char in nom_societe:
        if char.isdigit():
            nom_societe = nom_societe.replace(char, '')
    signs = ['_', '-', 'OK', 'Ok', 'NON', 'non', 'Non']
    for sign in signs:
        nom_societe = nom_societe.replace(sign, ' ')
    data_dirigeant['Société'] = nom_societe.strip()
    data_partenaire['Société'] = nom_societe.strip()


def get_dirigeant_and_partner(content):
    # DIRIGEANT
    if len(content.split('dirigeant(e)(s) :')) == 1:
        dirigeant = content.split('dirigeant(e)(s):')[1]
    else:
        dirigeant = content.split('dirigeant(e)(s) :')[1]
    dirigeant = dirigeant.split()

    if dirigeant[0] == 'nom' or dirigeant[0] == 'prénom':
        nom_dirigeant, prenom_dirigeant = dirigeant[7].capitalize(), dirigeant[8].capitalize()
        dirigeant = f'{prenom_dirigeant} {nom_dirigeant}'
    else:
        nom_dirigeant, prenom_dirigeant = dirigeant[1].capitalize(), dirigeant[0].capitalize()
        dirigeant = f'{prenom_dirigeant} {nom_dirigeant}'
    data_dirigeant['Nom'] = nom_dirigeant
    data_dirigeant['Prénom'] = prenom_dirigeant
    data_dirigeant['Prénom & Nom'] = dirigeant

    # PARTNER
    if len(content.split('partenaire(s) :')) == 1:
        partner = content.split('partenaire(s):')[1]
    else:
        partner = content.split('partenaire(s) :')[1]
    partner = partner.split()

    if partner[0] == 'nom' or partner[0] == 'prénom':
        if partner[7] == '›' or partner[7] == "forme":
            prenom_partner = ''
            nom_partner = ''
        else:
            liste = ['sas', ':']
            if partner[9] in liste:
                prenom_partner = partner[8].capitalize().replace(',', '')
                nom_partner = partner[7].capitalize().replace(',', '')

            else:
                prenom_partner = partner[8].capitalize().replace(',', '')
                nom_partner = f"{partner[7].capitalize().replace(',', '')} {partner[9].capitalize()}"

    else:
        prenom_partner = partner[0].capitalize()
        nom_partner = partner[1].capitalize()

    my_list = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9', ',', '.']  # LIST OF CHARACTERS I DON'T WANT IN A NAME
    for char in my_list:
        if char in prenom_partner:
            prenom_partner = prenom_partner.replace(char, '')
    for char in my_list:
        if char in nom_partner:
            nom_partner = nom_partner.replace(char, '')

    partner = f'{prenom_partner} {nom_partner}'
    data_partenaire['Nom'] = nom_partner
    data_partenaire['Prénom'] = prenom_partner
    data_partenaire['Prénom & Nom'] = partner


def department_date_comite(content):
    department = content.split('département :')[1]
    department = department.split()
    department = department[0]

    data_dirigeant['Code postal (bureau)'] = department
    data_partenaire['Code postal (bureau)'] = department

    date_comite = content.split("comité :")[1]  # CAN BE TAKEN FROM FILENAME AS WELL
    date_comite = date_comite.split()[0]

    data_dirigeant['Date Comité'] = date_comite
    data_partenaire['Date Comité'] = date_comite


def comment(content):
    date_comite = data_dirigeant['Date Comité']
    montant_pret = content.split('rsi demandé')[1]
    montant_pret = montant_pret.split()[0]

    prescripteur = content.split('prescripteur :')[1]
    prescripteur = prescripteur.split()
    type_prescripteur = ""

    for x in range(0, 6):
        if prescripteur[x] != '›':
            type_prescripteur += f" {prescripteur[x]}"
        elif prescripteur[x] == '›':
            break
    activite = content.split('activité de l’entreprise')[1]
    activite = activite.split()
    activite = activite[:20]
    data_dirigeant['Comment : Apporteur / Sté Apportée ou  Comité / Apporteur / Métier'] = \
        f"Comité du {date_comite} - {montant_pret}K € - {type_prescripteur} - {' '.join(activite)}"


def get_emails_and_phone_numbers(content):
    mails, phones = [], []
    email_rexes = [
        # EMAILS
        r'(^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)',
        r'\S+@\S+'
    ]
    phone_rexes = [
        # PHONE NUMBERS
        r'0[1-3-4-6-7-9].\d\d.\d\d.\d\d.\d\d',
        r'0[1-3-4-6-7-9]\d\d\d\d\d\d\d\d',
        r'\+33.[1-3-4-6-7-9].\d\d\d\d\d\d\d\d',
        r'\+33.[1-3-4-6-7-9].\d\d.\d\d.\d\d.\d\d',
        r'\+33[1-3-4-6-7-9]\d\d\d\d\d\d'
    ]

    for email_rex in email_rexes:
        mails += re.findall(email_rex, content)
    for phone_rex in phone_rexes:
        phones += re.findall(phone_rex, content)

    if len(data_partenaire['Prénom & Nom']) > 1:
        if len(mails) == 0:
            data_dirigeant['Adresse de messagerie'] = 'Non Renseignée'
        elif len(mails) == 1:
            data_dirigeant['Adresse de messagerie'] = mails[0]
        elif len(mails) == 2:
            data_dirigeant['Adresse de messagerie'] = mails[0]
            data_partenaire['Adresse de messagerie'] = mails[1]
        if len(phones) == 0:
            data_dirigeant['Téléphone mobile'] = 'Non Renseigné'
        elif len(phones) == 1:
            data_dirigeant['Téléphone mobile'] = phones[0]
        elif len(phones) == 2:
            data_dirigeant['Téléphone mobile'] = phones[0]
            data_partenaire['Téléphone mobile'] = phones[1]

    elif len(data_partenaire['Prénom & Nom']) == 0:
        if len(mails) == 0:
            data_dirigeant['Adresse de messagerie'] = 'Non Renseignée'
        elif len(mails) == 1:
            data_dirigeant['Adresse de messagerie'] = mails[0]
        elif len(mails) == 2:
            data_dirigeant['Adresse de messagerie'] = mails[0]
        if len(phones) == 0:
            data_dirigeant['Téléphone mobile'] = 'Non Renseigné'
        elif len(phones) == 1:
            data_dirigeant['Téléphone mobile'] = phones[0]
        elif len(phones) == 2:
            data_dirigeant['Téléphone mobile'] = phones[0]


def loop():  # GETTING DATA OF ALL FILES
    df = pd.DataFrame()
    with os.scandir('Fichiers/') as folder:
        for file in folder:
            if file.name.endswith(".pdf") and file.is_file():
                file_text = ""
                content = get_content(fitz.open(f'Fichiers/{file.name}'), file_text)

                get_name(file.name)
                get_dirigeant_and_partner(content)
                department_date_comite(content)
                comment(content)
                get_emails_and_phone_numbers(content)
                # CREATING DATAFRAME
                df = df.append(data_dirigeant, ignore_index=True)
                if len(data_partenaire['Nom']) > 0:
                    df = df.append(data_partenaire, ignore_index=True)
                # CREATING EXCEL FILE
                df.to_excel('OUTPUT/Data.xlsx', index=False)


loop()
