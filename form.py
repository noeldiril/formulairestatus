import streamlit as st
from docx import Document
from num2words import num2words
from mail import MonEmail
from datetime import datetime
import streamlit.components.v1 as components
from num2words import num2words


def main():
    # Paramètres SMTP
    serveur_smtp = 'smtp.office365.com'
    port_smtp = 587
    utilisateur = 'aiformulas@outlook.fr'
    mot_de_passe = 'ndengineering95200.'

    # Création de l'objet MonEmail avec les paramètres SMTP
    email_sender = MonEmail(serveur_smtp, port_smtp, utilisateur, mot_de_passe)

    # Paramètres de l'e-mail
    destinataire = 'dirilnoel1@gmail.com'
    sujet = 'Formulaire création société : '
    message = """
    Bonjour,

    Vous trouverez ci-joints, les documents necessaires à la création de la sosciété.

    Cordialement,
    AIFormulas
    """

    components.html("""
        <script>
        const inputs = window.parent.document.querySelectorAll('input');
        inputs.forEach(input => {
            input.addEventListener('keydown', function(event) {
                if (event.key === 'Enter') {
                    event.preventDefault();
                }
            });
        });
        </script>
    """,
    height=0)

    # Affichage du texte au-dessus des boutons radio
    st.title("Choisissez votre type de société :")

    # Liste des options pour les boutons radio
    options = ["SAS (2 personnes)", "SAS (3 personnes)", "SASU"]

    # Affichage des boutons radio horizontalement
    selected_option = st.radio("", options, horizontal=True)

     # Récupérer la date actuelle
    aujourd_hui = datetime.now()
    date_aujourd_hui = aujourd_hui.strftime("%d/%m/%Y")

    # Affichage du formulaire en fonction de l'option sélectionnée
    if selected_option == "SAS (2 personnes)":
        st.header("Société par actions simplifiée (2 personnes)")
        with st.form(key='sas2', clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                denominationsas2 = st.text_input("Nom de la société : ", key="denominationsas2")
                capitalsas2 = st.text_input("Capital de la société : ", key="capitalsas2",on_change=None)
                sexe_associe_1sas2 = st.radio("Sexe du premier associé", key="sexe_associe_1sas2", options=["Monsieur", "Madame"])
                prenom_associe_1sas2 = st.text_input("Prénom du premier associé : ", key="prenom_associe_1sas2")
                nom_associe_1sas2 = st.text_input("Nom du premier associé : ", key="nom_associe_1sas2")
                date_associe_1sas2 = st.text_input("Date de naissance du premier associé (JJ/MM/AAAA): ", key="date_associe_1sas2")
                lieu_naissance_associe_1sas2 = st.text_input("Lieu de naissance du premier associé: ", key="lieu_naissance_associe_1sas2")
                code_postal_associe_1sas2 = st.text_input("Departement de naissance du premier associé : ", key="code_postal_associe_1sas2")
                nationalite_associe_1sas2 = st.text_input("Nationalité du premier associé : ", key="nationalite_associe_1sas2")
                adresse_associe_1sas2 = st.text_input("Adresse du premier associé : ", key="adresse_associe_1sas2")
                apport_associe_1sas2 = st.text_input("Apport du premier associé : ", key="apport_associe_1sas2")
                pere_associe_1sas2 = st.text_input("Nom et prénom du père du premier associé : ", key="pere_associe_1sas2")
                mere_associe_1sas2 = st.text_input("Nom et prénom de la mère du premier associé : ", key="mere_associe_1sas2")
                annee_bilansas2 = st.text_input("Année du premier bilan : ", key="annee_bilansas2")

            with col2:
                activitesas2 = st.text_input("Activité de la société : ", key="activitesas2")
                siege_socialsas2 = st.text_input("Siège social de la société : ", key="siege_socialsas2")
                sexe_associe_2sas2 = st.radio("Sexe du second associé", key="sexe_associe_2sas2", options=["Monsieur", "Madame"])
                prenom_associe_2sas2 = st.text_input("Prénom du second associé : ", key="prenom_associe_2sas2")
                nom_associe_2sas2 = st.text_input("Nom du second associé : ", key="nom_associe_2sas2")
                date_associe_2sas2 = st.text_input("Date de naissance du second associé (JJ/MM/AAAA): ", key="date_associe_2sas2")
                lieu_naissance_associe_2sas2 = st.text_input("Lieu de naissance du second associé: ", key="lieu_naissance_associe_2sas2")
                code_postal_associe_2sas2 = st.text_input("Departement de naissance du second associé : ", key="code_postal_associe_2sas2")
                nationalite_associe_2sas2 = st.text_input("Nationalité du second associé : ", key="nationalite_associe_2sas2")
                adresse_associe_2sas2 = st.text_input("Adresse du second associé : ", key="adresse_associe_2sas2")
                apport_associe_2sas2 = st.text_input("Apport du second associé : ", key="apport_associe_2sas2")
                ville_siege_socialsas2 = st.text_input("Ville du siège social : ", key="ville_siege_socialsas2")


            submittedsas2 = st.form_submit_button("Valider")

            if submittedsas2:
                if all([denominationsas2, capitalsas2, activitesas2, siege_socialsas2, prenom_associe_1sas2, nom_associe_1sas2, date_associe_1sas2,
                        lieu_naissance_associe_1sas2, code_postal_associe_1sas2, nationalite_associe_1sas2, adresse_associe_1sas2,
                        annee_bilansas2, prenom_associe_2sas2, nom_associe_2sas2, date_associe_2sas2, lieu_naissance_associe_2sas2,
                        code_postal_associe_2sas2, nationalite_associe_2sas2, adresse_associe_2sas2, apport_associe_1sas2,apport_associe_2sas2,ville_siege_socialsas2,pere_associe_1sas2]):
                    
                    remplacements_sas2 = {
                        '<CAPITAL>': capitalsas2,
                        '<SIEGESOCIAL>': siege_socialsas2,
                        '<DENOMINATION>': denominationsas2,
                        '<SEXEASSOCIE1>': sexe_associe_1sas2,
                        '<PRENOMASSOCIE1>': prenom_associe_1sas2,
                        '<NOMASSOCIE1>': nom_associe_1sas2.upper(),
                        '<DATEDENAISSANCEASSOCIE1<': date_associe_1sas2,
                        '<LIEUDENAISSANCEASSOCIE1>': lieu_naissance_associe_1sas2,
                        '<CODEPOSTALASSOCIE1>': code_postal_associe_1sas2,
                        '<NATIONALITEASSOCIE1>': nationalite_associe_1sas2,
                        '<ADRESSEDOMICILEASSOCIE1>': adresse_associe_1sas2,
                        '<APPORTASSOCIE1' : apport_associe_1sas2,
                        '<SEXEASSOCIE2>': sexe_associe_2sas2,
                        '<PRENOMASSOCIE2>': prenom_associe_2sas2,
                        '<NOMASSOCIE2>': nom_associe_2sas2.upper(),
                        '<DATEDENAISSANCEASSOCIE2<': date_associe_2sas2,
                        '<LIEUDENAISSANCEASSOCIE2>': lieu_naissance_associe_2sas2,
                        '<CODEPOSTALASSOCIE2>': code_postal_associe_2sas2,
                        '<NATIONALITEASSOCIE2>': nationalite_associe_2sas2,
                        '<ADRESSEDOMICILEASSOCIE2>': adresse_associe_2sas2,
                        '<APPORTASSOCIE2' : apport_associe_2sas2,
                        '<ACTIVITE>': activitesas2,
                        '<CAPITALENCHIFFRE>': capitalsas2,
                        '<ANNEEPREMIERBILAN>': annee_bilansas2,
                        '<VILLESIEGESOCIAL>' : ville_siege_socialsas2,
                        '<PARAPHESASSOCIE1>' : get_initials(prenom_associe_1sas2,nom_associe_1sas2),
                        '<PARAPHESASSOCIE2>' : get_initials(prenom_associe_2sas2,nom_associe_2sas2),
                        '<DATE>' : date_aujourd_hui,
                        '<PARENT1>' : pere_associe_1sas2,
                        '<PARENT2>' : mere_associe_1sas2,

                    }
                    
                    status_file = remplacer_valeurs('/Users/noeldiril/Desktop/Juridique/societe/sas2/Statuts SAS 2.docx', remplacements_sas2,"Status "+denominationsas2)
                    liste_action_file = remplacer_valeurs('/Users/noeldiril/Desktop/Juridique/societe/sas2/Liste souscripteurs actions SAS 2.docx', remplacements_sas2,"Liste souscripteurs "+denominationsas2)
                    dnc_file = remplacer_valeurs('/Users/noeldiril/Desktop/Juridique/societe/DNC.docx', remplacements_sas2,"DNC "+denominationsas2)

                    if status_file:
                        pieces_jointes = [status_file, liste_action_file, dnc_file]
                        email_sender.envoyer_email(destinataire, sujet+ denominationsas2, message, pieces_jointes)
                        st.success("Votre formulaire a bien été saisi et envoyé par e-mail.")
                    else:
                        st.error("Problème dans la mise à jour du fichier.")
                else:
                    st.error("Veuillez remplir tous les champs.")

    elif selected_option == "SAS (3 personnes)":
        st.header("Société par actions simplifiée (3 personnes)")
        with st.form(key='sas3', clear_on_submit=True):
            col1, col2, col3 = st.columns(3)
            with col1:
                denominationsas3 = st.text_input("Nom de la société : ", key="denominationsas3")
                capitalsas3 = st.text_input("Capital de la société : ", key="capitalsas3")
                sexe_associe_1sas3 = st.radio("Sexe du premier associé", key="sexe_associe_1sas3", options=["Monsieur", "Madame"])
                prenom_associe_1sas3 = st.text_input("Prénom du premier associé : ", key="prenom_associe_1sas3")
                nom_associe_1sas3 = st.text_input("Nom du premier associé : ", key="nom_associe_1sas3")
                date_associe_1sas3 = st.text_input("Date de naissance du premier associé (JJ/MM/AAAA): ", key="date_associe_1sas3")
                lieu_naissance_associe_1sas3 = st.text_input("Lieu de naissance du premier associé: ", key="lieu_naissance_associe_1sas3")
                code_postal_associe_1sas3 = st.text_input("Departement de naissance du premier associé : ", key="code_postal_associe_1sas3")
                nationalite_associe_1sas3 = st.text_input("Nationalité du premier associé : ", key="nationalite_associe_1sas3")
                adresse_associe_1sas3 = st.text_input("Adresse du premier associé : ", key="adresse_associe_1sas3")
                apport_associe_1sas3 = st.text_input("Apport du premier associé : ", key="apport_associe_1sas3")
                pere_associe_1sas3 = st.text_input("Nom et prénom du père du premier associé : ", key="pere_associe_1sas3")

            with col2:
                activitesas3 = st.text_input("Activité de la société : ", key="activitesas3")
                siege_socialsas3 = st.text_input("Siège social de la société : ", key="siege_socialsas3")
                sexe_associe_2sas3 = st.radio("Sexe du second associé", key="sexe_associe_2sas3", options=["Monsieur", "Madame"])
                prenom_associe_2sas3 = st.text_input("Prénom du second associé : ", key="prenom_associe_2sas3")
                nom_associe_2sas3 = st.text_input("Nom du second associé : ", key="nom_associe_2sas3")
                date_associe_2sas3 = st.text_input("Date de naissance du second associé (JJ/MM/AAAA): ", key="date_associe_2sas3")
                lieu_naissance_associe_2sas3 = st.text_input("Lieu de naissance du second associé: ", key="lieu_naissance_associe_2sas3")
                code_postal_associe_2sas3 = st.text_input("Departement de naissance du second associé : ", key="code_postal_associe_2sas3")
                nationalite_associe_2sas3 = st.text_input("Nationalité du second associé : ", key="nationalite_associe_2sas3")
                adresse_associe_2sas3 = st.text_input("Adresse du second associé : ", key="adresse_associe_2sas3")
                apport_associe_2sas3 = st.text_input("Apport du second associé : ", key="apport_associe_2sas3")
                mere_associe_1sas3 = st.text_input("Nom et prénom de la mère du premier associé : ", key="mere_associe_1sas3")

            with col3:
                annee_bilansas3 = st.text_input("Année du premier bilan: ", key="annee_bilansas3")
                ville_siege_socialsas3 = st.text_input("Ville du siège social : ", key="ville_siege_socialsas3")
                sexe_associe_3sas3 = st.radio("Sexe du troisième associé", key="sexe_associe_3sas3", options=["Monsieur", "Madame"])
                prenom_associe_3sas3 = st.text_input("Prénom du troisième associé : ", key="prenom_associe_3sas3")
                nom_associe_3sas3 = st.text_input("Nom du troisième associé : ", key="nom_associe_3sas3")
                date_associe_3sas3 = st.text_input("Date de naissance du troisième associé (JJ/MM/AAAA): ", key="date_associe_3sas3")
                lieu_naissance_associe_3sas3 = st.text_input("Lieu de naissance du troisième associé: ", key="lieu_naissance_associe_3sas3")
                code_postal_associe_3sas3 = st.text_input("Departement de naissance du troisième associé : ", key="code_postal_associe_3sas3")
                nationalite_associe_3sas3 = st.text_input("Nationalité du troisième associé : ", key="nationalite_associe_3sas3")
                adresse_associe_3sas3 = st.text_input("Adresse du troisième associé : ", key="adresse_associe_3sas3")
                apport_associe_3sas3 = st.text_input("Apport du troisième associé : ", key="apport_associe_3sas3")


            submittedsas3 = st.form_submit_button("Valider")

            if submittedsas3:
                if all([denominationsas3, capitalsas3, activitesas3, siege_socialsas3, prenom_associe_1sas3, nom_associe_1sas3, date_associe_1sas3,
                        lieu_naissance_associe_1sas3, code_postal_associe_1sas3, nationalite_associe_1sas3, adresse_associe_1sas3,
                        annee_bilansas3, prenom_associe_2sas3, nom_associe_2sas3, date_associe_2sas3, lieu_naissance_associe_2sas3,
                        code_postal_associe_2sas3, nationalite_associe_2sas3, adresse_associe_2sas3, apport_associe_1sas3,apport_associe_2sas3,
                        ville_siege_socialsas3,pere_associe_1sas3, mere_associe_1sas3,prenom_associe_3sas3,nom_associe_3sas3,date_associe_3sas3,
                        lieu_naissance_associe_3sas3,code_postal_associe_3sas3,nationalite_associe_3sas3,adresse_associe_3sas3,apport_associe_3sas3]):
                    
                    remplacements_sas3 = {
                        '<CAPITAL>': capitalsas3,
                        '<SIEGESOCIAL>': siege_socialsas3,
                        '<DENOMINATION>': denominationsas3,
                        '<SEXEASSOCIE1>': sexe_associe_1sas3,
                        '<PRENOMASSOCIE1>': prenom_associe_1sas3,
                        '<NOMASSOCIE1>': nom_associe_1sas3.upper(),
                        '<DATEDENAISSANCEASSOCIE1<': date_associe_1sas3,
                        '<LIEUDENAISSANCEASSOCIE1>': lieu_naissance_associe_1sas3,
                        '<CODEPOSTALASSOCIE1>': code_postal_associe_1sas3,
                        '<NATIONALITEASSOCIE1>': nationalite_associe_1sas3,
                        '<ADRESSEDOMICILEASSOCIE1>': adresse_associe_1sas3,
                        '<APPORTASSOCIE1' : apport_associe_1sas3,
                        '<SEXEASSOCIE2>': sexe_associe_2sas3,
                        '<PRENOMASSOCIE2>': prenom_associe_2sas3,
                        '<NOMASSOCIE2>': nom_associe_2sas3.upper(),
                        '<DATEDENAISSANCEASSOCIE2<': date_associe_2sas3,
                        '<LIEUDENAISSANCEASSOCIE2>': lieu_naissance_associe_2sas3,
                        '<CODEPOSTALASSOCIE2>': code_postal_associe_2sas3,
                        '<NATIONALITEASSOCIE2>': nationalite_associe_2sas3,
                        '<ADRESSEDOMICILEASSOCIE2>': adresse_associe_2sas3,
                        '<APPORTASSOCIE2' : apport_associe_2sas3,
                        '<ACTIVITE>': activitesas3,
                        '<CAPITALENCHIFFRE>': capitalsas3,
                        '<ANNEEPREMIERBILAN>': annee_bilansas3,
                        '<VILLESIEGESOCIAL>' : ville_siege_socialsas3,
                        '<SEXEASSOCIE3>': sexe_associe_3sas3,
                        '<PRENOMASSOCIE3>': prenom_associe_3sas3,
                        '<NOMASSOCIE3>': nom_associe_3sas3.upper(),
                        '<DATEDENAISSANCEASSOCIE3<': date_associe_3sas3,
                        '<LIEUDENAISSANCEASSOCIE3>': lieu_naissance_associe_3sas3,
                        '<CODEPOSTALASSOCIE3>': code_postal_associe_3sas3,
                        '<NATIONALITEASSOCIE3>': nationalite_associe_3sas3,
                        '<ADRESSEDOMICILEASSOCIE3>': adresse_associe_3sas3,
                        '<APPORTASSOCIE3' : apport_associe_3sas3,
                        '<PARAPHESASSOCIE1>' : get_initials(prenom_associe_1sas3,nom_associe_1sas3),
                        '<PARAPHESASSOCIE2>' : get_initials(prenom_associe_2sas3,nom_associe_2sas3),
                        '<PARAPHESASSOCIE3>' : get_initials(prenom_associe_3sas3,nom_associe_3sas3),
                        '<PARENT1>' : pere_associe_1sas3,
                        '<PARENT2>' : mere_associe_1sas3,
                        '<DATE>' : date_aujourd_hui
                    }
                    
                    status_file = remplacer_valeurs('/Users/noeldiril/Desktop/Juridique/societe/sas2/Statuts SAS 2.docx', remplacements_sas3,"Status "+denominationsas3)
                    liste_action_file = remplacer_valeurs('/Users/noeldiril/Desktop/Juridique/societe/sas2/Liste souscripteurs actions SAS 2.docx', remplacements_sas3,"Liste souscripteurs "+denominationsas3)
                    dnc_file = remplacer_valeurs('/Users/noeldiril/Desktop/Juridique/societe/DNC.docx', remplacements_sas3,"DNC "+denominationsas3)

                    if status_file:
                        pieces_jointes = [status_file, liste_action_file, dnc_file]
                        email_sender.envoyer_email(destinataire, sujet+ denominationsas3, message, pieces_jointes)
                        st.success("Votre formulaire a bien été saisi et envoyé par e-mail.")
                    else:
                        st.error("Problème dans la mise à jour du fichier.")
                else:
                    st.error("Veuillez remplir tous les champs.")

    elif selected_option == "SASU":
        st.header("Société par actions simplifiée unipersonnelle")
        with st.form(key='sasu', clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                denominationsasu = st.text_input("Nom de la société : ", key="denominationsasu")
                capitalsasu = st.text_input("Capital de la société : ", key="capitalsasu")
                sexe_associe_1sasu = st.radio("Sexe du président ", key="sexe_associe_1sasu", options=["Monsieur", "Madame"])
                prenom_associe_1sasu = st.text_input("Prénom du président  : ", key="prenom_associe_1sasu")
                nom_associe_1sasu = st.text_input("Nom du président  : ", key="nom_associe_1sasu")
                date_associe_1sasu = st.text_input("Date de naissance du président  (JJ/MM/AAAA): ", key="date_associe_1sasu")
                lieu_naissance_associe_1sasu = st.text_input("Lieu de naissance du président : ", key="lieu_naissance_associe_1sasu")
                code_postal_associe_1sasu = st.text_input("Departement de naissance du président  : ", key="code_postal_associe_1sasu")
                nationalite_associe_1sasu = st.text_input("Nationalité du président  : ", key="nationalite_associe_1sasu")
                
            with col2:
                activitesasu = st.text_input("Activité de la société : ", key="activitesasu")
                siege_socialsasu = st.text_input("Siège social de la société : ", key="siege_socialsasu")
                ville_siege_socialsasu = st.text_input("Ville du siège social : ", key="ville_siege_socialsasu")
                adresse_associe_1sasu = st.text_input("Adresse du président  : ", key="adresse_associe_1sasu")
                apport_associe_1sasu = st.text_input("Apport du président  : ", key="apport_associe_1sasu")
                pere_associe_1sasu = st.text_input("Nom et prénom du père du président  : ", key="pere_associe_1sasu")
                mere_associe_1sasu = st.text_input("Nom et prénom de la mère du président  : ", key="mere_associe_1sasu")
                annee_bilansasu = st.text_input("Année du premier bilan : ", key="annee_bilansasu")


            submittedsasu = st.form_submit_button("Valider")

            if submittedsasu:
                if all([denominationsasu, capitalsasu, activitesasu, siege_socialsasu, prenom_associe_1sasu, nom_associe_1sasu, date_associe_1sasu,
                        lieu_naissance_associe_1sasu, code_postal_associe_1sasu, nationalite_associe_1sasu, adresse_associe_1sasu,
                        annee_bilansasu, apport_associe_1sasu,ville_siege_socialsasu,pere_associe_1sasu, mere_associe_1sasu]):
                    
                    remplacements_sasu = {
                        '<CAPITAL>': capitalsasu,
                        '<SIEGESOCIAL>': siege_socialsasu,
                        '<DENOMINATION>': denominationsasu,
                        '<SEXEASSOCIE1>': sexe_associe_1sasu,
                        '<PRENOMASSOCIE1>': prenom_associe_1sasu,
                        '<NOMASSOCIE1>': nom_associe_1sasu.upper(),
                        '<DATEDENAISSANCEASSOCIE1<': date_associe_1sasu,
                        '<LIEUDENAISSANCEASSOCIE1>': lieu_naissance_associe_1sasu,
                        '<CODEPOSTALASSOCIE1>': code_postal_associe_1sasu,
                        '<NATIONALITEASSOCIE1>': nationalite_associe_1sasu,
                        '<ADRESSEDOMICILEASSOCIE1>': adresse_associe_1sasu,
                        '<APPORTASSOCIE1' : apport_associe_1sasu,
                        '<ACTIVITE>': activitesasu,
                        '<CAPITALENCHIFFRE>': capitalsasu,
                        '<ANNEEPREMIERBILAN>': annee_bilansasu,
                        '<VILLESIEGESOCIAL>' : ville_siege_socialsasu,
                        '<PARAPHESASSOCIE1>' : get_initials(prenom_associe_1sasu,nom_associe_1sasu),
                        '<PARENT1>' : pere_associe_1sasu,
                        '<PARENT2>' : mere_associe_1sasu,
                        '<DATE>' : date_aujourd_hui
                    }
                    
                    status_file = remplacer_valeurs('/Users/noeldiril/Desktop/Juridique/societe/sasu/Statuts SASU.docx', remplacements_sasu,"Status "+denominationsasu)
                    liste_action_file = remplacer_valeurs('/Users/noeldiril/Desktop/Juridique/societe/sasu/Liste souscripteurs actions SASu.docx', remplacements_sasu,"Liste souscripteurs "+denominationsasu)
                    dnc_file = remplacer_valeurs('/Users/noeldiril/Desktop/Juridique/societe/DNC.docx', remplacements_sasu,"DNC "+denominationsasu)

                    if status_file:
                        pieces_jointes = [status_file, liste_action_file, dnc_file]
                        email_sender.envoyer_email(destinataire, sujet+ denominationsasu, message, pieces_jointes)
                        st.success("Votre formulaire a bien été saisi et envoyé par e-mail.")
                    else:
                        st.error("Problème dans la mise à jour du fichier.")
                else:
                    st.error("Veuillez remplir tous les champs.")


def get_initials(prenom, nom):
    return f"{prenom[0].upper()}{nom[0].upper()}"

def remplacer_valeurs(docx_file, remplacements, nomfichier):
    doc = Document(docx_file)

    def remplacer_dans_paragraph(paragraph, remplacements):
        for old_value, new_value in remplacements.items():
            if old_value in paragraph.text:
                for run in paragraph.runs:
                    if old_value in run.text:
                        run.text = run.text.replace(old_value, new_value)

    def remplacer_dans_table(table, remplacements):
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    remplacer_dans_paragraph(paragraph, remplacements)

    # Remplacement dans les paragraphes principaux
    for paragraph in doc.paragraphs:
        remplacer_dans_paragraph(paragraph, remplacements)

    # Remplacement dans les tableaux du corps principal du document
    for table in doc.tables:
        remplacer_dans_table(table, remplacements)

    # Remplacement dans les sections (en-têtes et pieds de page)
    for section in doc.sections:
        # Remplacement dans les en-têtes
        for header in section.header.paragraphs:
            remplacer_dans_paragraph(header, remplacements)
        for table in section.header.tables:
            remplacer_dans_table(table, remplacements)

        # Remplacement dans les pieds de page
        for footer in section.footer.paragraphs:
            remplacer_dans_paragraph(footer, remplacements)
        for table in section.footer.tables:
            remplacer_dans_table(table, remplacements)

    # Sauvegarde du nouveau document
    new_docx_file = nomfichier + ".docx"
    doc.save(new_docx_file)

    return new_docx_file





if __name__ == "__main__":
    main()
