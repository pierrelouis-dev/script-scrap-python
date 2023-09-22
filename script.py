import csv
import unicodedata
from googlesearch import search
import pandas as pd

personnes_non_trouvees = []  # Liste pour stocker les personnes non trouvées
nombre_personnes_non_trouvees = 0  # Compteur pour le nombre de personnes non trouvées

with open('base_eleve.csv', newline='') as csvfile:
    diplomes = list(csv.DictReader(csvfile))

with open('liste_entreprise.csv', newline='') as csvfile:
    entreprises = list(csv.DictReader(csvfile))

def remove_accents(input_str):
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    return u"".join([c for c in nfkd_form if not unicodedata.combining(c)])

personnes_trouvees = []  # Liste pour stocker les personnes trouvées

for diplome in diplomes:
    nom = remove_accents(diplome['nom'])
    prenom = remove_accents(diplome['prenom'])
    personne_trouvee = False

    for entreprise in entreprises:
        if nom.lower() in remove_accents(entreprise['nomUniteLegale'].lower()) or nom.lower() in remove_accents(entreprise['nomUsageUniteLegale'].lower()) and (
                prenom.lower() in remove_accents(entreprise['prenom1UniteLegale'].lower()) or
                prenom.lower() in remove_accents(entreprise['prenom2UniteLegale'].lower()) or
                prenom.lower() in remove_accents(entreprise['prenom3UniteLegale'].lower()) or
                prenom.lower() in remove_accents(entreprise['prenom4UniteLegale'].lower()) or
                prenom.lower() in remove_accents(entreprise['prenomUsuelUniteLegale'].lower())
        ):
            personne_trouvee_info = {
                'Nom': nom,
                'Prénom': prenom,
                'Code postal': entreprise['codePostalEtablissement'],
                'Adresse': f"{entreprise['numeroVoieEtablissement']} {entreprise['typeVoieEtablissement']} {entreprise['libelleVoieEtablissement']} {entreprise['codeCommuneEtablissement']} {entreprise['libelleCommuneEtablissement']}",
                'Entreprise créée le': entreprise['dateCreationEtablissement'],
                'SIREN': entreprise['siren'],
                'SIRET': entreprise['siret']
            }
            personnes_trouvees.append(personne_trouvee_info)
            personne_trouvee = True
            break

    if not personne_trouvee:
        nombre_personnes_non_trouvees += 1
        personnes_non_trouvees.append((nom, prenom))

# Créer des DataFrames pour les personnes trouvées et non trouvées
df_personnes_trouvees = pd.DataFrame(personnes_trouvees)
df_personnes_non_trouvees = pd.DataFrame(personnes_non_trouvees, columns=['Nom', 'Prénom'])

# Exporter les DataFrames vers des fichiers Excel
df_personnes_trouvees.to_excel('personnes_trouvees.xlsx', index=False)
df_personnes_non_trouvees.to_excel('personnes_non_trouvees.xlsx', index=False)
