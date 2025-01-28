#!/usr/bin/env python
# coding: utf-8

# # Notebook PréparaIon des données

# ## 1. Compréhension des données

# ### Identifier/collecter/compléter les données nécessaires à la réalisation du projet : 
# >J'ai commencé par un pré-nettoyage des 4 datasets initiaux, qui contenaient de nombreuses données inutiles pour l'étude. J'ai fusionné les données des étudiants et du personnel dans un seul fichier, ce qui est plus pratique pour visualiser les répartitions et effectuer des analyses globales. Ce pré-nettoyage me permettra de mieux me familiariser avec les données avant d'effectuer un nettoyage plus approfondi après la phase de compréhension.

# But final: rassembler les fichiers 2023 et 2024 ensemble et finir avec un seul et unique df nettoyé et exploitable pour nos KPI.

# ### Questionnaire UGA étudiants - 2023 et 2024

# ## Plan d'action :
# 
# -> Voici comment je vais procéder: 
# D’abord je vais commencer par le 2023, je vais supprimer les colonnes que nous n’utiliserons pas dans notre projet, et garder que celles dont on aura besoin pour calculer les KPI. Je chargerai ensuite le 2024 et je ferai en sorte d’avoir la meme structure (donc les memes colonnes que le 2023) pour ensuite les rassembler. 

# In[128]:


import pandas as pd

# Charger le deuxième fichier Excel contenant les données des étudiants
file_path_students = "/Users/amaloubaba/Downloads/UGA-enquete-mobilite-2023-etudiants.xlsx"
df_students_2023 = pd.read_excel(file_path_students)


# Afin de facilier l'analyse et l'affichage, je. vais commencer par supprimer les colonnes qui ne me seront pas pas utiles pour ce projet. 

# In[129]:


# Afficher toutes les colonnes disponibles dans le fichier
print("Liste des colonnes dans le fichier :")
for col in df_students_2023.columns:
    print(col)


# In[130]:


# Liste des colonnes à supprimer
colonnes_a_supprimer = [
    '1. Dans quel établissement principal êtes-vous inscrit.e ?',
    '2. En quel diplôme êtes-vous inscrit.e ?',
    '3. Autre :',
    '4. Quel est votre site d’enseignement principal ?',
    '5. Autre :',
    "6. Aurez-vous une interruption de présence sur votre lieu d'enseignement cette année (ex : stage, césure...) ?",
    '7. Hors vacances universitaires, combien de semaines durera cette interruption de présence ?',
    '10. Distance totale ALLER / RETOUR en marche (km / jour)',
    '13. Distance totale ALLER / RETOUR en trottinette électrique (km / jour)',
    '14. Distance totale ALLER / RETOUR en moto / scooter (km / jour)',
    '15. Nombre de personnes sur le 2 roues motorisées',
    '18. Nombre de personnes dans la voiture',
    '19. Distance totale ALLER / RETOUR en bus (km / jour)',
    '20. Distance totale ALLER / RETOUR en tramway (km / jour)',
    '21. Distance totale ALLER / RETOUR en train (km / jour)',
    "23. Parmi vos [En_moyenne_lorsque_vous_netes_pas_en_vac] jour(s) de déplacements par semaine pour vous rendre sur votre lieu d'études, combien de jours ce deuxième trajet concerne-t-il ?",
    '25. Distance totale ALLER / RETOUR en marche (km / jour)',
    '28. Distance totale ALLER / RETOUR en trottinette électrique (km / jour)',
    '29. Distance totale ALLER / RETOUR en moto / scooter (km / jour)',
    '30. Nombre de personnes sur le 2 roues motorisées',
    '33. Nombre de personnes dans la voiture',
    '34. Distance totale ALLER / RETOUR en bus (km / jour)',
    '35. Distance totale ALLER / RETOUR en tramway (km / jour)',
    '36. Distance totale ALLER / RETOUR en train (km / jour)',
    '37. Clé',
    '39. Date de dernière modification',
    '40. Date de dernier enregistrement',
    '41. Temps de saisie',
    '42. Langue',
    '43. Progression',
    '44. Dernière question saisie',
    '45. Origine',
    '46. Appareil utilisé pour la saisie'
]

# Supprimer les colonnes inutiles
df_stu_2023_clean = df_students_2023.drop(columns=colonnes_a_supprimer, errors='ignore')

# Remplacer les en-têtes par la deuxième ligne (actuelle) et supprimer la première ligne
df_stu_2023_clean.columns = df_stu_2023_clean.iloc[0]
df_cleaned_2023 = df_stu_2023_clean[1:].reset_index(drop=True)

df_cleaned_2023


# In[131]:


# Renommer la colonne contenant NaN en 'etude_duree'
df_cleaned_2023.columns = df_cleaned_2023.columns.fillna('etude_duree')

df_cleaned_2023


# In[132]:


# Afficher toutes les colonnes disponibles dans le fichier
print("Liste des colonnes dans le fichier :")
for col in df_cleaned_2023.columns:
    print(col)


# In[133]:


# Sauvegarde des résultats
output_file_path = "/Users/amaloubaba/Downloads/UGA_etudiants_2023__.xlsx"
df_cleaned_2023.to_excel(output_file_path, index=False)


# ## Nous y voyons plus clair, passons maintenant au fichier 2024.

# In[134]:


# Charger le deuxième fichier Excel contenant les données des étudiants
file_path_students = "/Users/amaloubaba/Downloads/_UGA-enquÃªte-mobilitÃ©-2024-Ã©tudiants.xlsx"
df_students_2024 = pd.read_excel(file_path_students)


# In[135]:


# Afficher toutes les colonnes disponibles dans le fichier
print("Liste des colonnes dans le fichier :")
for col in df_students_2024.columns:
    print(col)


# supprimons les colonnes que nous n'utiliserons pas et gardons les mêmes que celles qu'on a gardé pour l'année 2023

# In[136]:


# Liste des colonnes à garder:
#colonnes_a_sgarder = [
#    'N°Obs',
#    '6. nb_jours_semaine',
#    '11. D1_modes_transport',
#    '13. D1_vélo_km',
#    '14. D1_VAE_km',
#    '18. D1_voiture_km',
#    '19. D1_voiture_motorisation',
#    '28. D2_déclaration',
#    '32. D2_modes_transport',
#    '34. D2_vélo_km',
#    '35. D2_VAE_km',
#    '39. D2_voiture_km',
#    '40. D2_voiture_motorisation',
#    '66. Date de saisie'
#]


# Colonnes à supprimer
colonnes_a_supprimer = [
    '1. composantes', '2. diplome', '3. Autre diplôme :', '4. site', '5. Autre site :', '7. interruption', 
    '12. D1_marche_km', '15. D1_trottinette_km', '16. D1_moto_km', '17. D1_moto_nb_passagers', 
    '20. D1_voiture_nb_passagers', '21. covoiturage_intention', '22. Avec qui covoiturez-vous ?', 
    '23. voiture_raisons', '24. Autre(s) raison(s) :', '25. D1_bus_km', '26. D1_tramway_km', 
    '27. D1_train_km', '29. D2_code_postal', '30. D2_commune', '31. no_code_postal', 
    '33. D2_marche_km', '36. D2_trottinette_km', '37. D2_moto_km', '38. D2_moto_nb_passagers', 
    '41. D2_voiture_nb_passagers', '42. D2_bus_km', '43. D2_tramway_km', '44. D2_train_km', 
    '46. De quel type de mobilité s\'agit-il ?', 
    '47. Si vous ne connaissez pas encore votre lieu de stage ou d\'alternance et vos modes de déplacement pour vous y rendre, cochez la case ci-après.',
    '48. stage_duree', '55. stage_modes_transport', '56. stage_marche_km', '57. stage_vélo_km', 
    '58. stage_VAE_km', '59. stage_trottinette_km', '60. stage_moto_km', '61. stage_voiture_km', 
    '62. stage_bus_km', '63. stage_tramway_km', '64. stage_train_km', '65. Clé', 
    '67. Date de dernière modification', '68. Date de dernier enregistrement', 
    '69. Temps de saisie', '70. Langue', '71. Progression', '72. Dernière question saisie', 
    '73. Origine', '74. Appareil utilisé pour la saisie'
]



# Supprimer les colonnes inutiles
df_stu_2024_clean = df_students_2024.drop(columns=colonnes_a_supprimer, errors='ignore')


df_stu_2024_clean


# In[137]:


# Sauvegarde des résultats
output_file_path = "/Users/amaloubaba/Downloads/UGA_etudiants_2024__.xlsx"
df_stu_2024_clean.to_excel(output_file_path, index=False)


# ## Concaténons les deux fichiers:

# In[139]:


df_cleaned_2023


# In[140]:


# Afficher toutes les colonnes disponibles dans le fichier
print("Liste des colonnes dans le fichier :")
for col in df_cleaned_2023.columns:
    print(col)


# In[141]:


df_stu_2024_clean


# In[142]:


# Afficher toutes les colonnes disponibles dans le fichier
print("Liste des colonnes dans le fichier :")
for col in df_stu_2024_clean.columns:
    print(col)


# Nous remarquons que les 2 fichiers n'ont pas les mêmes noms de colonnes, afin d'éviter toute erreur lors de la fusion, nous allons les modifier 

# In[145]:


import pandas as pd

# Charger les fichiers Excel
file_2023 = '/Users/amaloubaba/Downloads/UGA_etudiants_2023__.xlsx'
file_2024 = '/Users/amaloubaba/Downloads/UGA_etudiants_2024__.xlsx'

df_2023 = pd.read_excel(file_2023)
df_2024 = pd.read_excel(file_2024)

# Renommer les colonnes de 2024 pour qu'elles correspondent à celles de 2023
rename_columns_2024 = {
    "N°Obs": "n_obs",
    "6. nb_jours_semaine": "etude_duree",
    "11. D1_modes_transport": "d1_mode",
    "13. D1_vélo_km": "d1_vélo_km",
    "14. D1_VAE_km": "d1_VAE_km",
    "18. D1_voiture_km": "d1_voiture_km",
    "19. D1_voiture_motorisation": "d1_voiture_motorisation",
    "28. D2_déclaration": "d2_déclaration",
    "32. D2_modes_transport": "d2_mode",
    "34. D2_vélo_km": "d2_vélo_km",
    "35. D2_VAE_km": "d2_VAE_km",
    "39. D2_voiture_km": "d2_voiture_km",
    "40. D2_voiture_motorisation": "d2_voiture_motorisation",
    "66. Date de saisie": "date_saisie"
}

df_2024.rename(columns=rename_columns_2024, inplace=True)

# Vérifier que les colonnes sont alignées
assert list(df_2023.columns) == list(df_2024.columns), "Les colonnes ne sont pas alignées après le renommage."

# Fusionner les deux fichiers
df_final = pd.concat([df_2023, df_2024], ignore_index=True)

# Sauvegarder le fichier fusionné
output_file = '/Users/amaloubaba/Downloads/UGA_etudiants_2023_2024_fusionnés.xlsx'
df_final.to_excel(output_file, index=False)

print(f"Fichier fusionné enregistré sous : {output_file}")


# ## UGA - personnel 2023/2024

# Nous allons procéder de la même manière qu'avec les UGA- personnel. Par conséquent, je détaillerai moins les parties ci dessous (car même procédé)

# In[146]:


######## 2023 personnel 
# Charger le fichier Excel contenant les données du personnel 2023
file_path_students = "/Users/amaloubaba/Downloads/UGA-enquete-mobilite-2023-personnel.xlsx"
df_perso_2023 = pd.read_excel(file_path_students)

# Afficher toutes les colonnes disponibles dans le fichier
print("Liste des colonnes dans le fichier :")
for col in df_perso_2023.columns:
    print(col)
    


# In[147]:


# Liste des colonnes à supprimer
colonnes_a_supprimer = [
    '1. Êtes-vous employé par l\'UGA ?',
    '2. Quel est votre statut à l\'UGA ?',
    '3. Autre',
    '6. Distance totale ALLER / RETOUR en marche (km / jour)',
    '9. Distance totale ALLER / RETOUR en trottinette électrique (km / jour)',
    '10. Distance totale ALLER / RETOUR en moto / scooter (km / jour)',
    '11. Nombre de personnes sur le 2 roues motorisées',
    '14. Nombre de personnes dans la voiture',
    '17. Distance totale ALLER / RETOUR en bus (km / jour)',
    '18. Distance totale ALLER / RETOUR en tramway (km / jour)',
    '19. Distance totale ALLER / RETOUR en train (km / jour)',
    '21. Parmi vos [En_moyenne_lorsque_vous_netes_pas_en_vac] jour(s) de déplacements par semaine pour vous rendre sur votre lieu de travail, combien de jours ce deuxième trajet concerne-t-il ?',
    '23. Distance totale ALLER / RETOUR en marche (km / jour)',
    '26. Distance totale ALLER / RETOUR en trottinette électrique (km / jour)',
    '27. Distance totale ALLER / RETOUR en moto / scooter (km / jour)',
    '28. Nombre de personnes sur le 2 roues motorisées',
    '31. Nombre de personnes dans la voiture',
    '34. Distance totale ALLER / RETOUR en bus (km / jour)',
    '35. Distance totale ALLER / RETOUR en tramway (km / jour)',
    '36. Distance totale ALLER / RETOUR en train (km / jour)',
    '37. Souhaitez-vous préciser votre structure d\'appartenance (service, direction, laboratoire...) pour lui permettre d\'utiliser les données renseignées pour réaliser son bilan de gaz à effet de serre ?',
    '38. Votre structure d\'appartenance est :',
    '39. Quel est le nom de votre structure de recherche ?',
    '40. Quel est le nom de votre structure (direction, service...) ?',
    '41. Clé',
    '43. Date de dernière modification',
    '44. Date de dernier enregistrement',
    '45. Temps de saisie',
    '46. Langue',
    '47. Progression',
    '48. Dernière question saisie',
    '49. Origine',
    '50. Appareil utilisé pour la saisie'
]

# Supprimer les colonnes inutiles
df_per_2023_clean = df_perso_2023.drop(columns=colonnes_a_supprimer, errors='ignore')

# Remplacer les en-têtes par la deuxième ligne (actuelle) et supprimer la première ligne
df_per_2023_clean.columns = df_per_2023_clean.iloc[0]
df_pers_cleaned_2023 = df_per_2023_clean[1:].reset_index(drop=True)

df_pers_cleaned_2023


# In[148]:


#vérifions les colonnes
# Afficher toutes les colonnes disponibles dans le fichier
print("Liste des colonnes dans le fichier :")
for col in df_pers_cleaned_2023.columns:
    print(col)


# In[149]:


# Sauvegarde des résultats
output_file_path = "/Users/amaloubaba/Downloads/UGA_personnel_2023__.xlsx"
df_pers_cleaned_2023.to_excel(output_file_path, index=False)


# In[150]:


######## 2024 personnel 
# Charger le fichier Excel contenant les données du personnel 2023
file_path_students = "/Users/amaloubaba/Downloads/_UGA-enquÃªte-mobilitÃ©-2024-personnel.xlsx"
df_personel_2024 = pd.read_excel(file_path_students)

# Afficher toutes les colonnes disponibles dans le fichier
print("Liste des colonnes dans le fichier :")
for col in df_personel_2024.columns:
    print(col)


# In[151]:


df_personel_2024


# In[152]:


# Liste des colonnes à supprimer
colonnes_a_supprimer = [
    '1. statut_agent',
    '2. Autre :',
    '3. site',
    '4. Autre site :',
    '5. type_structure',
    '6. composante',
    '7. Autre :',
    '9. Autre :',
    '11. Autre :',
    '17. D1_marche_km',
    '20. D1_trottinette_km',
    '21. D1_moto_km',
    '22. D1_moto_nb_passagers',
    '25. D1_voiture_nb_passagers',
    '26. covoiturage_intention',
    '27. covoiturage_qui',
    '28. voiture_raisons',
    '29. Autre(s) raison(s) :',
    '30. D1_bus_km',
    '31. D1_tramway_km',
    '32. D1_train_km',
    '34. D2_jours_semaine',
    '39. D2_marche_km',
    '42. D2_trottinette_km',
    '43. D2_moto_km',
    '44. D2_moto_nb_passagers',
    '47. D2_voiture_nb_passagers',
    '48. D2_bus_km',
    '49. D2_tramway_km',
    '50. D2_train_km',
    '52. Clé',
    '54. Date de dernière modification',
    '55. Date de dernier enregistrement',
    '56. Temps de saisie',
    '57. Langue',
    '58. Progression',
    '59. Dernière question saisie',
    '60. Origine',
    '61. Appareil utilisé pour la saisie'
]

# Supprimer les colonnes inutiles
df_per_2024_clean = df_personel_2024.drop(columns=colonnes_a_supprimer, errors='ignore')


df_per_2024_clean


# In[157]:


#vérifions les colonnes
# Afficher toutes les colonnes disponibles dans le fichier
print("Liste des colonnes dans le fichier :")
for col in df_per_2024_clean.columns:
    print(col)


# In[153]:


# Sauvegarde des résultats
output_file_path = "/Users/amaloubaba/Downloads/UGA_personnel_2024__.xlsx"
df_per_2024_clean.to_excel(output_file_path, index=False)


# ## Concaténons les deux fichiers:

# In[159]:


import pandas as pd

# Charger les fichiers Excel
file2_2023 = '/Users/amaloubaba/Downloads/UGA_personnel_2023__.xlsx'
file2_2024 = '/Users/amaloubaba/Downloads/UGA_personnel_2024__.xlsx'

df2_2023 = pd.read_excel(file2_2023)
df2_2024 = pd.read_excel(file2_2024)

# Renommer les colonnes de 2024 pour qu'elles correspondent à celles de 2023
rename_columns_2024 = {
    "N°Obs": "n_obs",
    "12. nb_jours_semaine": "nb_jours_semaine",
    "16. D1_modes_transport": "d1_mode",
    "18. D1_vélo_km": "d1_vélo_km",
    "19. D1_VAE_km": "d1_VAE_km",
    "23. D1_voiture_km": "d1_voiture_km",
    "24. D1_voiture_motorisation": "d1_voiture_motorisation",
    "33. D2_déclaration": "d2_déclaration",
    "38. D2_modes_transport": "d2_mode",
    "40. D2_vélo_km": "d2_vélo_km",
    "41. D2_VAE_km": "d2_VAE_km",
    "45. D2_voiture_km": "d2_voiture_km",
    "46. D2_voiture_motorisation": "d2_voiture_motorisation",
    "53. Date de saisie": "date_saisie"
}

df2_2024.rename(columns=rename_columns_2024, inplace=True)

# Vérifier que les colonnes sont alignées
assert list(df2_2023.columns) == list(df2_2024.columns), "Les colonnes ne sont pas alignées après le renommage."

# Fusionner les deux fichiers
df2_final = pd.concat([df2_2023, df2_2024], ignore_index=True)

# Sauvegarder le fichier fusionné
output_file = '/Users/amaloubaba/Downloads/UGA_personnel_2023_2024_fusionnés.xlsx'
df2_final.to_excel(output_file, index=False)

print(f"Fichier fusionné enregistré sous : {output_file}")


# # Fusion des étudiants et personnel
# 

# vérifions s'ils ont des colonnes similaires pour éviter les erreurs

# In[161]:


#vérifions les colonnes
# Afficher toutes les colonnes disponibles dans le fichier
print("Liste des colonnes dans le fichier étudiant:")
for col in df_final.columns:
    print(col)
    
#vérifions les colonnes
# Afficher toutes les colonnes disponibles dans le fichier
print("Liste des colonnes dans le fichier personnel :")
for col in df2_final.columns:
    print(col)


# Fusionnons ces deux fichier et créons une colonne qui indique si c'est un étudiant -> ETU ou si c'est un personnel -> PER, enfin les données seront triées par la colonne (date de saisie) pour un ordre chronologique 

# In[166]:


import pandas as pd

# Chemins des fichiers
file_etudiants = '/Users/amaloubaba/Downloads/UGA_etudiants_2023_2024_fusionnés.xlsx'
file_personnel = '/Users/amaloubaba/Downloads/UGA_personnel_2023_2024_fusionnés.xlsx'

# Charger les fichiers Excel
df_etudiants = pd.read_excel(file_etudiants)
df_personnel = pd.read_excel(file_personnel)

# Ajouter une colonne 'type_personne' pour indiquer ETU ou PER
df_etudiants['type_personne'] = 'ETU'
df_personnel['type_personne'] = 'PER'

# Fusionner les deux DataFrames sans trier
df_fusion = pd.concat([df_etudiants, df_personnel], ignore_index=True)

# Sauvegarder le résultat dans un nouveau fichier Excel
output_file = '/Users/amaloubaba/Downloads/UGA_2023_2024_ETU_PER__.xlsx'
df_fusion.to_excel(output_file, index=False)

print(f"Fichier fusionné et trié enregistré sous : {output_file}")


# # 2. Dictionnaire des données 

# Passons maintenant directement au Dictionnaire des données, qui inclura la signification des données manipulées, leur format, et leur volumétrie.

# In[168]:


import pandas as pd

# Étape 1 : Charger le fichier Excel
# Remplacez 'file_path' par le chemin vers votre fichier
file_path = '/Users/amaloubaba/Downloads/UGA_2023_2024_ETU_PER__.xlsx'
df = pd.read_excel(file_path)

# Étape 2 : Créer un dictionnaire contenant les informations sur chaque colonne
# On extrait les informations suivantes : 
# - Nom de la colonne
# - Type de données de la colonne (int, float, object, etc.)
# - Quelques exemples de valeurs présentes dans la colonne
# - Le nombre de valeurs non nulles dans la colonne

column_info = pd.DataFrame({
    "Colonne": df.columns,                        # Liste des noms de colonnes
    "Type de Donnée": df.dtypes,                 # Types de données des colonnes
    "Exemples de Valeurs": [                     # On collecte quelques exemples de valeurs pour chaque colonne
        df[col].dropna().unique()[:5]            # Supprime les valeurs nulles et prend les 5 premières valeurs uniques
        for col in df.columns
    ],
    "Nombre de valeurs non nulles": df.notnull().sum()  # Compte le nombre de valeurs non nulles pour chaque colonne
})

# Étape 3 : Calculer la volumétrie globale du fichier
# On collecte :
# - Le nombre total de lignes (nombre d'observations)
# - Le nombre total de colonnes (variables)

volumetrie = {
    "Nombre total de lignes": len(df),           # Compte le nombre total de lignes dans le fichier
    "Nombre total de colonnes": len(df.columns)  # Compte le nombre total de colonnes dans le fichier
}

# Étape 4 : Afficher ou sauvegarder les résultats
# Affiche le dictionnaire des données (vous pouvez aussi l'exporter vers un fichier si nécessaire)
print("Dictionnaire des données :")
print(column_info)

# Affiche la volumétrie globale
print("\nVolumétrie globale :")
print(volumetrie)

# Facultatif : sauvegarder le dictionnaire des données dans un fichier Excel pour consultation
output_path = '/Users/amaloubaba/Downloads/Dictionnaire_des_données.xlsx'
column_info.to_excel(output_path, index=False)  # Sauvegarde le dictionnaire des données dans un fichier Excel
print(f"\nDictionnaire des données sauvegardé sous : {output_path}")


# In[172]:


import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Charger les données
file_path = '/Users/amaloubaba/Downloads/UGA_2023_2024_ETU_PER__.xlsx'
df = pd.read_excel(file_path)

# Aperçu des données
apercu = df.head()

# Statistiques descriptives pour colonnes numériques
stats_numeriques = df.describe()

# Statistiques descriptives pour colonnes non numériques
stats_categoriques = df.describe(include=['object'])

# Distribution des colonnes numériques
numerical_columns = df.select_dtypes(include=['float64', 'int64']).columns
distribution_numeriques = {}
for col in numerical_columns:
    plt.figure()
    sns.histplot(df[col], kde=True, bins=30)
    plt.title(f"Distribution de la colonne : {col}")
    plt.xlabel(col)
    plt.ylabel("Fréquence")
    plt.show()

# Répartition des colonnes catégoriques (si présentes)
categorical_columns = df.select_dtypes(include=['object']).columns
for col in categorical_columns:
    plt.figure()
    sns.countplot(data=df, y=col, order=df[col].value_counts().index)
    plt.title(f"Répartition des catégories pour la colonne : {col}")
    plt.xlabel("Nombre")
    plt.ylabel(col)
    plt.show()

# Étude d'impact sur les objectifs métier (exemple avec d1_mode si présent)
if 'd1_mode' in df.columns:
    plt.figure(figsize=(10, 6))
    sns.countplot(data=df, y='d1_mode', order=df['d1_mode'].value_counts().index)
    plt.title("Répartition des modes de transport (d1_mode)")
    plt.xlabel("Nombre d'observations")
    plt.ylabel("Modes de transport")
    plt.show()

# Résultat des valeurs manquantes
valeurs_nulles = df.isnull().sum()



# ## Analyons ces données 

# ### Visualisation des données

# In[174]:


# Visualiser un aperçu des données (5 premières lignes)
print("Aperçu des données :")
print(df.head())

# Afficher les informations générales sur le DataFrame
print("\nInformations générales sur les données :")
print(df.info())

# Répartition des valeurs manquantes par colonne
missing_values = df.isnull().sum()
print("\nNombre de valeurs manquantes par colonne :")
print(missing_values)


# Commentaires:
# Les données sont bien structurées avec des colonnes variées (numériques, catégoriques, dates). Cependant, certaines colonnes comme d1_vélo_km et d2_vélo_km présentent un taux important de valeurs manquantes (> 70 %), ce qui nécessite un plan de gestion. L’ensemble des données est cohérent en surface, prêt pour des analyses détaillées.
# 
# 

# ### Statistiques descriptives 

# In[175]:


# Statistiques descriptives
print("Statistiques descriptives :")
print(df.describe())

# Valeurs aberrantes
aberrant_values = df[df['d1_voiture_km'] > 100000]
print("\nValeurs aberrantes dans 'd1_voiture_km' :")
print(aberrant_values[['n_obs', 'd1_voiture_km']])


# Les statistiques descriptives révèlent des tendances attendues pour la majorité des colonnes numériques, comme une médiane cohérente pour les distances parcourues (d1_voiture_km : 22 km). Cependant, des valeurs aberrantes, telles qu’un maximum de 300,000 km, doivent être investiguées et potentiellement corrigées pour garantir des analyses fiables.

# ### Étude de l'impact des modes de transport sur les distances parcourues

# In[177]:


# Étudier la distance moyenne parcourue par mode de transport (d1_mode et d1_vélo_km)
if 'd1_mode' in df.columns and 'd1_vélo_km' in df.columns:
    avg_distance_by_mode = df.groupby('d1_mode')['d1_vélo_km'].mean().sort_values(ascending=False)
    print("\nDistances moyennes parcourues par mode de transport (en km) :")
    print(avg_distance_by_mode)




# Les données montrent que certains modes de transport, notamment ceux incluant des vélos ou des véhicules individuels, affichent des distances moyennes significativement plus élevées. Cette information peut orienter les politiques de mobilité, comme la promotion des modes de transport à faible impact environnemental pour les trajets plus courts. Les modes de transport combinés offrent également des opportunités pour une approche multimodale.

# ## Retraitement des dates

# In[178]:


# Étape : Retraitement des dates (extraire uniquement l'année)
if 'date_saisie' in df.columns:
    # Convertir les dates en format datetime si ce n'est pas déjà fait
    df['date_saisie'] = pd.to_datetime(df['date_saisie'], errors='coerce')

    # Extraire uniquement l'année dans une nouvelle colonne
    df['année_saisie'] = df['date_saisie'].dt.year

    # Vérifier les données transformées
    print("Années extraites :")
    print(df['année_saisie'].value_counts())
    
    # Afficher les premières lignes des colonnes concernées pour valider la transformation
    print("\nAperçu des colonnes après transformation :")
    print(df[['date_saisie', 'année_saisie']].head())


# ## Analysons la qualité des données

# In[179]:


# Étape : Analyser la qualité des données (valeurs manquantes et erronées)

# Calculer le pourcentage de valeurs manquantes pour chaque colonne
missing_values_percentage = (df.isnull().sum() / len(df)) * 100
print("Pourcentage de valeurs manquantes par colonne :")
print(missing_values_percentage)

# Vérifier les valeurs aberrantes dans des colonnes spécifiques
# Exemple : détecter les distances aberrantes dans 'd1_vélo_km' et 'd1_voiture_km'
if 'd1_vélo_km' in df.columns:
    velo_aberrant = df[df['d1_vélo_km'] > 200]  # Plus de 200 km à vélo en une journée
    print("\nValeurs aberrantes détectées dans 'd1_vélo_km' (plus de 200 km) :")
    print(velo_aberrant[['n_obs', 'd1_vélo_km']])

if 'd1_voiture_km' in df.columns:
    voiture_aberrant = df[df['d1_voiture_km'] > 1000]  # Plus de 1,000 km en voiture en une journée
    print("\nValeurs aberrantes détectées dans 'd1_voiture_km' (plus de 1,000 km) :")
    print(voiture_aberrant[['n_obs', 'd1_voiture_km']])


# Les données présentent des problèmes de qualité. Plusieurs colonnes comme d1_vélo_km (71 % de valeurs manquantes) et d1_voiture_km (73 % de valeurs manquantes) nécessitent un traitement. Des anomalies sont détectées, comme 1,200 km pour d1_vélo_km et 300,000 km pour d1_voiture_km, qui doivent être corrigées pour garantir des analyses fiables

# # 2. PréparaIon des données

# ## 1. Sélectionner les données prises en compte

# In[185]:


# Étape : Préparation des données - Sélection des données prises en compte

# Filtrer les données pour ne garder que les modes de transport "vélo", "vélo électrique", et "voiture"
selected_modes = ['Vélo', 'Vélo électrique', 'Voiture']

# Filtrer les données pour D1_mode ou D2_mode
filtered_df = df[
    (df['d1_mode'].isin(selected_modes)) | 
    (df['d2_mode'].isin(selected_modes))
]

# Afficher un aperçu des données filtrées
print("Aperçu des données après filtrage des modes de transport sélectionnés :")
print(filtered_df.head())

# Vérifier la répartition des modes de transport après filtrage
print("\nRépartition des modes de transport dans 'd1_mode' et 'd2_mode' :")
print("D1_mode :")
print(filtered_df['d1_mode'].value_counts())
print("\nD2_mode :")
print(filtered_df['d2_mode'].value_counts())




# Ce code filtre les données pour ne conserver que les lignes où le mode de transport correspond à Vélo, Vélo électrique, ou Voiture. 

# In[186]:


df_filtered


# ### Explication de la méthode isin et de notre logique de filtrage
# 
# La méthode utilisée avec isin est une approche flexible qui permet de conserver un maximum de données pertinentes tout en respectant nos critères de sélection. Au lieu de filtrer strictement les lignes qui contiennent uniquement les modes "Vélo", "Vélo électrique" ou "Voiture" dans les deux colonnes d1_mode et d2_mode, cette méthode vérifie si au moins une des deux colonnes contient un des modes sélectionnés. Si par exemple une ligne contient "Vélo" dans d1_mode et "Tramway" dans d2_mode, elle sera retenue car d1_mode correspond à un mode pertinent. Cette logique garantit que nous ne perdons pas de données utiles, même si certains modes non pertinents (comme "Tramway") apparaissent dans une des colonnes.
# 
# 
# Ce choix est motivé par le fait que les modes non pertinents, comme "Tramway" ou "Marche", n’auront pas d’impact direct sur nos calculs. Nous pouvons ainsi nous concentrer uniquement sur les modes pertinents pour l’analyse (consommation et émissions de CO2) tout en maximisant la quantité de données disponibles. Cette approche permet une exploitation plus complète des données collectées et réduit le risque de biais lié à l’élimination excessive de lignes, ce qui est particulièrement important dans les analyses quantitatives ou les calculs de KPI.
# 
# 

# ## 2. Gestion des valeurs aberrantes et suppression des données de 2025

# In[188]:


# Nettoyage des données : Gestion des valeurs aberrantes et suppression des données de 2025

# Étape 1 : Gestion des valeurs aberrantes
# Fixer des seuils raisonnables pour les distances (exemple : vélo < 200 km, voiture < 1000 km)
df_cleaned = df_filtered[(df_filtered['d1_vélo_km'] <= 200) | (df_filtered['d1_voiture_km'] <= 1000)]

# Étape 2 : Suppression des données de 2025
if 'année_saisie' in df_cleaned.columns:
    df_cleaned = df_cleaned[df_cleaned['année_saisie'] < 2025]

# Vérifier les résultats après nettoyage
print("\nAperçu des données après nettoyage (valeurs aberrantes et suppression de 2025) :")
print(df_cleaned.head())

# Vérifier les dimensions après nettoyage
print("\nDimensions du dataframe après nettoyage :")
print(df_cleaned.shape)

# Vérification des années restantes
print("\nAnnées présentes dans le dataframe après nettoyage :")
print(df_cleaned['année_saisie'].value_counts())


# Après nettoyage, le jeu de données contient 4,497 lignes et 17 colonnes. Les valeurs aberrantes ont été traitées en fixant des seuils raisonnables pour les distances : moins de 200 km pour les déplacements à vélo et moins de 1,000 km pour les déplacements en voiture. Les observations correspondant à l'année 2025 ont également été supprimées.
# 
# Le jeu de données nettoyé conserve deux années principales : 2,575 observations pour 2024 et 1,922 pour 2023. Cela garantit une meilleure cohérence des données, en excluant les anomalies et les années hors du périmètre d'étude. L'aperçu montre que seules les informations pertinentes et conformes aux critères ont été retenues, préparant ainsi le jeu de données pour l'analyse des KPI liés à la consommation et aux émissions de CO2.

# # Calcul des KPI

# ## Consommation

# In[192]:


# Charger le fichier contenant les données de consommation énergétique
file_path_energy = "/Users/amaloubaba/Downloads/Consommation.xlsx"
df_energy = pd.read_excel(file_path_energy)

# Nettoyer les données de consommation pour extraire les valeurs numériques
df_energy['Consommation'] = df_energy['Consommation (kWh/km)'].str.extract(r'([\d\.]+)').astype(float)

# Dictionnaire des consommations par type de véhicule
consumption_dict = {
    'vélo électrique': df_energy.loc[df_energy['Type de véhicule'].str.lower() == 'vélo électrique', 'Consommation'].values[0],
    'vélo': 0,  # Pas de consommation pour les vélos normaux
    'voiture essence': df_energy.loc[df_energy['Type de véhicule'].str.lower() == 'voiture essence', 'Consommation'].values[0],
    'voiture diesel': df_energy.loc[df_energy['Type de véhicule'].str.lower() == 'voiture diesel', 'Consommation'].values[0]
}

# Nettoyer la colonne "Motorisation de la voiture" pour éviter les incohérences
df_cleaned['d1_voiture_motorisation'] = df_cleaned['d1_voiture_motorisation'].fillna('inconnu').astype(str).str.lower().str.strip()

# Fonction pour calculer la consommation énergétique
def calculate_energy_consumption(row):
    energy_kwh = 0
    fuel_liters = 0

    # Vélo électrique (kWh/km)
    if row['d1_mode'] == 'Vélo électrique' and pd.notnull(row['d1_vélo_km']):
        energy_kwh += row['d1_vélo_km'] * consumption_dict['vélo électrique']
    
    # Vélo normal (0 consommation énergétique)
    if row['d1_mode'] == 'Vélo' and pd.notnull(row['d1_vélo_km']):
        energy_kwh += 0  # Valeur explicitement ajoutée

    # Voiture (L/km)
    if row['d1_mode'] == 'Voiture' and pd.notnull(row['d1_voiture_km']):
        motorization = row['d1_voiture_motorisation']
        if 'essence' in motorization:
            fuel_liters += row['d1_voiture_km'] * consumption_dict['voiture essence']
        elif 'diesel' in motorization:
            fuel_liters += row['d1_voiture_km'] * consumption_dict['voiture diesel']

    return pd.Series([energy_kwh, fuel_liters])

# Appliquer la fonction au DataFrame
df_cleaned[['Consommation_énergétique_totale (kWh)', 'Consommation_carburant_totale (L)']] = df_cleaned.apply(calculate_energy_consumption, axis=1)

# Ajouter une colonne pour le vélo normal (consommation toujours 0 pour clarté)
df_cleaned['Consommation_pour_vélo (kWh)'] = df_cleaned.apply(
    lambda row: row['d1_vélo_km'] * consumption_dict['vélo'] if row['d1_mode'] == 'Vélo' else 0, axis=1
)

# Vérification des résultats
print(df_cleaned[['d1_mode', 'd1_vélo_km', 'd1_voiture_km', 'd1_voiture_motorisation',
                  'Consommation_énergétique_totale (kWh)', 'Consommation_carburant_totale (L)', 
                  'Consommation_pour_vélo (kWh)']].head())


# Sauvegarde des résultats dans un fichier
output_file_path = "/Users/amaloubaba/Downloads/UGA_personnel_avec_conso.xlsx"
df_cleaned.to_excel(output_file_path, index=False)
print(f"Fichier avec consommation énergétique sauvegardé : {output_file_path}")


# In[194]:


df_cleaned


# In[196]:


# Calcul des totaux pour les consommations énergétiques et de carburant

# Total de la consommation énergétique (kWh)
total_energy_kwh = df_cleaned['Consommation_énergétique_totale (kWh)'].sum()

# Total de la consommation de carburant (L)
total_fuel_liters = df_cleaned['Consommation_carburant_totale (L)'].sum()

# Total de la consommation pour les vélos (kWh)
total_bike_energy = df_cleaned['Consommation_pour_vélo (kWh)'].sum()

# Affichage des résultats
print("Total de la consommation énergétique (kWh) :", total_energy_kwh)
print("Total de la consommation de carburant (L) :", total_fuel_liters)
print("Total de la consommation pour les vélos (kWh) :", total_bike_energy)


# Analyse des résultats :
# 
# Total de la consommation énergétique (kWh) : 9.625 kWh
# Cette consommation provient exclusivement des vélos électriques. Elle est relativement faible, ce qui est cohérent avec l'efficacité énergétique des vélos électriques (faible consommation en kWh/km).
# 
# 
# Total de la consommation de carburant (L) : 4,277.068 L
# Cette consommation provient des voitures, incluant celles à motorisation essence et diesel.
# Ce chiffre est élevé, reflétant l'importance des trajets en voiture dans les données. Les émissions associées à cette consommation seront significatives en raison des facteurs d'émission élevés pour les carburants fossiles.
# 
# 
# Total de la consommation pour les vélos (kWh) : 0.0 kWh
# Cela concerne les vélos normaux, qui n'ont aucune consommation énergétique, ce qui est conforme à leur absence de motorisation.
# 
# 
# 
# ----> Ces résultats soulignent l'importance de promouvoir des alternatives comme les vélos électriques ou normaux pour réduire la dépendance énergétique et les émissions liées aux voitures. Une analyse complémentaire pourrait inclure le calcul des émissions de CO2 totales associées à ces consommations.

# ## Émissions CO2

# In[201]:


import pandas as pd

# Charger le fichier contenant les données d'émissions de CO2
file_path_emissions = "/Users/amaloubaba/Downloads/emissions.xlsx"
df_emissions = pd.read_excel(file_path_emissions)

# Nettoyer la colonne des émissions CO2 pour supprimer les unités et convertir en float
df_emissions['Émissions CO₂ (g/km)'] = df_emissions['Émissions CO₂ (g/km)'].astype(str).str.replace(' g/km', '', regex=False).str.replace(',', '.', regex=False).astype(float)

# Création d'un dictionnaire des émissions de CO₂ par mode de transport
co2_emission_factors = {
    row['Type de transport'].strip().lower(): row['Émissions CO₂ (g/km)'] / 1000  # Conversion en kg/km
    for _, row in df_emissions.iterrows()
}

# Remplacer la valeur NaN pour le vélo normal (aucune émission)
co2_emission_factors['vélo normal'] = 0.0

# Vérifier le dictionnaire
print("\nFacteurs d'émission de CO2 (kg/km) :", co2_emission_factors)

# Nettoyer la colonne "Motorisation de la voiture"
df_cleaned['d1_voiture_motorisation'] = df_cleaned['d1_voiture_motorisation'].astype(str).str.lower().str.strip().fillna('inconnu')
df_cleaned['d2_voiture_motorisation'] = df_cleaned['d2_voiture_motorisation'].astype(str).str.lower().str.strip().fillna('inconnu')

# Fonction pour calculer les émissions de CO2 pour chaque mode de transport
def calculate_co2_emissions(row):
    # Emissions pour D1_mode
    if row['d1_mode'] == 'Vélo électrique':
        co2_d1 = row['d1_vélo_km'] * co2_emission_factors['vélo électrique']
    elif row['d1_mode'] == 'Voiture' and 'essence' in row['d1_voiture_motorisation']:
        co2_d1 = row['d1_voiture_km'] * co2_emission_factors['voiture essence']
    elif row['d1_mode'] == 'Voiture' and ('diesel' in row['d1_voiture_motorisation'] or 'diésel' in row['d1_voiture_motorisation']):
        co2_d1 = row['d1_voiture_km'] * co2_emission_factors['voiture diésel']
    else:
        co2_d1 = 0.0

    # Emissions pour D2_mode
    if row['d2_mode'] == 'Vélo électrique':
        co2_d2 = row['d2_vélo_km'] * co2_emission_factors['vélo électrique']
    elif row['d2_mode'] == 'Voiture' and 'essence' in row['d2_voiture_motorisation']:
        co2_d2 = row['d2_voiture_km'] * co2_emission_factors['voiture essence']
    elif row['d2_mode'] == 'Voiture' and ('diesel' in row['d2_voiture_motorisation'] or 'diésel' in row['d2_voiture_motorisation']):
        co2_d2 = row['d2_voiture_km'] * co2_emission_factors['voiture diésel']
    else:
        co2_d2 = 0.0

    return co2_d1 + co2_d2

# Appliquer la fonction pour calculer les émissions totales
df_cleaned['Émissions_CO2_totales (kg)'] = df_cleaned.apply(calculate_co2_emissions, axis=1)

# Vérifier les résultats
print("\nAperçu des données avec émissions calculées :")
print(df_cleaned[['d1_mode', 'd1_voiture_km', 'd1_voiture_motorisation', 'd2_mode', 'd2_voiture_km', 'Émissions_CO2_totales (kg)']].head())

# Sauvegarder les résultats
output_file_path = "/Users/amaloubaba/Downloads/UGA_personnel_avec_co2.xlsx"
df_cleaned.to_excel(output_file_path, index=False)

print(f"Fichier sauvegardé à l'emplacement : {output_file_path}")


# In[202]:


# Ajouter une colonne pour les émissions de CO2 en tonnes
df_cleaned['Émissions_CO2_totales (tonnes)'] = df_cleaned['Émissions_CO2_totales (kg)'] / 1000

# Ajouter une colonne pour le nombre d'arbres équivalents
# Hypothèse : 1 arbre capture environ 50 kg de CO2 en 1,5 an
df_cleaned['Arbres_équivalents'] = df_cleaned['Émissions_CO2_totales (kg)'] / (50 / 1.5)

# Vérifier les résultats
print("\nAperçu des données avec émissions en tonnes et arbres équivalents :")
print(df_cleaned[['Émissions_CO2_totales (kg)', 'Émissions_CO2_totales (tonnes)', 'Arbres_équivalents']].head())

# Sauvegarder le fichier enrichi
output_file_path_enriched = "/Users/amaloubaba/Downloads/UGA_personnel_avec_co2_et_arbres.xlsx"
df_cleaned.to_excel(output_file_path_enriched, index=False)

print(f"Fichier enrichi sauvegardé à l'emplacement : {output_file_path_enriched}")



# In[203]:


df_cleaned


# Les résultats montrent que les émissions de CO₂ sont nulles pour les vélos normaux, comme attendu, et très faibles pour les vélos électriques (exprimées en kg). En revanche, les voitures (essence ou diesel) génèrent des émissions significatives. Par exemple, une ligne indique une émission totale de 1,237 kg (ou 0,001237 tonnes), nécessitant environ 0,037 arbres pour être compensée. Ces données mettent en évidence l'impact environnemental des modes de transport motorisés par rapport aux alternatives plus durables.

# ## Calories dépensées

# In[204]:


# Définition des calories brûlées par heure pour chaque type de transport
calories_per_hour = {
    'vélo électrique': 246,  # calories par heure
    'vélo': 354,             # calories par heure (vélo normal)
    'voiture essence': 0,    # calories par heure
    'voiture diesel': 0      # calories par heure
}

# Supposons des vitesses moyennes pour convertir la distance en temps (heures)
vitesse_velo_normal = 15  # km/h pour vélo normal
vitesse_velo_electrique = 20  # km/h pour vélo électrique
vitesse_voiture = 50  # km/h pour les voitures (valeur estimée)

# Conversion des distances en temps (en heures)
df_cleaned['Temps vélo (h)'] = df_cleaned['d1_vélo_km'] / vitesse_velo_normal
df_cleaned['Temps vélo électrique (h)'] = df_cleaned['d1_vélo_km'] / vitesse_velo_electrique
df_cleaned['Temps voiture (h)'] = df_cleaned['d1_voiture_km'] / vitesse_voiture

# Calcul des calories dépensées pour chaque type de transport
df_cleaned['Calories vélo'] = df_cleaned['Temps vélo (h)'] * calories_per_hour['vélo']
df_cleaned['Calories vélo électrique'] = df_cleaned['Temps vélo électrique (h)'] * calories_per_hour['vélo électrique']
df_cleaned['Calories voiture essence'] = df_cleaned['Temps voiture (h)'] * calories_per_hour['voiture essence']
df_cleaned['Calories voiture diesel'] = df_cleaned['Temps voiture (h)'] * calories_per_hour['voiture diesel']

# Ajouter une colonne "Calories dépensées" en sommant toutes les valeurs
df_cleaned['Calories dépensées'] = (
    df_cleaned['Calories vélo'].fillna(0) +
    df_cleaned['Calories vélo électrique'].fillna(0) +
    df_cleaned['Calories voiture essence'].fillna(0) +
    df_cleaned['Calories voiture diesel'].fillna(0)
)

# Vérification des résultats
print("\nAperçu des données avec calories calculées :")
print(df_cleaned[['d1_vélo_km', 'd1_voiture_km', 'Temps vélo (h)', 'Temps vélo électrique (h)',
                  'Temps voiture (h)', 'Calories vélo', 'Calories vélo électrique', 
                  'Calories voiture essence', 'Calories voiture diesel', 
                  'Calories dépensées']].head())

# Sauvegarde des résultats
output_file_path_calories = "/Users/amaloubaba/Downloads/UGA_personnel_avec_calories.xlsx"
df_cleaned.to_excel(output_file_path_calories, index=False)
print(f"Fichier mis à jour avec les calories dépensées : {output_file_path_calories}")


# Les résultats montrent que le vélo normal est le mode de transport le plus exigeant sur le plan physique, avec une dépense calorique notable (ex. 188,8 calories pour 8 km). Le vélo électrique, bien que moins intense (98,4 calories pour la même distance), reste bénéfique. En revanche, les voitures (essence ou diesel) n'entraînent aucune dépense calorique, soulignant leur caractère passif. Ces chiffres mettent en avant l'impact positif des modes actifs sur la santé.

# In[205]:


# Sauvegarde des résultats finaux
output_file_path_calories = "/Users/amaloubaba/Downloads/UGA_2023_2024_FINAL.xlsx"
df_cleaned.to_excel(output_file_path_calories, index=False)
print(f"Fichier final : {output_file_path_calories}")


# In[ ]:




