import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import math
import openpyxl

# Création d'une base de données fictive
np.random.seed(42) #reproductibilité 

# Données pour les actifs (crédits)
actifs = pd.DataFrame({
    'Type': 'Crédit',
    'Montant': np.round(np.random.uniform(10000, 500000, 500)),
    'Taux': np.random.choice([0.03,0.04, 0.05], 500),
    'Maturité': np.random.choice([i for i in range(1,21)], 500)
})

# Données pour les passifs (livret A, DAV)
passifs = pd.DataFrame({
    'Type': np.random.choice(['Livret A', 'DAV'], 500),
    'Montant': np.round(np.random.uniform(10000, 350000, 500)),
    'Taux': np.full(500, 0.02),  # Taux fixe pour Livret A et DAV
    'Maturité': np.random.choice([i for i in range(1,21)], 500)
})

# Fusionnement des actifs et passifs dans une seule base de données
base_de_donnees = pd.concat([actifs, passifs]).reset_index(drop=True)

base_de_donnees.to_excel('base_de_donnees.xlsx', index=False)

# Fonction helper pour la normalisation des axes
def normaliser_ordonnees(ax, valeurs):
    max_val = np.max(np.abs(valeurs))
    if max_val >= 1e9:
        ax.set_ylabel('Milliards')
        ax.set_yticklabels(['{:.1f}'.format(y / 1e9) for y in ax.get_yticks()])
    elif max_val >= 1e6:
        ax.set_ylabel('Millions')
        ax.set_yticklabels(['{:.1f}'.format(y / 1e6) for y in ax.get_yticks()])
    else:
        ax.set_ylabel('Valeur')

def calculer_et_tracer_gap_de_taux(base_de_donnees):
    # Durée de l'analyse
    annee_max = base_de_donnees['Maturité'].max()
    annees = np.arange(1, annee_max + 1)  # De 1 à l'année max

    # Initialisation des flux
    flux_credits = np.zeros_like(annees, dtype=float)
    flux_passifs = np.zeros_like(annees, dtype=float)

    # Calcul des flux pour les actifs (crédits)
    for _, actif in base_de_donnees[base_de_donnees['Type'] == 'Crédit'].iterrows():
        maturite = actif['Maturité']
        taux = actif['Taux']
        montant = actif['Montant']
        for annee in annees:
            if annee <= maturite:
                flux_credits[annee-1] += montant * (1 + taux)**annee - montant

    # Calcul des flux pour les passifs (Livret A, DAV)
    for _, passif in base_de_donnees[base_de_donnees['Type'].isin(['Livret A', 'DAV'])].iterrows():
        maturite = passif['Maturité']
        taux = passif['Taux']
        montant = passif['Montant']
        for annee in annees:
            if annee <= maturite:
                flux_passifs[annee-1] -= montant * (1 + taux)**annee - montant

    # Calcul du gap de taux annuel
    gap_de_taux_annuel = flux_passifs + flux_credits

    # Traçage du graphique du gap de taux
    plt.figure(figsize=(20, 5))
    plt.bar(annees, gap_de_taux_annuel, color='orange', label='Gap de taux')
    plt.xlabel('Années')
    plt.ylabel('Gap de taux')
    plt.title('Gap de taux annuel pour plusieurs actifs et passifs')
    plt.axhline(0, color='black', linewidth=0.8)
    plt.xticks(annees)
    plt.legend()
    plt.grid(axis='y')
    ax = plt.gca()
    normaliser_ordonnees(ax, gap_de_taux_annuel)
    plt.show()

    return gap_de_taux_annuel

def calculer_et_tracer_gap_avec_couverture_neutre(base_de_donnees):
    annee_max = base_de_donnees['Maturité'].max()
    annees = np.arange(1, annee_max + 1)

    flux_credits = np.zeros_like(annees, dtype=float)
    flux_passifs = np.zeros_like(annees, dtype=float)

    for _, actif in base_de_donnees[base_de_donnees['Type'] == 'Crédit'].iterrows():
        for annee in annees:
            if annee <= actif['Maturité']:
                flux_credits[annee - 1] += actif['Montant'] * (1 + actif['Taux']) ** annee - actif['Montant']

    for _, passif in base_de_donnees[base_de_donnees['Type'].isin(['Livret A', 'DAV'])].iterrows():
        for annee in annees:
            if annee <= passif['Maturité']:
                flux_passifs[annee - 1] -= passif['Montant'] * (1 + passif['Taux']) ** annee - passif['Montant']

    gap_de_taux_avant_couverture = flux_passifs + flux_credits
    swaps_neutres = -gap_de_taux_avant_couverture
    gap_de_taux_apres_couverture = gap_de_taux_avant_couverture + swaps_neutres

    plt.figure(figsize=(20, 5))
    plt.bar(annees - 0.2, gap_de_taux_avant_couverture, width=0.4, color='green', label='Gap avant couverture')
    plt.bar(annees - 0.2, swaps_neutres, width=0.4, color='orange', label='Swap emprunteur taux fixe')
    plt.bar(annees + 0.2, gap_de_taux_apres_couverture, width=0.4, color='blue', label='Gap après couverture neutre')
    plt.xlabel('Années')
    plt.title('Gap de taux avant et après couverture neutre')
    plt.axhline(0, color='black', linewidth=0.8)
    plt.xticks(annees)
    plt.legend()
    plt.grid(axis='y')

    ax = plt.gca()
    normaliser_ordonnees(ax, np.concatenate([gap_de_taux_avant_couverture, gap_de_taux_apres_couverture]))
    plt.show()

    return gap_de_taux_avant_couverture, gap_de_taux_apres_couverture, swaps_neutres


def impact_mni(base_de_donnees, variation_taux):
    gap_de_taux_annuel = calculer_et_tracer_gap_de_taux(base_de_donnees)
    annee_max = base_de_donnees['Maturité'].max()
    annees = np.arange(1, annee_max + 1)

    impact_mni_avant_couverture = gap_de_taux_annuel * variation_taux / 100

    plt.figure(figsize=(20, 5))
    plt.bar(annees, impact_mni_avant_couverture, color='red', label='Impact sur MNI')
    plt.xlabel('Années')
    plt.ylabel('Impact sur MNI (%)')
    plt.title('Impact de la variation de taux sur la MNI')
    plt.axhline(0, color='black', linewidth=0.8)
    plt.xticks(annees)
    plt.legend()
    plt.grid(axis='y')

    ax = plt.gca()
    normaliser_ordonnees(ax, impact_mni_avant_couverture)
    plt.show()

#Impact de choc sur la VAN et sur le Bilan
    
def choc_parallele(taux, maturite, delta):
    return taux + delta

def choc_pentification(taux, maturite, delta_court, delta_long):
    return np.where(maturite <= 1, taux + delta_court, taux + delta_long)

def choc_aplatissement(taux, maturite, delta_court, delta_long):
    return np.where(maturite > 1, taux + delta_long, taux - delta_court)

def choc_hausse_taux_courts(taux, maturite, delta):
    return np.where(maturite <= 1, taux + delta, taux)

def choc_baisse_taux_courts(taux, maturite, delta):
    return np.where(maturite <= 1, taux - delta, taux)

def calculer_cash_flows(montants, taux, maturites):
    return montants / ((1 + taux) ** maturites)

# Cette fonction applique le choc sur les taux et calcule la VAN
def appliquer_choc_et_calculer_van(base_de_donnees, fonction_choc, *args):
    taux_ajustes = fonction_choc(base_de_donnees['Taux'], base_de_donnees['Maturité'], *args)
    base_modifiee = base_de_donnees.assign(Taux=taux_ajustes)
    van = sum(calculer_cash_flows(base_modifiee['Montant'], base_modifiee['Taux'], base_modifiee['Maturité']))
    return van

# Fonction principale pour calculer la VAN par choc
def calculer_van_par_choc(base_de_donnees, choc_haut, choc_bas, delta_court, delta_long):
    van_par_choc = {
        'Déplacement parallèle vers le haut': appliquer_choc_et_calculer_van(base_de_donnees, choc_parallele, choc_haut),
        'Déplacement parallèle vers le bas': appliquer_choc_et_calculer_van(base_de_donnees, choc_parallele, choc_bas),
        'Pentification de la courbe': appliquer_choc_et_calculer_van(base_de_donnees, choc_pentification, delta_court, delta_long),
        'Aplatissement de la courbe': appliquer_choc_et_calculer_van(base_de_donnees, choc_aplatissement, delta_court, delta_long),
        'Hausse des taux courts': appliquer_choc_et_calculer_van(base_de_donnees, choc_hausse_taux_courts, delta_court),
        'Baisse des taux courts': appliquer_choc_et_calculer_van(base_de_donnees, choc_baisse_taux_courts, delta_court)
    }
    
    return van_par_choc

def calculer_impact_sur_bilan(base_de_donnees, choc_haut, choc_bas, delta_court, delta_long):
    # Calcul de la VAN de base sans choc
    van_base = sum(calculer_cash_flows(base_de_donnees['Montant'], base_de_donnees['Taux'], base_de_donnees['Maturité']))

    # Utilisation de la fonction précédente pour obtenir la VAN par choc
    van_par_choc = calculer_van_par_choc(base_de_donnees, choc_haut, choc_bas, delta_court, delta_long)

    # Calcul de l'impact sur le bilan pour chaque scénario de choc
    impacts_sur_bilan = {scenario: van - van_base for scenario, van in van_par_choc.items()}

    return impacts_sur_bilan


#Calcul de RTIG

def calculer_sensibilite(base_de_donnees, delta_taux):
    """Calcul simplifié de la sensibilité des actifs/passifs."""
    sensibilites = -base_de_donnees['Montant'] * base_de_donnees['Maturité'] * delta_taux / (1 + base_de_donnees['Taux'] * base_de_donnees['Maturité'])
    return sensibilites

def calculer_cash_flows(montant, taux, maturite):
    return montant / ((1 + taux) ** maturite)

def appliquer_choc_taux(taux, choc):
    return taux + choc

def calculer_var_99(base_de_donnees, delta_taux):
    # Calcul de la sensibilité pour un delta de taux
    base_de_donnees['Sensibilité'] = calculer_sensibilite(base_de_donnees, delta_taux)

    # Application  d'un choc de hausse et baisse de taux
    base_de_donnees['Taux_Apres_Choc_Hausse'] = appliquer_choc_taux(base_de_donnees['Taux'], delta_taux)
    base_de_donnees['Taux_Apres_Choc_Baisse'] = appliquer_choc_taux(base_de_donnees['Taux'], -delta_taux)

    # Calcul des cash-flows actualisés après les chocs
    base_de_donnees['Cash_Flows_Central'] = base_de_donnees.apply(
        lambda x: calculer_cash_flows(x['Montant'], x['Taux'], x['Maturité']), axis=1
    )
    base_de_donnees['Cash_Flows_Apres_Choc_Hausse'] = base_de_donnees.apply(
        lambda x: calculer_cash_flows(x['Montant'], x['Taux_Apres_Choc_Hausse'], x['Maturité']), axis=1
    )
    base_de_donnees['Cash_Flows_Apres_Choc_Baisse'] = base_de_donnees.apply(
        lambda x: calculer_cash_flows(x['Montant'], x['Taux_Apres_Choc_Baisse'], x['Maturité']), axis=1
    )

    # Calcul des ΔVAN pour une hausse et une baisse de taux
    base_de_donnees['Delta_VAN_Hausse'] = base_de_donnees['Cash_Flows_Apres_Choc_Hausse'] - base_de_donnees['Cash_Flows_Central']
    base_de_donnees['Delta_VAN_Baisse'] = base_de_donnees['Cash_Flows_Apres_Choc_Baisse'] - base_de_donnees['Cash_Flows_Central']

    # Calcul de la VaR à 99% pour les ΔVAN
    var_99_hausse = np.percentile(base_de_donnees['Delta_VAN_Hausse'], 1)
    var_99_baisse = np.percentile(base_de_donnees['Delta_VAN_Baisse'], 1)
    besoin_en_fonds_propres = max(-var_99_hausse, -var_99_baisse)

    # Retour de la VaR et du besoin en fonds propres
    return var_99_hausse, var_99_baisse, besoin_en_fonds_propres