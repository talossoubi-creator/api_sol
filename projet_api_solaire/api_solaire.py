# Rapport solaire automatique via API Open-Meteo
# Recupere les donnees d'irradiation et de temperature pour une localisation,
# analyse les performances journalieres et genere un rapport Excel avec graphiques.

import requests
import pandas as pd
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.drawing.image import Image


def recuperer_donnees(latitude, longitude, date_debut, date_fin):
    """Appelle l'API Open-Meteo et retourne les donnees JSON."""
    url = "https://archive-api.open-meteo.com/v1/archive"
    params = {
        "latitude": latitude,
        "longitude": longitude,
        "start_date": date_debut,
        "end_date": date_fin,
        "hourly": "temperature_2m,shortwave_radiation",
        "timezone": "Europe/Paris"
    }
    response = requests.get(url, params=params)
    return response.json()


def json_vers_dataframe(data):
    """Transforme le JSON de l'API en DataFrame pandas propre."""
    hourly = data["hourly"]
    df = pd.DataFrame(hourly)
    df["time"] = pd.to_datetime(df["time"])
    return df


def analyser(df):
    """Agrege les donnees par jour et calcule les indicateurs."""
    df["date"] = df["time"].dt.date
    daily = df.groupby("date").agg(
        irr_moy=("shortwave_radiation", "mean"),
        irr_sum=("shortwave_radiation", "sum"),
        temp_moy=("temperature_2m", "mean")
    )
    # Conversion Wh/m2 en kWh/m2
    daily["irr_kwh"] = daily["irr_sum"] / 1000
    return daily


def generer_rapport(df, nom_fichier, lieu):
    """Genere un rapport Excel avec graphiques d'irradiation et temperature."""
    # Graphique irradiation
    df["irr_kwh"].plot(kind="line")
    plt.title("Irradiation journaliere - " + lieu)
    plt.ylabel("kWh/m2/jour")
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig("graph_irradiation.png")
    plt.close()

    # Graphique temperature
    df["temp_moy"].plot(kind="line")
    plt.title("Temperature moyenne - " + lieu)
    plt.ylabel("°C")
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig("graph_temperature.png")
    plt.close()

    # Export Excel avec insertion des graphiques
    df.to_excel(nom_fichier, sheet_name="tableau du rapport")
    wb = openpyxl.load_workbook(nom_fichier)
    ws = wb.active
    img1 = Image("graph_irradiation.png")
    ws.add_image(img1, "G2")
    img2 = Image("graph_temperature.png")
    ws.add_image(img2, "G30")
    wb.save(nom_fichier)


# --- Execution ---
# Changer les coordonnees et les dates pour analyser un autre site
data = recuperer_donnees(43.69, 5.76, "2026-05-01", "2026-05-31")
df = json_vers_dataframe(data)
daily = analyser(df)
generer_rapport(daily, "rapport_solaire.xlsx", "Cadarache")

