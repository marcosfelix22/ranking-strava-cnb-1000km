import requests
import pandas as pd
import os
from datetime import datetime

# --- CONFIGURAÇÃO ---
CLIENT_ID = os.environ.get('CLIENT_ID')
CLIENT_SECRET = os.environ.get('CLIENT_SECRET')
REFRESH_TOKEN = os.environ.get('REFRESH_TOKEN')
CLUB_ID = '1921916'
NOME_ARQUIVO = 'Ranking_CNB_1000km_2026.xlsx'

# --- TOKEN ---
def obter_access_token():
    payload = {
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'refresh_token': REFRESH_TOKEN,
        'grant_type': 'refresh_token'
    }
    res = requests.post("https://www.strava.com/oauth/token", data=payload).json()
    return res.get('access_token')

# --- FORMATAÇÃO ---
def formatar_km(valor):
    return f"{valor:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".") + " km"

def formatar_alt(valor):
    return f"{int(valor):,}".replace(",", ".") + " m"

# --- CARREGA DADOS EXISTENTES ---
if os.path.exists(NOME_ARQUIVO):
    try:
        df_atividades = pd.read_excel(NOME_ARQUIVO, sheet_name='Atividades')
    except:
        df_atividades = pd.DataFrame(columns=[
            'activity_id', 'atleta', 'data', 'distancia_km', 'altimetria_m'
        ])
else:
    df_atividades = pd.DataFrame(columns=[
        'activity_id', 'atleta', 'data', 'distancia_km', 'altimetria_m'
    ])

# --- TOKEN ---
access_token = obter_access_token()

if access_token:
    print("Buscando atividades do clube...")

    novas_atividades = []

    for pagina in range(1, 11):

        url = f"https://www.strava.com/api/v3/clubs/{CLUB_ID}/activities"

        resp = requests.get(
            url,
            headers={'Authorization': f'Bearer {access_token}'},
            params={'per_page': 200, 'page': pagina}
        ).json()

        if not resp or 'errors' in resp:
            break

        for act in resp:

            # ID real do Strava
            activity_id = act.get('id')

            if not activity_id:
                continue

            # Evita duplicação
            if activity_id in df_atividades['activity_id'].values:
                continue

            # Apenas corrida
            if act.get('sport_type') != 'Run':
                continue

            # Data
            try:
                data_atividade = datetime.strptime(
                    act['start_date'],
                    '%Y-%m-%dT%H:%M:%SZ'
                )
            except:
                continue

            # Apenas 2026
            if data_atividade.year != 2026:
                continue

            # Distância
            dist_km = act.get('distance', 0) / 1000
            if dist_km <= 0:
                continue

            alt = act.get('total_elevation_gain', 0)

            # Atleta
            atleta = act.get('athlete', {})
            nome = f"{athlete.get('firstname','')} {athlete.get('lastname','')}".strip().upper()

            novas_atividades.append({
                'activity_id': activity_id,
                'atleta': nome,
                'data': data_atividade,
                'distancia_km': dist_km,
                'altimetria_m': alt
            })

    # --- JUNTA HISTÓRICO ---
    if len(novas_atividades) > 0:
        df_atividades = pd.concat([
            df_atividades,
            pd.DataFrame(novas_atividades)
        ], ignore_index=True)

    # remove duplicados reais
    df_atividades = df_atividades.drop_duplicates(subset=['activity_id'])

    # --- RANKING ---
    ranking = df_atividades.groupby('atleta').agg({
        'distancia_km': 'sum',
        'altimetria_m': 'sum',
        'activity_id': 'count'
    }).reset_index()

    ranking.columns = [
        'Atleta',
        'KM Total',
        'Altimetria (m)',
        'Treinos'
    ]

    ranking = ranking.sort_values(by='KM Total', ascending=False)

    # posição
    ranking.insert(0, 'Posição', range(1, len(ranking) + 1))

    # meta
    ranking['KM Total'] = pd.to_numeric(ranking['KM Total'], errors='coerce')

ranking['Meta 1000km (%)'] = (ranking['KM Total'] / 1000 * 100).round(1)

    # formato visual
    df_visual = ranking.copy()
    df_visual['KM Total'] = df_visual['KM Total'].apply(formatar_km)
    df_visual['Altimetria (m)'] = df_visual['Altimetria (m)'].apply(formatar_alt)

    # --- SALVAR ---
    with pd.ExcelWriter(NOME_ARQUIVO) as writer:
        df_visual.to_excel(writer, sheet_name='Ranking', index=False)
        df_atividades.to_excel(writer, sheet_name='Atividades', index=False)

    print("Ranking atualizado com sucesso!")
