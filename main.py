import datetime as dt
import os
import smtplib
import time
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
import requests
from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# CONFIGURAÇÕES E PARÂMETROS
# ==============================================================================
TIMEZONE = "America/Sao_Paulo"
hoje = dt.datetime.now(ZoneInfo(TIMEZONE)).date()
DATA_INICIO = (hoje - dt.timedelta(days=20)).strftime("%Y-%m-%d")  # D-20
DATA_FIM = (hoje + dt.timedelta(days=14)).strftime("%Y-%m-%d")      # D+14

ARQUIVO_TEMPLATE_QUADS = "template_radar_sp.xlsx"  # Base v2.2 com colunas A:H nas abas Quads
ARQUIVO_SAIDA = f"radar_chuva_sp_{DATA_INICIO}_a_{DATA_FIM}.xlsx"

# Parâmetros de performance e resiliência da API Open-Meteo
BATCH_SIZE = 100                 # Coordenadas por requisição (evita HTTP 414)
PAUSA_ENTRE_LOTES = 10.0         # Segundos entre lotes (dilui cota de ~600 coords/min)
PRECISAO_COORDENADAS = 2         # 2 = grade ~1,1 km (~736 pts) | None = exata (~1.711 pts)

BAIRROS_SP = {
    "Sé (Centro)": (-23.5505, -46.6333),
    "Bela Vista": (-23.5590, -46.6534),
    "Liberdade": (-23.5583, -46.6350),
    "Pinheiros": (-23.5629, -46.6821),
    "Itaim Bibi": (-23.5852, -46.6819),
    "Moema": (-23.6003, -46.6664),
    "Vila Mariana": (-23.5875, -46.6344),
    "Campo Belo": (-23.6193, -46.6694),
    "Santo Amaro": (-23.6544, -46.7106),
    "Butantã": (-23.5709, -46.7106),
    "Lapa": (-23.5275, -46.7028),
    "Vila Leopoldina": (-23.5286, -46.7291),
    "Santana": (-23.5064, -46.6236),
    "Tucuruvi": (-23.4805, -46.6031),
    "Freguesia do Ó": (-23.4917, -46.6822),
    "Pirituba": (-23.4772, -46.7183),
    "Perus": (-23.4014, -46.7472),
    "Brasilândia": (-23.4547, -46.6825),
    "Tatuapé": (-23.5405, -46.5765),
    "Mooca": (-23.5605, -46.5975),
    "Penha": (-23.5303, -46.5406),
    "Itaquera": (-23.5361, -46.4478),
    "São Miguel Paulista": (-23.4989, -46.4444),
    "Guaianases": (-23.5461, -46.4106),
    "Cidade Tiradentes": (-23.5875, -46.4028),
    "Vila Prudente": (-23.5814, -46.5789),
    "Ipiranga": (-23.5911, -46.6079),
    "Cidade Ademar": (-23.6764, -46.6667),
    "Capão Redondo": (-23.6683, -46.7742),
    "Campo Limpo": (-23.6473, -46.7581),
    "Grajaú": (-23.7628, -46.6975),
    "Parelheiros": (-23.8153, -46.7517),
}

COORDENADAS_ESTATICAS_SP = {
    "São Paulo": (-23.5505, -46.6333),
    "Guarulhos": (-23.4628, -46.5333),
    "Campinas": (-22.9056, -47.0608),
    "São Bernardo do Campo": (-23.6939, -46.5650),
    "Santo André": (-23.6639, -46.5333),
    "Sorocaba": (-23.5017, -47.4581),
    "Osasco": (-23.5325, -46.7917),
    "Ribeirão Preto": (-21.1775, -47.8103),
    "São José dos Campos": (-23.1794, -45.8869),
    "São José do Rio Preto": (-20.8203, -49.3797),
    "Mogi das Cruzes": (-23.5206, -46.1854),
    "Jundiaí": (-23.1864, -46.8842),
    "Piracicaba": (-22.7253, -47.6492),
    "Santos": (-23.9608, -46.3339),
    "Mauá": (-23.6678, -46.4614),
    "Diadema": (-23.6861, -46.6228),
    "Carapicuíba": (-23.5225, -46.8356),
    "Bauru": (-22.3147, -49.0606),
    "Itaquaquecetuba": (-23.4861, -46.3483),
    "Franca": (-20.5386, -47.4008),
    "Praia Grande": (-24.0058, -46.4028),
    "São Vicente": (-23.9631, -46.3919),
    "Barueri": (-23.5106, -46.8761),
    "Taubaté": (-23.0264, -45.5553),
    "Suzano": (-23.5414, -46.3103),
    "Limeira": (-22.5647, -47.4017),
    "Guarujá": (-23.9931, -46.2564),
    "Sumaré": (-22.8219, -47.2669),
    "Cotia": (-23.6039, -46.9189),
    "Taboão da Serra": (-23.6261, -46.7800),
    "Indaiatuba": (-23.0903, -47.2181),
    "São Carlos": (-22.0175, -47.8908),
    "Embu das Artes": (-23.6489, -46.8522),
    "Araraquara": (-21.7944, -48.1756),
    "Jacareí": (-23.3053, -45.9658),
    "Americana": (-22.7392, -47.3314),
    "Marília": (-22.2139, -49.9458),
    "Itapevi": (-23.5489, -46.9342),
    "Hortolândia": (-22.8583, -47.2200),
    "Presidente Prudente": (-22.1256, -51.3889),
    "Rio Claro": (-22.4114, -47.5614),
    "Araçatuba": (-21.2089, -50.4328),
    "Ferraz de Vasconcelos": (-23.5411, -46.3686),
    "Santa Bárbara d'Oeste": (-22.7553, -47.4142),
    "Francisco Morato": (-23.2817, -46.7456),
    "Itapecerica da Serra": (-23.7172, -46.8489),
    "Itu": (-23.2642, -47.2992),
    "Bragança Paulista": (-22.9528, -46.5419),
    "Pindamonhangaba": (-22.9244, -45.4617),
    "Itapetininga": (-23.5917, -48.0531),
    "São Caetano do Sul": (-23.6228, -46.5542),
    "Franco da Rocha": (-23.3289, -46.7289),
    "Mogi Guaçu": (-22.3719, -46.9428),
    "Jaú": (-22.2961, -48.5586),
    "Botucatu": (-22.8858, -48.4450),
    "Atibaia": (-23.1169, -46.5564),
    "Santana de Parnaíba": (-23.4439, -46.9178),
    "Araras": (-22.3572, -47.3842),
    "Cubatão": (-23.8950, -46.4253),
    "Valinhos": (-22.9708, -46.9961),
    "Sertãozinho": (-21.1356, -47.9903),
    "Jandira": (-23.5278, -46.9028),
    "Ribeirão Pires": (-23.7139, -46.4136),
    "Birigui": (-21.2886, -50.3400),
    "Votorantim": (-23.5461, -47.4378),
    "Caraguatatuba": (-23.6228, -45.4128),
    "Várzea Paulista": (-23.2128, -46.8272),
    "Tatuí": (-23.3556, -47.8569),
    "Barretos": (-20.5572, -48.5678),
    "Itatiba": (-23.0058, -46.8389),
    "Guaratinguetá": (-22.8164, -45.1925),
    "Catanduva": (-21.1378, -48.9728),
    "Salto": (-23.2008, -47.2869),
    "Poá": (-23.5289, -46.3444),
    "Ourinhos": (-22.9789, -49.8706),
    "Paulínia": (-22.7611, -47.1539),
    "Resende": (-22.4689, -44.4497),
}


# ==============================================================================
# LEITURA DE LOCALIDADES E BASE QUADS
# ==============================================================================
def normalizar_coord(lat, lon):
    lat_f, lon_f = float(lat), float(lon)
    if PRECISAO_COORDENADAS is not None:
        return (round(lat_f, PRECISAO_COORDENADAS), round(lon_f, PRECISAO_COORDENADAS))
    return (lat_f, lon_f)


def obter_cidades_monitoradas():
    cidades = list(COORDENADAS_ESTATICAS_SP.keys())
    return pd.DataFrame([{"cidade": nome, "populacao": 0} for nome in cidades])


def carregar_base_quads(caminho_template=ARQUIVO_TEMPLATE_QUADS):
    """Lê os metadados (colunas A:H) da aba 'Quads (Prob%)' do template v2.2."""
    if not os.path.exists(caminho_template):
        raise FileNotFoundError(f"Template de Quads não encontrado: {caminho_template}")

    df_quads = pd.read_excel(
        caminho_template,
        sheet_name="Quads (Prob%)",
        header=1,
        usecols="A:H",
    ).dropna(subset=["LATITUDE", "LONGITUDE"])

    df_quads["LATITUDE"] = pd.to_numeric(df_quads["LATITUDE"], errors="raise")
    df_quads["LONGITUDE"] = pd.to_numeric(df_quads["LONGITUDE"], errors="raise")
    return df_quads


def geocodificar(nome):
    if nome in COORDENADAS_ESTATICAS_SP:
        return COORDENADAS_ESTATICAS_SP[nome]
    url = "https://geocoding-api.open-meteo.com/v1/search"
    params = {"name": nome, "count": 5, "language": "pt", "format": "json", "countryCode": "BR"}
    for _ in range(3):
        try:
            r = requests.get(url, params=params, timeout=15)
            r.raise_for_status()
            results = r.json().get("results", [])
            time.sleep(0.3)
            for res in results:
                if "são paulo" in (res.get("admin1") or "").lower():
                    return (res["latitude"], res["longitude"])
            if results:
                return (results[0]["latitude"], results[0]["longitude"])
            return None
        except Exception:
            time.sleep(1.5)
    return None


# ==============================================================================
# CONSULTA UNIFICADA EM LOTE NA OPEN-METEO
# ==============================================================================
def consultar_open_meteo_em_lotes(lista_coords_unicas, data_inicio, data_fim, timezone=TIMEZONE):
    """
    Consulta uma lista deduplicada de tuplas (lat, lon) em lotes de BATCH_SIZE,
    retornando dois dicionários indexados por (lat, lon):
      - cache_prob[(lat, lon)] -> pd.Series (datas -> prob 0.0..1.0)
      - cache_vol[(lat, lon)]  -> pd.Series (datas -> vol mm)
    """
    todas_datas = list(pd.date_range(data_inicio, data_fim, freq="D").date)
    cache_prob = {}
    cache_vol = {}

    total_lotes = (len(lista_coords_unicas) + BATCH_SIZE - 1) // BATCH_SIZE
    session = requests.Session()
    session.headers.update({"User-Agent": "radar-meteorologico-sp/2.2"})

    for idx_lote in range(total_lotes):
        lote = lista_coords_unicas[idx_lote * BATCH_SIZE : (idx_lote + 1) * BATCH_SIZE]
        params = {
            "latitude": ",".join(str(c[0]) for c in lote),
            "longitude": ",".join(str(c[1]) for c in lote),
            "hourly": "precipitation_probability",
            "daily": "precipitation_sum",
            "timezone": timezone,
            "start_date": str(data_inicio),
            "end_date": str(data_fim),
        }

        dados = None
        for tentativa in range(1, 6):
            try:
                r = session.get("https://api.open-meteo.com/v1/forecast", params=params, timeout=(15, 60))
                if r.status_code == 429:
                    retry_after = r.headers.get("Retry-After")
                    espera = max(65, int(retry_after)) if retry_after and retry_after.isdigit() else 65
                    print(f"   [Rate Limit 429] Lote {idx_lote + 1}/{total_lotes} (tentativa {tentativa}/5), aguardando {espera}s...")
                    time.sleep(espera)
                    continue
                r.raise_for_status()
                dados = r.json()
                break
            except Exception as e:
                espera = 5.0 * tentativa
                print(f"   [Tentativa {tentativa}/5] Erro no lote {idx_lote + 1}/{total_lotes}: {e}. Aguardando {espera}s...")
                time.sleep(espera)

        if dados is None:
            raise RuntimeError(f"Falha definitiva ao consultar o lote {idx_lote + 1}/{total_lotes} na Open-Meteo.")

        if isinstance(dados, dict):
            dados = [dados]

        for coord, res in zip(lote, dados):
            serie_prob = pd.Series(index=todas_datas, dtype=float)
            serie_vol = pd.Series(index=todas_datas, dtype=float)

            if "hourly" in res and "precipitation_probability" in res["hourly"]:
                probs_limpas = [
                    (p / 100.0 if p is not None else np.nan)
                    for p in res["hourly"]["precipitation_probability"]
                ]
                df_h = pd.DataFrame({
                    "time": pd.to_datetime(res["hourly"]["time"]),
                    "prob": probs_limpas,
                })
                df_h["data"] = df_h["time"].dt.date
                serie_prob.update(df_h.groupby("data")["prob"].mean())

            if "daily" in res and "precipitation_sum" in res["daily"]:
                df_d = pd.DataFrame({
                    "time": pd.to_datetime(res["daily"]["time"]),
                    "vol": res["daily"]["precipitation_sum"],
                })
                df_d["data"] = df_d["time"].dt.date
                serie_vol.update(df_d.set_index("data")["vol"])

            cache_prob[coord] = serie_prob
            cache_vol[coord] = serie_vol

        print(f"   Lote {idx_lote + 1}/{total_lotes} concluído ({len(lote)} coordenadas).")
        if idx_lote + 1 < total_lotes:
            time.sleep(PAUSA_ENTRE_LOTES)

    return cache_prob, cache_vol


def montar_dataframes_entidades(dict_entidades, cache_prob, cache_vol, todas_datas):
    nomes = list(dict_entidades.keys())
    df_prob = pd.DataFrame(index=nomes, columns=todas_datas, dtype=float)
    df_vol = pd.DataFrame(index=nomes, columns=todas_datas, dtype=float)
    for nome, coord in dict_entidades.items():
        chave = normalizar_coord(*coord)
        df_prob.loc[nome] = cache_prob[chave]
        df_vol.loc[nome] = cache_vol[chave]
    return df_prob, df_vol


def montar_matrizes_quads(df_quads_meta, cache_prob, cache_vol, todas_datas):
    chaves = [normalizar_coord(lat, lon) for lat, lon in zip(df_quads_meta["LATITUDE"], df_quads_meta["LONGITUDE"])]
    matriz_prob = np.vstack([cache_prob[c].reindex(todas_datas).to_numpy(dtype=float) for c in chaves])
    matriz_vol = np.vstack([cache_vol[c].reindex(todas_datas).to_numpy(dtype=float) for c in chaves])
    return matriz_prob, matriz_vol


# ==============================================================================
# GERAÇÃO DO EXCEL (6 ABAS - PADRÃO v2.2)
# ==============================================================================
def gerar_excel(
    df_cidades_prob,
    df_cidades_vol,
    df_bairros_prob,
    df_bairros_vol,
    df_quads_meta,
    matriz_quads_prob,
    matriz_quads_vol,
    caminho_saida,
):
    wb = Workbook()

    font_header = Font(name="Calibri", size=11, bold=True, color="000000")
    font_normal = Font(name="Calibri", size=11, bold=False, color="000000")
    font_quad_meta = Font(name="Calibri", size=10, bold=False, color="000000")
    font_quad_hdr_azul = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_quad_hdr_amar = Font(name="Calibri", size=10, bold=True, color="000000")

    fill_header = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    fill_estado = PatternFill(start_color="E6EDF5", end_color="E6EDF5", fill_type="solid")
    fill_azul = PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")
    fill_amarelo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    dias_semana_pt = {"Mon": "Seg", "Tue": "Ter", "Wed": "Qua", "Thu": "Qui", "Fri": "Sex", "Sat": "Sáb", "Sun": "Dom"}
    meses_pt = {
        1: "Jan", 2: "Fev", 3: "Mar", 4: "Abr", 5: "Mai", 6: "Jun",
        7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez",
    }

    def aplicar_color_scale(ws, intervalo, tipo_dado):
        if tipo_dado == "prob":
            rule = ColorScaleRule(
                start_type="num", start_value=0.0, start_color="63BE7B",
                mid_type="num", mid_value=0.5, mid_color="FFEB84",
                end_type="num", end_value=1.0, end_color="F8696B",
            )
        else:
            rule = ColorScaleRule(
                start_type="num", start_value=0.0, start_color="63BE7B",
                mid_type="num", mid_value=10.0, mid_color="FFEB84",
                end_type="num", end_value=25.0, end_color="F8696B",
            )
        ws.conditional_formatting.add(intervalo, rule)

    def criar_aba_padrao(ws, df_dados, eh_cidades, tipo_dado):
        ws.views.sheetView[0].showGridLines = True
        ws.append(["Dia"] + [dias_semana_pt.get(d.strftime("%a"), d.strftime("%a")) for d in df_dados.columns])
        ws.append(["Cidade" if eh_cidades else "Bairro"] + [f"{meses_pt[d.month]}-{d.day:02d}" for d in df_dados.columns])

        for loc, row_vals in df_dados.iterrows():
            if eh_cidades and loc == df_dados.index[1]:
                ws.append(["Cidades"] + [f"{meses_pt[d.month]}-{d.day:02d}" for d in df_dados.columns])
            ws.append([loc] + [None if pd.isna(v) else float(v) for v in row_vals.values])

        max_row = ws.max_row
        max_col = ws.max_column

        for r in range(1, max_row + 1):
            p_cel = ws.cell(r, 1).value
            for c in range(1, max_col + 1):
                cell = ws.cell(r, c)
                cell.border = thin_border
                if r in (1, 2) or p_cel == "Cidades":
                    cell.font = font_header
                    cell.fill = fill_header
                    cell.alignment = align_left if c == 1 else align_center
                else:
                    cell.font = font_header if p_cel == "Estado de São Paulo" else font_normal
                    cell.fill = fill_estado if p_cel == "Estado de São Paulo" else PatternFill()
                    cell.alignment = align_left if c == 1 else align_center
                if c > 1 and isinstance(cell.value, (int, float)):
                    cell.number_format = "0%" if tipo_dado == "prob" else "0.0"

        aplicar_color_scale(ws, f"B3:{get_column_letter(max_col)}{max_row}", tipo_dado)
        ws.column_dimensions["A"].width = 28
        for c in range(2, max_col + 1):
            ws.column_dimensions[get_column_letter(c)].width = 10

    def criar_aba_quads(ws, df_meta, matriz_valores, datas, tipo_dado):
        ws.views.sheetView[0].showGridLines = True

        # Linha 1: A1:H1 vazias; I1:AQ1 com dias da semana
        ws.append([None] * 8 + [dias_semana_pt.get(d.strftime("%a"), d.strftime("%a")) for d in datas])

        # Linha 2: Metadados A2:H2 + datas I2:AQ2
        cabecalho_meta = list(df_meta.columns[:8])
        ws.append(cabecalho_meta + [f"{meses_pt[d.month]}-{d.day:02d}" for d in datas])

        # Linhas 3..N: Metadados + série meteorológica
        meta_rows = df_meta.iloc[:, :8].where(pd.notna(df_meta.iloc[:, :8]), None).values.tolist()
        for meta_row, vals_row in zip(meta_rows, matriz_valores):
            ws.append(meta_row + [None if np.isnan(v) else float(v) for v in vals_row])

        max_row = ws.max_row
        max_col = ws.max_column

        # Estilização Linha 1 (apenas colunas >= 9)
        for c in range(9, max_col + 1):
            cell = ws.cell(1, c)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = align_center
            cell.border = thin_border

        # Estilização Linha 2
        for c in range(1, max_col + 1):
            cell = ws.cell(2, c)
            cell.border = thin_border
            cell.alignment = align_center
            if c <= 3:
                cell.font = font_quad_hdr_azul
                cell.fill = fill_azul
            elif c <= 8:
                cell.font = font_quad_hdr_amar
                cell.fill = fill_amarelo
            else:
                cell.font = font_header
                cell.fill = fill_header

        # Estilização Linhas de Dados (3..max_row)
        num_fmt = "0%" if tipo_dado == "prob" else "0.0"
        for r in range(3, max_row + 1):
            for c in range(1, max_col + 1):
                cell = ws.cell(r, c)
                cell.border = thin_border
                cell.alignment = align_center
                if c <= 8:
                    cell.font = font_quad_meta
                    if c == 2:
                        cell.number_format = "0"
                else:
                    cell.font = font_normal
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = num_fmt

        aplicar_color_scale(ws, f"I3:{get_column_letter(max_col)}{max_row}", tipo_dado)

        larguras_meta = {"A": 12, "B": 5, "C": 12, "D": 48, "E": 38, "F": 13, "G": 13, "H": 31}
        for col_letra, largura in larguras_meta.items():
            ws.column_dimensions[col_letra].width = largura
        for c in range(9, max_col + 1):
            ws.column_dimensions[get_column_letter(c)].width = 11

    # Ordem exata das abas conforme template v2.2
    datas_regua = list(df_cidades_prob.columns)

    ws1 = wb.active
    ws1.title = "Cidades (Prob%)"
    criar_aba_padrao(ws1, df_cidades_prob, True, "prob")

    ws2 = wb.create_sheet(title="Quads (Prob%)")
    criar_aba_quads(ws2, df_quads_meta, matriz_quads_prob, datas_regua, "prob")

    ws3 = wb.create_sheet(title="Quads (Vol mm)")
    criar_aba_quads(ws3, df_quads_meta, matriz_quads_vol, datas_regua, "vol")

    ws4 = wb.create_sheet(title="Cidades (Vol mm)")
    criar_aba_padrao(ws4, df_cidades_vol, True, "vol")

    ws5 = wb.create_sheet(title="Bairros (Prob%)")
    criar_aba_padrao(ws5, df_bairros_prob, False, "prob")

    ws6 = wb.create_sheet(title="Bairros (Vol mm)")
    criar_aba_padrao(ws6, df_bairros_vol, False, "vol")

    wb.save(caminho_saida)


# ==============================================================================
# ENVIO DE E-MAIL E TELEGRAM
# ==============================================================================
def enviar_email(caminho_arquivo):
    remetente = os.environ.get("EMAIL_REMETENTE") or os.environ.get("EMAIL_USER")
    senha = os.environ.get("EMAIL_SENHA") or
