import datetime as dt
import os
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import requests
import time

# ==============================================================================
# CONFIGURAÇÕES E PARÂMETROS
# ==============================================================================
hoje = dt.date.today()
DATA_INICIO = (hoje - dt.timedelta(days=20)).strftime("%Y-%m-%d")  # D-20
DATA_FIM = (hoje + dt.timedelta(days=14)).strftime("%Y-%m-%d")  # D+14

POPULACAO_MINIMA = 100_000
TIMEZONE = "America/Sao_Paulo"
ARQUIVO_SAIDA = f"radar_chuva_sp_{DATA_INICIO}_a_{DATA_FIM}.xlsx"

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

# ==============================================================================
# DICIONÁRIO DE COORDENADAS FIXAS (EVITA DEPENDÊNCIA DA API DE GEOCODIFICAÇÃO)
# ==============================================================================
COORDENADAS_ESTATICAS_SP = {
    "São Paulo": (-23.5505, -46.6333), "Guarulhos": (-23.4628, -46.5333),
    "Campinas": (-22.9056, -47.0608), "São Bernardo do Campo": (-23.6939, -46.5650),
    "Santo André": (-23.6639, -46.5333), "Sorocaba": (-23.5017, -47.4581),
    "Osasco": (-23.5325, -46.7917), "Ribeirão Preto": (-21.1775, -47.8103),
    "São José dos Campos": (-23.1794, -45.8869), "São José do Rio Preto": (-20.8203, -49.3797),
    "Mogi das Cruzes": (-23.5206, -46.1854), "Jundiaí": (-23.1864, -46.8842),
    "Piracicaba": (-22.7253, -47.6492), "Santos": (-23.9608, -46.3339),
    "Mauá": (-23.6678, -46.4614), "Diadema": (-23.6861, -46.6228),
    "Carapicuíba": (-23.5225, -46.8356), "Bauru": (-22.3147, -49.0606),
    "Itaquaquecetuba": (-23.4861, -46.3483), "Franca": (-20.5386, -47.4008),
    "Praia Grande": (-24.0058, -46.4028), "São Vicente": (-23.9631, -46.3919),
    "Barueri": (-23.5106, -46.8761), "Taubaté": (-23.0264, -45.5553),
    "Suzano": (-23.5414, -46.3103), "Limeira": (-22.5647, -47.4017),
    "Guarujá": (-23.9931, -46.2564), "Sumaré": (-22.8219, -47.2669),
    "Cotia": (-23.6039, -46.9189), "Taboão da Serra": (-23.6261, -46.7800),
    "Indaiatuba": (-23.0903, -47.2181), "São Carlos": (-22.0175, -47.8908),
    "Embu das Artes": (-23.6489, -46.8522), "Araraquara": (-21.7944, -48.1756),
    "Jacareí": (-23.3053, -45.9658), "Americana": (-22.7392, -47.3314),
    "Marília": (-22.2139, -49.9458), "Itapevi": (-23.5489, -46.9342),
    "Hortolândia": (-22.8583, -47.2200), "Presidente Prudente": (-22.1256, -51.3889),
    "Rio Claro": (-22.4114, -47.5614), "Araçatuba": (-21.2089, -50.4328),
    "Ferraz de Vasconcelos": (-23.5411, -46.3686), "Santa Bárbara d'Oeste": (-22.7553, -47.4142),
    "Francisco Morato": (-23.2817, -46.7456), "Itapecerica da Serra": (-23.7172, -46.8489),
    "Itu": (-23.2642, -47.2992), "Bragança Paulista": (-22.9528, -46.5419),
    "Pindamonhangaba": (-22.9244, -45.4617), "Itapetininga": (-23.5917, -48.0531),
    "São Caetano do Sul": (-23.6228, -46.5542), "Franco da Rocha": (-23.3289, -46.7289),
    "Mogi Guaçu": (-22.3719, -46.9428), "Jaú": (-22.2961, -48.5586),
    "Botucatu": (-22.8858, -48.4450), "Atibaia": (-23.1169, -46.5564),
    "Santana de Parnaíba": (-23.4439, -46.9178), "Araras": (-22.3572, -47.3842),
    "Cubatao": (-23.8950, -46.4253), "Valinhos": (-22.9708, -46.9961),
    "Sertãozinho": (-21.1356, -47.9903), "Jandira": (-23.5278, -46.9028),
    "Ribeirão Pires": (-23.7139, -46.4136), "Birigui": (-21.2886, -50.3400),
    "Votorantim": (-23.5461, -47.4378), "Caraguatatuba": (-23.6228, -45.4128),
    "Várzea Paulista": (-23.2128, -46.8272), "Tatuí": (-23.3556, -47.8569),
    "Barretos": (-20.5572, -48.5678), "Itatiba": (-23.0058, -46.8389),
    "Guaratinguetá": (-22.8164, -45.1925), "Catanduva": (-21.1378, -48.9728),
    "Salto": (-23.2008, -47.2869), "Poá": (-23.5289, -46.3444),
    "Ourinhos": (-22.9789, -49.8706), "Paulínia": (-22.7611, -47.1539),
    "Resende": (-22.4689, -44.4497),
}

# ==============================================================================
# FUNÇÕES DE DADOS (IBGE + GEO + OPEN-METEO)
# ==============================================================================
def buscar_cidades_sp_acima_de(populacao_minima):
  r_mun = requests.get(
      "https://servicodados.ibge.gov.br/api/v1/localidades/estados/SP/municipios",
      timeout=30,
  )
  municipios = {str(m["id"]): m["nome"] for m in r_mun.json()}

  r_pop = requests.get(
      "https://servicodados.ibge.gov.br/api/v3/agregados/6579/periodos/-1/variaveis/9324?localidades=N6[all]",
      timeout=60,
  )
  pop_data = r_pop.json()
  populacao = {}
  for var in pop_data:
    for res in var.get("resultados", []):
      for s in res.get("series", []):
        cod = str(s["localidade"]["id"])
        vals = s.get("serie", {})
        if vals:
          last_k = sorted(vals.keys())[-1]
          try:
            populacao[cod] = int(vals[last_k])
          except (TypeError, ValueError):
            pass

  linhas = []
  for cod, nome in municipios.items():
    p = populacao.get(cod)
    if p is not None and p >= populacao_minima:
      linhas.append({"cidade": nome, "populacao": p})

  return pd.DataFrame(linhas).sort_values("populacao", ascending=False)


def geocodificar(nome):
  # 1. Busca primeiro no dicionário fixo de cidades
  if nome in COORDENADAS_ESTATICAS_SP:
    return COORDENADAS_ESTATICAS_SP[nome]

  # 2. Se for uma cidade nova, faz consulta com timeout tratado e pausa
  url = "https://geocoding-api.open-meteo.com/v1/search"
  params = {
      "name": nome,
      "count": 5,
      "language": "pt",
      "format": "json",
      "countryCode": "BR",
  }

  for tentativa in range(1, 4):
    try:
      r = requests.get(url, params=params, timeout=15)
      r.raise_for_status()
      results = r.json().get("results", [])
      time.sleep(0.3)  # Pausa entre chamadas

      for res in results:
        if "são paulo" in (res.get("admin1") or "").lower():
          return (res["latitude"], res["longitude"])
      if results:
        return (results[0]["latitude"], results[0]["longitude"])
      return None
    except Exception as e:
      print(f"Tentativa {tentativa} falhou para '{nome}': {e}")
      time.sleep(1.5)

  return None


def buscar_probabilidade_chuva_media(
    dict_coordenadas, data_inicio, data_fim, timezone=TIMEZONE
):
  nomes = list(dict_coordenadas.keys())
  coords = list(dict_coordenadas.values())
  todas_datas = list(pd.date_range(data_inicio, data_fim, freq="D").date)
  resultado = pd.DataFrame(index=nomes, columns=todas_datas, dtype=float)

  params = {
      "latitude": ",".join(str(c[0]) for c in coords),
      "longitude": ",".join(str(c[1]) for c in coords),
      "hourly": "precipitation_probability",
      "timezone": timezone,
      "start_date": str(data_inicio),
      "end_date": str(data_fim),
  }

  r = requests.get(
      "https://api.open-meteo.com/v1/forecast", params=params, timeout=60
  )
  dados = r.json()
  if isinstance(dados, dict):
    dados = [dados]

  for nome, res in zip(nomes, dados):
    if "hourly" in res and "precipitation_probability" in res["hourly"]:
      df_h = pd.DataFrame({
          "time": pd.to_datetime(res["hourly"]["time"]),
          "prob": [
              p / 100.0 for p in res["hourly"]["precipitation_probability"]
          ],
      })
      df_h["data"] = df_h["time"].dt.date
      resultado.loc[nome] = df_h.groupby("data")["prob"].mean()

  return resultado


# ==============================================================================
# GERAÇÃO DO EXCEL
# ==============================================================================
def gerar_excel(df_cidades_completo, df_prob_bairros, caminho_saida):
  wb = Workbook()

  font_header = Font(name="Calibri", size=11, bold=True, color="000000")
  font_estado = Font(name="Calibri", size=11, bold=True, color="000000")
  font_normal = Font(name="Calibri", size=11, bold=False, color="000000")

  fill_header = PatternFill(
      start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"
  )
  fill_estado = PatternFill(
      start_color="E6EDF5", end_color="E6EDF5", fill_type="solid"
  )

  align_center = Alignment(horizontal="center", vertical="center")
  align_left = Alignment(horizontal="left", vertical="center")
  thin_border = Border(
      left=Side(style="thin", color="D9D9D9"),
      right=Side(style="thin", color="D9D9D9"),
      top=Side(style="thin", color="D9D9D9"),
      bottom=Side(style="thin", color="D9D9D9"),
  )

  dias_semana_pt = {
      "Mon": "Seg", "Tue": "Ter", "Wed": "Qua", "Thu": "Qui",
      "Fri": "Sex", "Sat": "Sáb", "Sun": "Dom",
  }
  meses_pt = {
      1: "Jan", 2: "Fev", 3: "Mar", 4: "Abr", 5: "Mai", 6: "Jun",
      7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez",
  }

  def criar_aba(ws, df_dados, eh_cidades):
    ws.views.sheetView[0].showGridLines = True

    # Linha 1: Dias
    row1 = ["Dia"] + [
        dias_semana_pt.get(d.strftime("%a"), d.strftime("%a"))
        for d in df_dados.columns
    ]
    ws.append(row1)

    # Linha 2: Datas
    row2 = ["Cidade" if eh_cidades else "Bairro"] + [
        f"{meses_pt[d.month]}-{d.day:02d}" for d in df_dados.columns
    ]
    ws.append(row2)

    # Dados
    for loc, row_vals in df_dados.iterrows():
      if eh_cidades and loc == df_cidades_completo.index[1]:
        ws.append(["Cidades"] + [
            f"{meses_pt[d.month]}-{d.day:02d}" for d in df_dados.columns
        ])
      ws.append([loc] + list(row_vals.values))

    # Formatação
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
        elif p_cel == "Estado de São Paulo":
          cell.font = font_estado
          cell.fill = fill_estado
          cell.alignment = align_left if c == 1 else align_center
          if c > 1:
            cell.number_format = "0%"
        else:
          cell.font = font_normal
          cell.alignment = align_left if c == 1 else align_center
          if c > 1:
            cell.number_format = "0%"

    color_scale = ColorScaleRule(
        start_type="num", start_value=0.0, start_color="63BE7B",
        mid_type="num", mid_value=0.5, mid_color="FFEB84",
        end_type="num", end_value=1.0, end_color="F8696B",
    )
    col_fim = get_column_letter(max_col)
    ws.conditional_formatting.add(f"B3:{col_fim}{max_row}", color_scale)

    ws.column_dimensions["A"].width = 28
    for c in range(2, max_col + 1):
      ws.column_dimensions[get_column_letter(c)].width = 10

  ws1 = wb.active
  ws1.title = "Cidades SP"
  criar_aba(ws1, df_cidades_completo, eh_cidades=True)

  ws2 = wb.create_sheet(title="SP Capital - Bairros")
  criar_aba(ws2, df_prob_bairros, eh_cidades=False)

  wb.save(caminho_saida)


# ==============================================================================
# ENVIO DE E-MAIL
# ==============================================================================
def enviar_email(caminho_arquivo):
  remetente = os.environ.get("EMAIL_REMETENTE")
  senha = os.environ.get("EMAIL_SENHA")
  destinatarios_raw = os.environ.get(
      "EMAIL_DESTINATARIOS", remetente
  )  # fallback se não definir
  destinatarios = [d.strip() for d in destinatarios_raw.split(",") if d.strip()]

  if not remetente or not senha:
    print("ERRO: Credenciais de e-mail não configuradas no ambiente.")
    return

  msg = MIMEMultipart()
  msg["From"] = remetente
  msg["To"] = ", ".join(destinatarios)
  msg["Subject"] = f"Radar Meteorológico SP - {hoje.strftime('%d/%m/%Y')}"

  corpo = """Olá,

Segue em anexo o relatório atualizado do Radar Meteorológico do Estado de São Paulo e Bairros da Capital.

Este e-mail foi gerado automaticamente.

Atenciosamente,
Equipe de Automação"""

  msg.attach(MIMEText(corpo, "plain"))

  with open(caminho_arquivo, "rb") as f:
    part = MIMEApplication(f.read(), Name=os.path.basename(caminho_arquivo))
    part[
        "Content-Disposition"
    ] = f'attachment; filename="{os.path.basename(caminho_arquivo)}"'
    msg.attach(part)

  with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
    server.login(remetente, senha)
    server.sendmail(remetente, destinatarios, msg.as_string())

  print(f"Report enviado com sucesso para {len(destinatarios)} destinatário(s)!")


# ==============================================================================
# EXECUÇÃO PRINCIPAL
# ==============================================================================
if __name__ == "__main__":
  print("1. Obtendo cidades via IBGE...")
  df_cidades = buscar_cidades_sp_acima_de(POPULACAO_MINIMA)

  print("2. Geocodificando localidades...")
  coords_cidades = {}
  for nome in df_cidades["cidade"]:
    c = geocodificar(nome)
    if c:
      coords_cidades[nome] = c

  print("3. Buscando previsão de chuva na Open-Meteo...")
  df_prob_cidades = buscar_probabilidade_chuva_media(
      coords_cidades, DATA_INICIO, DATA_FIM
  )
  df_prob_bairros = buscar_probabilidade_chuva_media(
      BAIRROS_SP, DATA_INICIO, DATA_FIM
  )

  print("4. Consolidando linha do Estado de São Paulo...")
  df_sp_media = pd.DataFrame(
      [df_prob_cidades.mean(axis=0)],
      index=["Estado de São Paulo"],
      columns=df_prob_cidades.columns,
  )
  df_cidades_completo = pd.concat([df_sp_media, df_prob_cidades])

  print("5. Gerando arquivo Excel...")
  gerar_excel(df_cidades_completo, df_prob_bairros, ARQUIVO_SAIDA)

  print("6. Enviando relatório por e-mail...")
  enviar_email(ARQUIVO_SAIDA)
  print("Processo concluído!")
