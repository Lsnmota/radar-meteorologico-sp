import datetime as dt
import os
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd

from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill
from openpyxl.utils import get_column_letter
import requests

# ==============================================================================
# CONFIGURAÇÕES E PARÂMETROS
# ==============================================================================
hoje = dt.date.today()
DATA_INICIO = (hoje - dt.timedelta(days=20)).strftime("%Y-%m-%d")  # D-20
DATA_FIM = (hoje + dt.timedelta(days=7)).strftime("%Y-%m-%d")  # D+7

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
  url = "https://geocoding-api.open-meteo.com/v1/search"
  params = {
      "name": nome,
      "count": 5,
      "language": "pt",
      "format": "json",
      "countryCode": "BR",
  }
  r = requests.get(url, params=params, timeout=30)
  results = r.json().get("results", [])
  for res in results:
    if "são paulo" in (res.get("admin1") or "").lower():
      return (res["latitude"], res["longitude"])
  if results:
    return (results[0]["latitude"], results[0]["longitude"])
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
      "Mon": "Seg",
      "Tue": "Ter",
      "Wed": "Qua",
      "Thu": "Qui",
      "Fri": "Sex",
      "Sat": "Sáb",
      "Sun": "Dom",
  }
  meses_pt = {
      1: "Jan",
      2: "Fev",
      3: "Mar",
      4: "Abr",
      5: "Mai",
      6: "Jun",
      7: "Jul",
      8: "Ago",
      9: "Set",
      10: "Out",
      11: "Nov",
      12: "Dez",
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
        start_type="num",
        start_value=0.0,
        start_color="63BE7B",
        mid_type="num",
        mid_value=0.5,
        mid_color="FFEB84",
        end_type="num",
        end_value=1.0,
        end_color="F8696B",
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

  print(
      f"Report enviado com sucesso para {len(destinatarios)} destinatário(s)!"
  )


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
