# Radar Meteorológico do Estado de SP

Este notebook gera uma planilha Excel com a **probabilidade de chuva** (%) para os próximos dias, para:

1. Todas as cidades do estado de São Paulo com mais de 100 mil habitantes (dados de população do **IBGE**, buscados automaticamente e sempre atualizados).
2. Os bairros da cidade de São Paulo (capital), em uma aba separada.

**Fontes de dados:**
- População dos municípios: API de localidades do IBGE + tabela SIDRA 6579 (estimativas populacionais).
- Coordenadas geográficas das cidades: API de geocodificação da [Open-Meteo](https://open-meteo.com).
- Previsão de chuva: API de previsão do tempo da [Open-Meteo](https://open-meteo.com) (gratuita, sem necessidade de chave de API).

**Limitações importantes:**
- A Open-Meteo fornece probabilidade de precipitação para um período de até ~92 dias no passado e ~16 dias no futuro, a partir da data em que o notebook é executado. Datas fora dessa janela não retornam dados.
- A resolução dos modelos meteorológicos é de alguns quilômetros. Isso significa que bairros muito próximos entre si podem apresentar valores iguais ou muito parecidos — a granularidade "por bairro" é uma aproximação, não uma medição hiperlocal.
- A lista de bairros da capital é uma lista curada (editável na célula de configuração), já que a API de geocodificação não cobre bairros de forma confiável.

Basta rodar as células em ordem. A célula 2 (Configurações) é o único lugar que você normalmente precisa editar.


## 1. Instalar/importar dependências

# As bibliotecas abaixo já vêm instaladas no Google Colab na maioria dos casos.
# O pip install garante que estejam disponíveis mesmo assim.
!pip install --quiet openpyxl requests pandas

import time
import datetime as dt

import requests
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter

print("Bibliotecas carregadas com sucesso.")


## 2. Configurações — ajuste aqui

- `DATA_INICIO` / `DATA_FIM`: período desejado, formato `AAAA-MM-DD`. Exemplo do enunciado: 1 a 20 de setembro.
  Lembre-se do limite da API: aproximadamente 92 dias no passado e 16 dias no futuro, contados a partir de **hoje** (data em que a célula é executada).
- `POPULACAO_MINIMA`: filtro de população dos municípios (padrão: 100.000).
- `BAIRROS_SP`: lista de bairros da capital que entrarão na segunda aba. Adicione, remova ou ajuste as coordenadas (latitude, longitude) livremente.


# ----------------------- PERÍODO DA PREVISÃO -----------------------
hoje = dt.date.today()

DATA_INICIO = (hoje - dt.timedelta(days=20)).strftime("%Y-%m-%d")  # D-20
DATA_FIM = (hoje + dt.timedelta(days=7)).strftime("%Y-%m-%d")  # D+7

# ----------------------- FILTRO DE POPULAÇÃO ------------------------
POPULACAO_MINIMA = 100_000

# ----------------------- FUSO HORÁRIO --------------------------------
TIMEZONE = "America/Sao_Paulo"

# ----------------------- NOME DO ARQUIVO DE SAÍDA --------------------
ARQUIVO_SAIDA = f"radar_chuva_sp_{DATA_INICIO}_a_{DATA_FIM}.xlsx"

# ----------------------- BAIRROS DA CAPITAL (aba 2) -------------------
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

print(f"Data de execução (D-0): {hoje.strftime('%Y-%m-%d')}")
print(f"Período: {DATA_INICIO} a {DATA_FIM}")
print(f"População mínima: {POPULACAO_MINIMA:,}".replace(",", "."))
print(f"Bairros configurados: {len(BAIRROS_SP)}")

## 3. Cidades de SP com população acima do limite (IBGE)

Busca em duas etapas, ligadas pelo código IBGE de cada município (mais confiável do que casar pelo nome):

1. Lista de todos os municípios do estado de SP (`/localidades/estados/SP/municipios`).
2. Estimativa populacional mais recente de cada município (tabela SIDRA 6579, variável 9324).


def buscar_municipios_sp():
    '''Retorna dict {codigo_ibge (str): nome_do_municipio}.'''
    url = "https://servicodados.ibge.gov.br/api/v1/localidades/estados/SP/municipios"
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    return {str(m["id"]): m["nome"] for m in r.json()}


def buscar_populacao_estimada():
    '''Retorna dict {codigo_ibge (str): populacao (int)} com a estimativa mais recente
    disponível na tabela SIDRA 6579 (Estimativas de População) para todos os municípios do Brasil.'''
    url = (
        "https://servicodados.ibge.gov.br/api/v3/agregados/6579/periodos/-1"
        "/variaveis/9324?localidades=N6[all]"
    )
    r = requests.get(url, timeout=60)
    r.raise_for_status()
    dados = r.json()

    populacao = {}
    for variavel in dados:
        for resultado in variavel.get("resultados", []):
            for serie in resultado.get("series", []):
                codigo = str(serie["localidade"]["id"])
                valores = serie.get("serie", {})
                if not valores:
                    continue
                ultimo_periodo = sorted(valores.keys())[-1]
                valor = valores[ultimo_periodo]
                try:
                    populacao[codigo] = int(valor)
                except (TypeError, ValueError):
                    continue  # valor não disponível ("...", "-", etc.)
    return populacao


def buscar_cidades_sp_acima_de(populacao_minima):
    municipios = buscar_municipios_sp()
    populacao = buscar_populacao_estimada()

    linhas = []
    for codigo, nome in municipios.items():
        pop = populacao.get(codigo)
        if pop is not None and pop >= populacao_minima:
            linhas.append({"codigo_ibge": codigo, "cidade": nome, "populacao": pop})

    df = pd.DataFrame(linhas).sort_values("populacao", ascending=False).reset_index(drop=True)
    return df


print("Buscando dados de população no IBGE...")
df_cidades = buscar_cidades_sp_acima_de(POPULACAO_MINIMA)
print(f"{len(df_cidades)} cidades encontradas com população >= {POPULACAO_MINIMA:,}.".replace(",", "."))
df_cidades.head(15)


## 4. Geocodificação (coordenadas de cada cidade)

Usa a API de geocodificação da Open-Meteo, com preferência por resultados cujo estado (`admin1`) seja São Paulo, para evitar confusão com cidades homônimas em outros estados.


import requests
import time

CACHE_GEOCODIFICACAO = {}

def geocodificar(nome, estado_preferido="São Paulo", pais="BR", max_tentativas=3):
    if nome in CACHE_GEOCODIFICACAO:
        return CACHE_GEOCODIFICACAO[nome]

    url = "https://geocoding-api.open-meteo.com/v1/search"
    params = {
        "name": nome,
        "count": 10,
        "language": "pt",
        "format": "json",
        "countryCode": pais
    }

    for tentativa in range(1, max_tentativas + 1):
        try:
            r = requests.get(url, params=params, timeout=60)
            r.raise_for_status()
            resultados = r.json().get("results", []) or []
            break
        except requests.RequestException as e:
            if tentativa == max_tentativas:
                print(f"AVISO: erro ao geocodificar '{nome}' após {max_tentativas} tentativas: {e}")
                CACHE_GEOCODIFICACAO[nome] = None
                return None
            time.sleep(2 * tentativa) # Pausa progressiva antes de tentar de novo

    escolhido = next(
        (res for res in resultados if estado_preferido.lower() in (res.get("admin1") or "").lower()),
        None,
    )
    if escolhido is None and resultados:
        escolhido = resultados[0]
        print(f"AVISO: '{nome}' — não encontrei resultado em {estado_preferido}, usando o melhor resultado disponível ({escolhido.get('admin1')}, {escolhido.get('country')}).")

    if escolhido is None:
        print(f"AVISO: '{nome}' — não foi possível geocodificar, cidade será omitida da planilha.")
        CACHE_GEOCODIFICACAO[nome] = None
        return None

    coord = (escolhido["latitude"], escolhido["longitude"])
    CACHE_GEOCODIFICACAO[nome] = coord
    return coord

print("Geocodificando cidades com sistema de nova tentativa...")
coordenadas_cidades = {}
for i, nome in enumerate(df_cidades["cidade"], start=1):
    coord = geocodificar(nome)
    if coord is not None:
        coordenadas_cidades[nome] = coord
    if i % 20 == 0:
        print(f"  {i}/{len(df_cidades)} processadas...")
    time.sleep(0.3)  # Aumentado levemente para respeitar o limite da API pública

print(f"Concluído: {len(coordenadas_cidades)}/{len(df_cidades)} cidades geocodificadas.")
coordenadas_bairros = dict(BAIRROS_SP)

## 5. Previsão de probabilidade de chuva (Open-Meteo)

Faz requisições em lotes (várias localidades por chamada) para reduzir o número de chamadas à API.
O resultado é uma tabela com uma linha por localidade e uma coluna por dia, com a probabilidade máxima diária de precipitação (0 a 1, onde 1 = 100%).




import pandas as pd
import requests

# ==============================================================================
# FUNÇÃO DE BUSCA NA API OPEN-METEO
# ==============================================================================


def buscar_probabilidade_chuva_media(
    dict_coordenadas,
    data_inicio,
    data_fim,
    timezone=TIMEZONE,
    tamanho_lote=50,
):
  """dict_coordenadas: {nome: (lat, lon)}.

  Retorna um DataFrame: índice = nomes, colunas = datas (datetime.date), valores
  = probabilidade média diária (0 a 1).
  """
  nomes = list(dict_coordenadas.keys())
  coords = list(dict_coordenadas.values())
  todas_datas = list(
      pd.date_range(data_inicio, data_fim, freq="D").date
  )
  resultado = pd.DataFrame(
      index=nomes, columns=todas_datas, dtype=float
  )

  for inicio in range(0, len(nomes), tamanho_lote):
    lote_nomes = nomes[inicio : inicio + tamanho_lote]
    lote_coords = coords[inicio : inicio + tamanho_lote]

    params = {
        "latitude": ",".join(str(c[0]) for c in lote_coords),
        "longitude": ",".join(str(c[1]) for c in lote_coords),
        "hourly": "precipitation_probability",
        "timezone": timezone,
        "start_date": str(data_inicio),
        "end_date": str(data_fim),
    }

    try:
      r = requests.get(
          "https://api.open-meteo.com/v1/forecast",
          params=params,
          timeout=60,
      )
      r.raise_for_status()
      dados = r.json()
    except requests.RequestException as e:
      print(f"AVISO: falha ao buscar previsão para o lote {lote_nomes}: {e}")
      continue

    if isinstance(dados, dict):
      dados = [dados]

    for nome, res in zip(lote_nomes, dados):
      if (
          "hourly" not in res
          or "precipitation_probability" not in res["hourly"]
      ):
        continue

      df_hourly = pd.DataFrame({
          "time": pd.to_datetime(res["hourly"]["time"]),
          "prob": res["hourly"]["precipitation_probability"],
      })
      df_hourly["prob"] = df_hourly["prob"] / 100.0  # Converte % para decimal (0 a 1)
      df_hourly["data"] = df_hourly["time"].dt.date

      # Média diária da probabilidade
      media_diaria = df_hourly.groupby("data")["prob"].mean()
      resultado.loc[nome] = media_diaria

  return resultado


# ==============================================================================
# EXECUÇÃO DA BUSCA PARA CIDADES E BAIRROS
# ==============================================================================

print("Buscando previsão de chuva para as cidades de SP...")
df_prob_cidades = buscar_probabilidade_chuva_media(
    coordenadas_cidades, DATA_INICIO, DATA_FIM
)

print("Buscando previsão de chuva para os bairros da Capital...")
df_prob_bairros = buscar_probabilidade_chuva_media(
    coordenadas_bairros, DATA_INICIO, DATA_FIM
)

print("Previsões obtidas com sucesso!")

## 6. Gerar a planilha Excel

Duas abas — **Cidades SP** e **SP Capital - Bairros** — cada uma com:
- cabeçalho duplo (dia da semana / data),
- valores em formato de porcentagem,
- escala de cores condicional (verde → amarelo → vermelho, de 0% a 100%).


import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# 1. PREPARAÇÃO DOS DADOS
# ==============================================================================

# Calcula a média do Estado de São Paulo (média diária entre todas as cidades)
sp_estado_media = pd.DataFrame(
    [df_prob_cidades.mean(axis=0)],
    index=["Estado de São Paulo"],
    columns=df_prob_cidades.columns,
)

# Consolida a tabela de cidades com o Estado no topo
df_cidades_completo = pd.concat([sp_estado_media, df_prob_cidades])

# ==============================================================================
# 2. CRIAÇÃO E ESTILIZAÇÃO DO WORKBOOK
# ==============================================================================

wb = Workbook()

# Estilos reutilizáveis
font_header_dias = Font(name="Calibri", size=11, bold=True, color="000000")
font_header_datas = Font(name="Calibri", size=11, bold=True, color="000000")
font_estado = Font(name="Calibri", size=11, bold=True, color="000000")
font_normal = Font(name="Calibri", size=11, bold=False, color="000000")

fill_header = PatternFill(
    start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"
)
fill_estado = PatternFill(
    start_color="E6EDF5", end_color="E6EDF5", fill_type="solid"
)  # Destaque leve para o Estado

align_center = Alignment(horizontal="center", vertical="center")
align_left = Alignment(horizontal="left", vertical="center")

thin_border = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)

# Mapeamento de dias da semana em português (3 letras)
dias_semana_pt = {
    "Mon": "Seg",
    "Tue": "Ter",
    "Wed": "Qua",
    "Thu": "Qui",
    "Fri": "Sex",
    "Sat": "Sáb",
    "Sun": "Dom",
}


def criar_aba_formatada(ws, df_dados, titulo_coluna_principal, eh_aba_cidades):
    ws.views.sheetView[0].showGridLines = True

    # --------------------------------------------------------------------------
    # Linha 1: Cabeçalho com Dias da Semana
    # --------------------------------------------------------------------------
    row1 = [titulo_coluna_principal]
    for d in df_dados.columns:
        dia_ing = d.strftime("%a")
        row1.append(dias_semana_pt.get(dia_ing, dia_ing))
    ws.append(row1)

    # --------------------------------------------------------------------------
    # Linha 2: Cabeçalho com Datas (ex: Set-01, Set-02)
    # --------------------------------------------------------------------------
    row2 = ["Cidade" if eh_aba_cidades else "Bairro"]
    for d in df_dados.columns:
        # Formata data como Mmm-DD (ex: Set-01, Ago-28)
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
        str_data = f"{meses_pt[d.month]}-{d.day:02d}"
        row2.append(str_data)
    ws.append(row2)

    # --------------------------------------------------------------------------
    # Linhas de Dados
    # --------------------------------------------------------------------------
    for localidade, row_vals in df_dados.iterrows():
        # Se for a seção de Cidades, insere a separação visual "Cidades" antes da lista
        if eh_aba_cidades and localidade == df_prob_cidades.index[0]:
            linha_subsep = ["Cidades"] + [
                f"{meses_pt[d.month]}-{d.day:02d}" for d in df_dados.columns
            ]
            ws.append(linha_subsep)

        linha = [localidade] + list(row_vals.values)
        ws.append(linha)

    # --------------------------------------------------------------------------
    # Formatação de Células, Estilos e Máscaras de Porcentagem
    # --------------------------------------------------------------------------
    max_row = ws.max_row
    max_col = ws.max_column

    for r in range(1, max_row + 1):
        primeira_celula = ws.cell(r, 1).value

        for c in range(1, max_col + 1):
            cell = ws.cell(r, c)
            cell.border = thin_border

            # Cabeçalhos
            if r in (1, 2) or primeira_celula == "Cidades":
                cell.font = font_header_datas
                cell.fill = fill_header
                cell.alignment = align_left if c == 1 else align_center

            # Linha consolidada do Estado de São Paulo
            elif primeira_celula == "Estado de São Paulo":
                cell.font = font_estado
                cell.fill = fill_estado
                if c == 1:
                    cell.alignment = align_left
                else:
                    cell.alignment = align_center
                    cell.number_format = "0%"

            # Linhas normais de dados (Cidades / Bairros)
            else:
                if c == 1:
                    cell.font = font_normal
                    cell.alignment = align_left
                else:
                    cell.font = font_normal
                    cell.alignment = align_center
                    cell.number_format = "0%"

    # --------------------------------------------------------------------------
    # Regra de Formatação Condicional (Escala de Cores: Verde -> Amarelo -> Vermelho)
    # --------------------------------------------------------------------------
    color_scale = ColorScaleRule(
        start_type="num",
        start_value=0.0,
        start_color="63BE7B",  # Verde (baixa probabilidade)
        mid_type="num",
        mid_value=0.5,
        mid_color="FFEB84",  # Amarelo (média probabilidade)
        end_type="num",
        end_value=1.0,
        end_color="F8696B",  # Vermelho (alta probabilidade)
    )

    col_fim_letra = get_column_letter(max_col)
    # Aplica a escala de cores apenas nas células com valores numéricos (%)
    ws.conditional_formatting.add(f"B3:{col_fim_letra}{max_row}", color_scale)

    # Ajusta largura das colunas
    ws.column_dimensions["A"].width = 28
    for c in range(2, max_col + 1):
        ws.column_dimensions[get_column_letter(c)].width = 10


# ------------------------------------------------------------------------------
# Montagem das Abas
# ------------------------------------------------------------------------------

# Aba 1: Cidades SP (com linha do Estado no topo)
ws1 = wb.active
ws1.title = "Cidades SP"
criar_aba_formatada(
    ws1, df_cidades_completo, titulo_coluna_principal="Dia", eh_aba_cidades=True
)

# Aba 2: SP Capital - Bairros
ws2 = wb.create_sheet(title="SP Capital - Bairros")
criar_aba_formatada(
    ws2,
    df_prob_bairros,
    titulo_coluna_principal="Dia",
    eh_aba_cidades=False,
)

# Salva o arquivo gerado
wb.save(ARQUIVO_SAIDA)
print(f"Planilha gerada com sucesso: '{ARQUIVO_SAIDA}'")

## 7. Baixar o arquivo

try:
    from google.colab import files
    files.download(ARQUIVO_SAIDA)
except ImportError:
    print(f"Ambiente fora do Google Colab — o arquivo foi salvo localmente como '{ARQUIVO_SAIDA}'.")
