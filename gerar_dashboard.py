"""
Gerador de Dashboard de Economia — Grid Energia

Lê uma planilha .xlsx no formato padrão, transforma os dados e gera um HTML
autossuficiente a partir do template `dashboard_template.html`.

Uso:
    python gerar_dashboard.py "Nome do Cliente" base_cliente.xlsx [arquivo_saida.html]

Formato esperado da planilha (header na linha 1, 11 colunas):
    Consumidor | Instalação | Referência | Consumo | Energia Compensada |
    Saldo | Vencimento | Valor | Economia Mensal | Valor Cemig | Status

Regras de transformação:
    - Descarta linhas com Status == "Aguardando Boleto" (faturamento ainda em avaliação)
    - Descarta colunas Consumidor e Status (não usadas na v1)
    - Calcula Economia Percentual = Economia Mensal / Valor Cemig
    - Descarta linhas sem Instalação
    - Descarta o último mês se tiver <= 30% da média de registros dos meses anteriores
    - Período padrão filtrado: janeiro do ano do último mês até o último mês com dados
"""

import sys
import re
import json
import unicodedata
from pathlib import Path
import pandas as pd


TEMPLATE_PATH = Path(__file__).parent / 'dashboard_template.html'
TEMPLATE_GRUPO_PATH = Path(__file__).parent / 'dashboard_grupo.html'
LIMIAR_DESCARTE = 0.30
STATUS_DESCARTAR = 'Aguardando Boleto'

COLUNAS_ESPERADAS = [
    'Consumidor', 'Instalação', 'Referência', 'Consumo', 'Energia Compensada',
    'Saldo', 'Vencimento', 'Valor', 'Economia Mensal', 'Valor Cemig', 'Status'
]

MESES_PT = {
    1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
    5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
    9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
}


def slugify(texto: str) -> str:
    texto = unicodedata.normalize('NFKD', texto).encode('ascii', 'ignore').decode('ascii')
    texto = re.sub(r'[^\w\s-]', '', texto).strip().lower()
    return re.sub(r'[-\s]+', '_', texto)


def fmt_mes_pt(ym: str) -> str:
    """'2026-04' -> 'Abril/2026'."""
    ano, mes = ym.split('-')
    return f"{MESES_PT[int(mes)]}/{ano}"


def ler_planilha(path_arquivo: Path) -> pd.DataFrame:
    ext = path_arquivo.suffix.lower()

    if ext in ('.xlsx', '.xlsm'):
        df = pd.read_excel(path_arquivo, header=0)
    elif ext == '.csv':
        # Detecta separador automaticamente (vírgula ou ponto e vírgula)
        with open(path_arquivo, 'r', encoding='utf-8-sig') as f:
            primeira_linha = f.readline()
        sep = ';' if primeira_linha.count(';') > primeira_linha.count(',') else ','
        df = pd.read_csv(path_arquivo, sep=sep, header=0, encoding='utf-8-sig')
        # Converter colunas de data (vêm como string no CSV)
        df['Referência'] = pd.to_datetime(df['Referência'], format='mixed', dayfirst=False)
        df['Vencimento'] = pd.to_datetime(df['Vencimento'], format='mixed', dayfirst=False, errors='coerce')
    else:
        raise ValueError(f"Formato não suportado: \'{ext}\'. Use .xlsx ou .csv")

    cols_faltando = set(COLUNAS_ESPERADAS) - set(df.columns)
    if cols_faltando:
        raise ValueError(
            f"Colunas faltando na planilha: {cols_faltando}\n"
            f"Esperado: {COLUNAS_ESPERADAS}\n"
            f"Recebido: {list(df.columns)}"
        )

    return df


def filtrar_aguardando_boleto(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """Remove linhas com status 'Aguardando Boleto'."""
    mask = df['Status'].astype(str).str.strip().str.lower() == STATUS_DESCARTAR.lower()
    qtd = int(mask.sum())
    return df[~mask].copy(), qtd


# ============================================================
# ROTEADOR INTELIGENTE — decide entre template Unitário e Grupo
# ============================================================

def extrair_nome_base(consumidor: str) -> str:
    """
    Extrai o "nome base" da razão social, removendo sufixos de filial
    e normalizando variações comuns de cadastro.

    Exemplos:
      "LOCALIZA RENT A CAR SA - ACUBA"     -> "localiza rent a car sa"
      "LOCALIZA RENT A CAR SA - ACPAT"     -> "localiza rent a car sa"
      "ARCELORMITTAL BRASIL S.A"           -> "arcelormittal brasil sa"
      "ARCELORMITTAL BRASIL S A"           -> "arcelormittal brasil sa"
      "ARCELORMITTAL BRASIL S/A"           -> "arcelormittal brasil sa"
      "ARCELORMITTAL BIOFLORESTAS LTDA."   -> "arcelormittal bioflorestas ltda"

    A regra é:
      1) Divide por " - " e fica com o trecho anterior ao primeiro hífen
         (que tipicamente carrega o nome da filial)
      2) Normaliza variações de sufixos jurídicos: S.A./S A/S/A -> SA;
         LTDA./EIRELI./ME./EPP. -> sem ponto final
      3) Colapsa múltiplos espaços
    """
    if not consumidor or pd.isna(consumidor):
        return ''
    nome = str(consumidor).strip()
    base = nome.split(' - ')[0].strip().lower()

    # Normaliza sufixos jurídicos comuns no final do nome.
    # A regex usa \b para casar apenas nessas palavras como tokens isolados,
    # não dentro de outras (ex.: não vai mexer em "MESA" por causa do "ME").
    SUFIXOS_JURIDICOS = r'\b(s[\s\./]*a|ltda|eireli|me|epp|mei|s/a|sa)\b\.?'
    def normaliza_sufixo(match):
        token = match.group(1).replace('.', '').replace('/', '').replace(' ', '')
        # 's/a', 's.a', 's a', 'sa' viram todos 'sa'
        if token in ('sa', 'sa', 'sa'):
            return 'sa'
        return token
    base = re.sub(SUFIXOS_JURIDICOS, normaliza_sufixo, base)

    # Colapsa múltiplos espaços em um só
    base = re.sub(r'\s+', ' ', base).strip()

    return base


def detectar_tipo_cliente(df: pd.DataFrame) -> tuple[str, list[str]]:
    """
    Decide entre 'unitario' e 'grupo' com base nos nomes-base distintos.
    Retorna (tipo, lista_de_nomes_base_unicos).
    """
    nomes_base = (
        df['Consumidor']
        .dropna()
        .map(extrair_nome_base)
        .replace('', pd.NA)
        .dropna()
        .unique()
        .tolist()
    )
    tipo = 'grupo' if len(nomes_base) > 1 else 'unitario'
    return tipo, sorted(nomes_base)


def descartar_ultimo_mes_se_incompleto(df: pd.DataFrame) -> tuple[pd.DataFrame, str | None]:
    df = df.copy()
    df['_ym'] = df['Referência'].dt.strftime('%Y-%m')
    contagem = df.groupby('_ym').size().sort_index()

    if len(contagem) < 2:
        return df.drop(columns=['_ym']), None

    ultimo_mes = contagem.index[-1]
    qtd_ultimo = contagem.iloc[-1]
    media_anteriores = contagem.iloc[:-1].mean()
    limiar = media_anteriores * LIMIAR_DESCARTE

    if qtd_ultimo <= limiar:
        df = df[df['_ym'] != ultimo_mes].copy()
        return df.drop(columns=['_ym']), ultimo_mes

    return df.drop(columns=['_ym']), None


def transformar_para_json(df: pd.DataFrame) -> list[dict]:
    df = df.dropna(subset=['Instalação']).copy()
    df['Instalação'] = df['Instalação'].apply(lambda x: str(int(x)))
    df['Referência'] = df['Referência'].dt.strftime('%Y-%m-%d')
    df['Vencimento'] = df['Vencimento'].apply(
        lambda x: x.strftime('%Y-%m-%d') if pd.notna(x) else ''
    )

    for col in ['Consumo', 'Energia Compensada', 'Saldo', 'Valor', 'Valor Cemig', 'Economia Mensal']:
        df[col] = df[col].astype(float)

    registros = []
    for _, r in df.iterrows():
        valor_cemig = float(r['Valor Cemig'])
        economia = float(r['Economia Mensal'])
        economia_pct = economia / valor_cemig if valor_cemig > 0 else 0.0

        # Consumidor e Status são preservados para a aba Inadimplência.
        # Strip remove espaços perdidos comuns em exports de CSV.
        consumidor = str(r['Consumidor']).strip() if pd.notna(r['Consumidor']) else ''
        status = str(r['Status']).strip() if pd.notna(r['Status']) else ''

        registros.append({
            'instalacao': r['Instalação'],
            'consumidor': consumidor,
            'referencia': r['Referência'],
            'consumo': float(r['Consumo']),
            'compensada': float(r['Energia Compensada']),
            'saldo': float(r['Saldo']),
            'vencimento': r['Vencimento'],
            'valorGrid': float(r['Valor']),
            'valorCemig': valor_cemig,
            'economia': economia,
            'economiaPct': economia_pct,
            'status': status
        })

    return registros


def calcular_periodos(registros: list[dict]) -> dict:
    meses = sorted({r['referencia'][:7] for r in registros})
    if not meses:
        raise ValueError("Nenhum registro válido após filtros — base vazia.")

    period_min = meses[0]
    period_max = meses[-1]
    ano_max = period_max[:4]

    default_ini = f"{ano_max}-01"
    if default_ini < period_min:
        default_ini = period_min
    default_fim = period_max

    subtitle = f"{fmt_mes_pt(period_min)} — {fmt_mes_pt(period_max)}"

    return {
        '__PERIOD_MIN__': period_min,
        '__PERIOD_MAX__': period_max,
        '__DEFAULT_INI__': default_ini,
        '__DEFAULT_FIM__': default_fim,
        '__SUBTITLE__': subtitle,
    }


def gerar_html(nome_cliente: str, registros: list[dict],
               periodos: dict, path_saida: Path,
               template_path: Path, nomes_consumidores: list[str] | None = None) -> None:
    template = template_path.read_text(encoding='utf-8')

    substituicoes = {
        '__CLIENTE_NOME__': nome_cliente.upper(),
        '__RAW_DATA__': json.dumps(registros, ensure_ascii=False, separators=(',', ':')),
        **periodos
    }

    # Marcador específico do template de grupo. Se o template for unitário,
    # ele simplesmente não terá esse marcador e a substituição é silenciosa.
    if '__CONSUMER_NAMES__' in template:
        substituicoes['__CONSUMER_NAMES__'] = json.dumps(
            nomes_consumidores or [], ensure_ascii=False, separators=(',', ':')
        )

    html = template
    for marker, valor in substituicoes.items():
        if marker not in html:
            raise RuntimeError(f"Marcador {marker} não encontrado no template.")
        html = html.replace(marker, valor)

    path_saida.write_text(html, encoding='utf-8')


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)

    nome_cliente = sys.argv[1]
    path_arquivo = Path(sys.argv[2])
    path_saida = Path(sys.argv[3]) if len(sys.argv) > 3 else \
        Path(f"dashboard_{slugify(nome_cliente)}.html")

    if not path_arquivo.exists():
        print(f"❌ Arquivo não encontrado: {path_arquivo}")
        sys.exit(1)

    if not TEMPLATE_PATH.exists():
        print(f"❌ Template não encontrado: {TEMPLATE_PATH}")
        sys.exit(1)

    print(f"📊 Lendo arquivo: {path_arquivo}")
    df = ler_planilha(path_arquivo)
    print(f"   {len(df):,} registros lidos")

    df, qtd_aguard = filtrar_aguardando_boleto(df)
    if qtd_aguard:
        print(f"⏸  {qtd_aguard} linhas descartadas (Status = 'Aguardando Boleto')")

    # ── Roteador Inteligente ──────────────────────────────────────────
    tipo_cliente, nomes_base = detectar_tipo_cliente(df)
    if tipo_cliente == 'grupo':
        print(f"🏢 Template de Grupo acionado — {len(nomes_base)} clientes distintos detectados:")
        for n in nomes_base:
            print(f"     • {n}")
        template_path = TEMPLATE_GRUPO_PATH
        if not template_path.exists():
            print(f"❌ Template de grupo não encontrado: {template_path}")
            sys.exit(1)
    else:
        print(f"👤 Cliente Unitário detectado: {nomes_base[0] if nomes_base else '(sem consumidor)'}")
        template_path = TEMPLATE_PATH

    df, mes_descartado = descartar_ultimo_mes_se_incompleto(df)
    if mes_descartado:
        print(f"⚠️  Mês incompleto descartado: {mes_descartado} (≤{int(LIMIAR_DESCARTE*100)}% da média)")

    registros = transformar_para_json(df)
    if not registros:
        print("❌ Nenhum registro válido após filtros — não é possível gerar dashboard.")
        sys.exit(1)

    periodos = calcular_periodos(registros)
    instalacoes = len({r['instalacao'] for r in registros})

    print(f"✓  Registros finais: {len(registros):,}")
    print(f"   Instalações únicas: {instalacoes}")
    print(f"   Período: {periodos['__PERIOD_MIN__']} → {periodos['__PERIOD_MAX__']}")
    print(f"   Filtro padrão: {periodos['__DEFAULT_INI__']} → {periodos['__DEFAULT_FIM__']}")
    print(f"📝 Cliente: {nome_cliente.upper()}")
    print(f"📄 Template: {template_path.name}")

    gerar_html(nome_cliente, registros, periodos, path_saida,
               template_path=template_path,
               nomes_consumidores=nomes_base if tipo_cliente == 'grupo' else None)
    print(f"✅ Dashboard gerado: {path_saida}  ({path_saida.stat().st_size:,} bytes)")


if __name__ == '__main__':
    main()
