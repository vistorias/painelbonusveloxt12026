# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import json
from pathlib import Path
import unicodedata, re

# ===================== CONFIG BÁSICA =====================
st.set_page_config(page_title="Painel de Bônus - VELOX", layout="wide")
st.title("🚀 Painel de Bônus Trimestral - VELOX")

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

# ===================== HELPERS =====================
def norm_txt(s: str) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    s = str(s).strip().upper()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = re.sub(r"\s+", " ", s)
    return s

def up(s):
    return norm_txt(s)

def texto_obs(valor):
    if pd.isna(valor):
        return ""
    s = str(valor).strip()
    return "" if s.lower() in ["none", "nan", ""] else s

def pct_safe(x):
    """Converte 0.035 ou 3.5 para fração (0-1)."""
    try:
        x = float(x)
        if x > 1:
            return x / 100.0
        return x
    except Exception:
        return 0.0

def fmt_pct(x):
    try:
        return f"{float(x) * 100:.2f}%"
    except Exception:
        return "0.00%"

def is_org_loja(item: str) -> bool:
    k = up(item)
    return "ORGANIZACAO DA LOJA" in k or ("ORGANIZACAO" in k and "LOJA" in k)

def is_lider_org(item: str) -> bool:
    k = up(item)
    return "LIDERANCA" in k and "ORGANIZACAO" in k

def is_producao_item(item: str) -> bool:
    k = up(item)
    # cobre PRODUÇÃO/PRODUCAO e variações abreviadas
    return ("PRODU" in k) or k.startswith("PRD") or k.startswith("PROD")

# ===================== PARÂMETROS (QUALIDADE GESTÃO) =====================
# Mesmo método da LOG (por cidade): cada cidade pesa igual no grupo do responsável.
QUALIDADE_GESTAO_METODO = "por_cidade"
META_ERROS_TOTAIS_GESTAO = 0.035  # 3,5%
META_ERROS_GG_GESTAO = 0.015      # 1,5%

# ===================== RESPONSABILIDADE (CIDADES) =====================
# Supervisores e Gerentes seguem o mesmo grupo de cidades:
# - ARYSON e MOISÉS: São Luís, Pedreiras, Grajaú (3 cidades)
# - LUCAS e JORGE: Imperatriz, Estreito (2 cidades)
# Pesos abaixo são usados para RATEAR PRODUÇÃO quando o item no pesos.json for "Produção" genérico.
_SUPERVISORES_CIDADES_RAW = {
    "ARYSON PAULINELLE GUTERES COSTA": {
        "SÃO LUIS": 1/3,
        "PEDREIRAS": 1/3,
        "GRAJAÚ": 1/3
    },
    "MOISÉS SANTOS DO NASCIMENTO": {
        "SÃO LUIS": 1/3,
        "PEDREIRAS": 1/3,
        "GRAJAÚ": 1/3
    },
    "LUCAS SAMPAIO NEVES": {
        "IMPERATRIZ": 0.5,
        "ESTREITO": 0.5
    },
    "JORGE ALEXANDRE BEZERRA DA COSTA": {
        "IMPERATRIZ": 0.5,
        "ESTREITO": 0.5
    }
}
RESP_CIDADES = {
    up(nome): {up(cidade): float(peso) for cidade, peso in cidades.items()}
    for nome, cidades in _SUPERVISORES_CIDADES_RAW.items()
}

# ===================== CARREGAMENTO ======================
def load_json(path: Path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

try:
    PESOS = load_json(DATA_DIR / "pesos_velox.json")
    INDICADORES = load_json(DATA_DIR / "empresa_indicadores_velox.json")
except Exception as e:
    st.error(f"Erro ao carregar JSONs em data/: {e}")
    st.stop()

FUNCOES_EXCLUIDAS = set(up(x) for x in PESOS.get("_funcoes_excluidas", []))

MESES = ["TRIMESTRE", "JANEIRO", "FEVEREIRO", "MARÇO"]
filtro_mes = st.radio("📅 Selecione o mês:", MESES, horizontal=True)

def ler_planilha(mes: str) -> pd.DataFrame:
    base = DATA_DIR / "RESUMO PARA PAINEL - VELOX.xlsx"
    if base.exists():
        return pd.read_excel(base, sheet_name=mes)
    candidatos = list(DATA_DIR.glob("RESUMO PARA PAINEL - VELOX*.xls*"))
    if not candidatos:
        st.error("Planilha não encontrada em data/ (RESUMO PARA PAINEL - VELOX.xlsx)")
        st.stop()
    return pd.read_excel(sorted(candidatos)[0], sheet_name=mes)

# ===================== QUALIDADE (GESTÃO) =====================
def calc_qualidade_gestao_por_cidade(
    cidades_resp: list,
    total_por_cidade: dict,
    gg_por_cidade: dict,
    meta_total: float = META_ERROS_TOTAIS_GESTAO,
    meta_gg: float = META_ERROS_GG_GESTAO
):
    """
    Retorna (frac_total, frac_gg, detalhes)
    frac_total = proporção de cidades que bateram Erros Totais (0..1)
    frac_gg    = proporção de cidades que bateram Erros GG (0..1)
    """
    detalhes = []

    cidades_total = [c for c in cidades_resp if c in total_por_cidade]
    cidades_gg = [c for c in cidades_resp if c in gg_por_cidade]

    if not cidades_total and not cidades_gg:
        # sem dados por cidade -> não paga qualidade da gestão (evita pagar no escuro)
        return 0.0, 0.0, ["Qualidade (gestão) — sem dados por cidade no JSON do mês"]

    # Erros Totais
    if cidades_total:
        ok_total = [c for c in cidades_total if float(total_por_cidade[c]) <= meta_total]
        nok_total = [c for c in cidades_total if c not in ok_total]
        frac_total = len(ok_total) / len(cidades_total)
        if nok_total:
            detalhes.append("Qualidade — Erros Totais (não bateu): " + ", ".join([c.title() for c in nok_total]))
    else:
        frac_total = 0.0
        detalhes.append("Qualidade — Erros Totais: sem dados por cidade")

    # Erros GG
    if cidades_gg:
        ok_gg = [c for c in cidades_gg if float(gg_por_cidade[c]) <= meta_gg]
        nok_gg = [c for c in cidades_gg if c not in ok_gg]
        frac_gg = len(ok_gg) / len(cidades_gg)
        if nok_gg:
            detalhes.append("Qualidade — Erros GG (não bateu): " + ", ".join([c.title() for c in nok_gg]))
    else:
        frac_gg = 0.0
        detalhes.append("Qualidade — Erros GG: sem dados por cidade")

    return float(frac_total), float(frac_gg), detalhes

# ===================== ELEGIBILIDADE =====================
def elegivel(valor_meta, obs):
    obs_u = up(obs)
    try:
        vm = float(valor_meta) if pd.notna(valor_meta) else 0.0
    except Exception:
        vm = 0.0
    if pd.isna(valor_meta) or vm == 0:
        return False, "Sem elegibilidade no mês"
    if "LICEN" in obs_u:
        return False, "Licença no mês"
    return True, ""

# ===================== CÁLCULO (POR MÊS) =====================
def calcula_mes(df_mes: pd.DataFrame, nome_mes: str) -> pd.DataFrame:
    ind_mes_raw = INDICADORES.get(nome_mes, {})

    # flags gerais
    ind_flags = {up(k): v for k, v in ind_mes_raw.items() if k != "producao_por_cidade"}

    def flag(chave: str, default=True):
        return bool(ind_flags.get(up(chave), default))

    # produção por cidade (boolean bateu)
    prod_cid_norm = {up(k): bool(v) for k, v in ind_mes_raw.get("producao_por_cidade", {}).items()}

    # qualidade por cidade (percentuais)
    qual_total_cid_norm = {up(k): pct_safe(v) for k, v in ind_mes_raw.get("qualidade_total_por_cidade", {}).items()}
    qual_gg_cid_norm = {up(k): pct_safe(v) for k, v in ind_mes_raw.get("qualidade_gg_por_cidade", {}).items()}

    df = df_mes.copy()

    # remove funções excluídas (se existirem na base)
    if "FUNÇÃO" in df.columns:
        df = df[~df["FUNÇÃO"].astype(str).apply(up).isin(FUNCOES_EXCLUIDAS)]

    def calcula_recebido(row):
        func = up(row.get("FUNÇÃO", ""))
        cidade = up(row.get("CIDADE", ""))
        nome = up(row.get("NOME", ""))
        obs = row.get("OBSERVAÇÃO", "")
        valor_meta = row.get("VALOR MENSAL META", 0)

        ok, motivo = elegivel(valor_meta, obs)
        perdeu_itens = []

        if not ok:
            return pd.Series({
                "MES": nome_mes, "META": 0.0, "RECEBIDO": 0.0, "PERDA": 0.0, "%": 0.0,
                "_badge": motivo, "_obs": texto_obs(obs), "perdeu_itens": perdeu_itens
            })

        metainfo = PESOS.get(func, PESOS.get(row.get("FUNÇÃO", ""), {}))
        total_func = float(metainfo.get("total", float(valor_meta) if pd.notna(valor_meta) else 0.0))
        itens = metainfo.get("metas", {})

        recebido, perdas = 0.0, 0.0

        for item, peso in itens.items():
            parcela = total_func * float(peso)
            item_norm = up(item)

            # ------------------- PRODUÇÃO -------------------
            if is_producao_item(item):
                # Supervisor/Gerente: rateia por cidades sob responsabilidade
                if func in [up("SUPERVISOR"), up("GERENTE")] and nome in RESP_CIDADES:
                    base_soma = sum(RESP_CIDADES[nome].values()) or 1.0
                    perdas_cids = []

                    for cid_norm, w in RESP_CIDADES[nome].items():
                        bateu = prod_cid_norm.get(cid_norm, True)  # se não existir no JSON, assume bateu
                        fatia = parcela * (float(w) / base_soma)

                        if bateu:
                            recebido += fatia
                        else:
                            perdas += fatia
                            perdas_cids.append(cid_norm.title())

                    if perdas_cids:
                        perdeu_itens.append("Produção – " + ", ".join(perdas_cids))
                    continue

                # demais funções: produção da própria cidade
                bateu_prod = prod_cid_norm.get(cidade, True)
                if bateu_prod:
                    recebido += parcela
                else:
                    perdas += parcela
                    perdeu_itens.append("Produção – " + (row.get("CIDADE", "") or "Cidade não informada"))
                continue

            # ------------------- QUALIDADE -------------------
            if item_norm == up("QUALIDADE"):
                # Supervisor/Gerente: 20% dividido em 10% (Totais) + 10% (GG), por cidade
                if func in [up("SUPERVISOR"), up("GERENTE")] and nome in RESP_CIDADES:
                    cidades_resp = list(RESP_CIDADES[nome].keys())

                    frac_total, frac_gg, detalhes = calc_qualidade_gestao_por_cidade(
                        cidades_resp=cidades_resp,
                        total_por_cidade=qual_total_cid_norm,
                        gg_por_cidade=qual_gg_cid_norm
                    )

                    metade = parcela * 0.5  # 10% + 10%

                    # Totais
                    recebido += metade * float(frac_total)
                    perdas += metade * (1.0 - float(frac_total))

                    # GG
                    recebido += metade * float(frac_gg)
                    perdas += metade * (1.0 - float(frac_gg))

                    if float(frac_total) < 1.0 or float(frac_gg) < 1.0:
                        perdeu_itens.append("Qualidade (gestão)")
                        perdeu_itens.extend(detalhes)
                    continue

                # fallback (se não mapeado): usa flag geral "qualidade"
                if flag("qualidade", True):
                    recebido += parcela
                else:
                    perdas += parcela
                    perdeu_itens.append("Qualidade")
                continue

            # ------------------- ORGANIZAÇÃO DA LOJA -------------------
            if is_org_loja(item):
                if flag("organizacao_da_loja", True):
                    recebido += parcela
                else:
                    perdas += parcela
                    perdeu_itens.append("Organização da Loja")
                continue

            # ------------------- LIDERANÇA & ORGANIZAÇÃO -------------------
            if is_lider_org(item) or item_norm == up("LIDERANÇA & ORGANIZAÇÃO"):
                if flag("Liderança & Organização", True):
                    recebido += parcela
                else:
                    perdas += parcela
                    perdeu_itens.append("Liderança & Organização")
                continue

            # ------------------- OUTRAS METAS -------------------
            # Se no futuro você quiser controlar via JSON, basta criar chaves no indicadores:
            # "pesquisa_de_satisfacao", "recursos_humanos", "treinamento", "adesao_pesquisa_de_clima", etc.
            chave_sugerida = item_norm.replace("&", "E")
            if flag(chave_sugerida, True):
                recebido += parcela
            else:
                perdas += parcela
                perdeu_itens.append(str(item))

        meta = total_func
        perc = 0.0 if meta == 0 else (recebido / meta) * 100.0

        return pd.Series({
            "MES": nome_mes, "META": meta, "RECEBIDO": recebido, "PERDA": perdas,
            "%": perc, "_badge": "", "_obs": texto_obs(obs), "perdeu_itens": perdeu_itens
        })

    calc = df.apply(calcula_recebido, axis=1)
    return pd.concat([df.reset_index(drop=True), calc], axis=1)

# ===================== LEITURA (TRIMESTRE OU MÊS) =====================
if filtro_mes == "TRIMESTRE":
    try:
        df_jan, df_fev, df_mar = [ler_planilha(m) for m in ["JANEIRO", "FEVEREIRO", "MARÇO"]]
        st.success("✅ Planilhas carregadas: JANEIRO, FEVEREIRO e MARÇO!")
    except Exception as e:
        st.error(f"Erro ao ler a planilha: {e}")
        st.stop()

    dados_full = pd.concat([
        calcula_mes(df_jan, "JANEIRO"),
        calcula_mes(df_fev, "FEVEREIRO"),
        calcula_mes(df_mar, "MARÇO")
    ], ignore_index=True)

    group_cols = ["CIDADE", "NOME", "FUNÇÃO", "DATA DE ADMISSÃO", "TEMPO DE CASA"]
    # se alguma coluna não existir, remove do group
    group_cols = [c for c in group_cols if c in dados_full.columns]

    agg = (dados_full
           .groupby(group_cols, dropna=False)
           .agg({
               "META": "sum",
               "RECEBIDO": "sum",
               "PERDA": "sum",
               "_obs": lambda x: ", ".join(sorted({s for s in x if s})),
               "_badge": lambda x: " / ".join(sorted({s for s in x if s}))
           })
           .reset_index())

    agg["%"] = agg.apply(lambda r: 0.0 if r["META"] == 0 else (r["RECEBIDO"] / r["META"]) * 100.0, axis=1)

    perdas_pessoa = (
        dados_full.assign(_lost=lambda d: d.apply(
            lambda r: [f"{it} ({r['MES']})" for it in r["perdeu_itens"]],
            axis=1))
        .groupby(group_cols, dropna=False)["_lost"]
        .sum()
        .apply(lambda L: ", ".join(sorted(set(L))))
        .reset_index()
        .rename(columns={"_lost": "INDICADORES_NAO_ENTREGUES"})
    )

    dados_calc = agg.merge(perdas_pessoa, on=group_cols, how="left")
    dados_calc["INDICADORES_NAO_ENTREGUES"] = dados_calc["INDICADORES_NAO_ENTREGUES"].fillna("")
else:
    try:
        df_mes = ler_planilha(filtro_mes)
        st.success(f"✅ Planilha de {filtro_mes} carregada!")
    except Exception as e:
        st.error(f"Erro ao ler a planilha: {e}")
        st.stop()

    dados_calc = calcula_mes(df_mes, filtro_mes)
    dados_calc["INDICADORES_NAO_ENTREGUES"] = dados_calc["perdeu_itens"].apply(
        lambda L: ", ".join(L) if isinstance(L, list) and L else ""
    )

# ===================== FILTROS =====================
st.markdown("### 🔎 Filtros")
col1, col2, col3, col4 = st.columns(4)

with col1:
    filtro_nome = st.text_input("Buscar por nome (contém)", "")

with col2:
    funcoes_validas = sorted([f for f in dados_calc.get("FUNÇÃO", pd.Series(dtype=str)).dropna().unique()
                             if up(f) in PESOS.keys() and up(f) not in FUNCOES_EXCLUIDAS])
    filtro_funcao = st.selectbox("Função", ["Todas"] + funcoes_validas)

with col3:
    cidades = ["Todas"] + sorted(dados_calc.get("CIDADE", pd.Series(dtype=str)).dropna().unique())
    filtro_cidade = st.selectbox("Cidade", cidades)

with col4:
    tempos = ["Todos"] + sorted(dados_calc.get("TEMPO DE CASA", pd.Series(dtype=str)).dropna().unique())
    filtro_tempo = st.selectbox("Tempo de casa", tempos)

dados_view = dados_calc.copy()
if filtro_nome and "NOME" in dados_view.columns:
    dados_view = dados_view[dados_view["NOME"].astype(str).str.contains(filtro_nome, case=False, na=False)]
if filtro_funcao != "Todas" and "FUNÇÃO" in dados_view.columns:
    dados_view = dados_view[dados_view["FUNÇÃO"] == filtro_funcao]
if filtro_cidade != "Todas" and "CIDADE" in dados_view.columns:
    dados_view = dados_view[dados_view["CIDADE"] == filtro_cidade]
if filtro_tempo != "Todos" and "TEMPO DE CASA" in dados_view.columns:
    dados_view = dados_view[dados_view["TEMPO DE CASA"] == filtro_tempo]

# ===================== RESUMO =====================
st.markdown("### 📊 Resumo Geral")
colA, colB, colC = st.columns(3)
with colA:
    st.success(f"💰 Total possível: R$ {dados_view['META'].sum():,.2f}")
with colB:
    st.info(f"📈 Recebido: R$ {dados_view['RECEBIDO'].sum():,.2f}")
with colC:
    st.error(f"📉 Deixou de ganhar: R$ {dados_view['PERDA'].sum():,.2f}")

# ===================== CARDS =====================
st.markdown("### 👥 Colaboradores")
cols = st.columns(3)

dados_view = dados_view.sort_values(by="%", ascending=False)

for idx, row in dados_view.iterrows():
    pct = float(row.get("%", 0.0)) if pd.notna(row.get("%", 0.0)) else 0.0
    meta = float(row.get("META", 0.0)) if pd.notna(row.get("META", 0.0)) else 0.0
    recebido = float(row.get("RECEBIDO", 0.0)) if pd.notna(row.get("RECEBIDO", 0.0)) else 0.0
    perdido = float(row.get("PERDA", 0.0)) if pd.notna(row.get("PERDA", 0.0)) else 0.0
    badge = row.get("_badge", "")
    obs_txt = texto_obs(row.get("_obs", ""))
    perdidos_txt = texto_obs(row.get("INDICADORES_NAO_ENTREGUES", ""))

    bg = "#f9f9f9" if not badge else "#eeeeee"

    with cols[idx % 3]:
        st.markdown(f"""
        <div style="border:1px solid #ccc;padding:16px;border-radius:12px;margin-bottom:12px;background:{bg}">
            <h4 style="margin:0">{str(row.get('NOME','')).title()}</h4>
            <p style="margin:4px 0;"><strong>{row.get('FUNÇÃO','')}</strong> — {row.get('CIDADE','')}</p>
            <p style="margin:4px 0;">
                <strong>Meta {'Trimestral' if filtro_mes=='TRIMESTRE' else 'Mensal'}:</strong> R$ {meta:,.2f}<br>
                <strong>Recebido:</strong> R$ {recebido:,.2f}<br>
                <strong>Deixou de ganhar:</strong> R$ {perdido:,.2f}<br>
                <strong>Cumprimento:</strong> {pct:.1f}%
            </p>
            <div style="height: 10px; background: #ddd; border-radius: 5px; overflow: hidden;">
                <div style="width: {max(0.0, min(100.0, pct)):.1f}%; background: black; height: 100%;"></div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        if badge:
            st.caption(f"⚠️ {badge}")
        if obs_txt:
            st.caption(f"🗒️ {obs_txt}")
        if perdidos_txt and "100%" not in perdidos_txt:
            st.caption(f"🔻 Indicadores não entregues: {perdidos_txt}")