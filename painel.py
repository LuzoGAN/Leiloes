#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Painel visual do leilão Copart (Streamlit)
==========================================
Como usar:
  pip install -r requirements.txt
  python -m playwright install chromium   (só na primeira vez)
  streamlit run painel.py

O que ele faz:
- Você cola o link do saleListResult na barra lateral e clica em "Extrair agora"
  ou "Iniciar ao vivo".
- O painel chama o copart_leilao.py em segundo plano (navegador real faz a
  extração) e mostra a tabela atualizando na tela: item do topo em destaque,
  métricas, filtros e log do monitoramento.
- O Excel continua sendo salvo normalmente a cada atualização.
"""

import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

BASE = Path(__file__).parent
SCRIPT = BASE / "copart_leilao.py"

URL_EXEMPLO = (
    "https://www.copart.com.br/saleListResult/auctionId/10490"
    "?location=Goi%C3%A2nia%20-%20GO&saleDate=1791212400000&liveAuction=true"
    "&from=&yardNum=9549&auctioneer=Ygor%20Ferreira%20Brasil%20-%20Juceg%20074%2F20"
)

st.set_page_config(page_title="Copart - Painel do Leilão", layout="wide")

if "proc" not in st.session_state:
    st.session_state.proc = None
if "saida_atual" not in st.session_state:
    st.session_state.saida_atual = str(BASE / "leilao_ao_vivo.xlsx")


def proc_rodando() -> bool:
    p = st.session_state.proc
    return bool(p is not None and p.poll() is None)


def parar_processo():
    p = st.session_state.proc
    if p is not None and p.poll() is None:
        p.terminate()
        try:
            p.wait(timeout=10)
        except Exception:
            p.kill()
    st.session_state.proc = None


@st.cache_data(ttl=5, show_spinner=False)
def carregar_excel(arquivo: str, _versao: int):
    """Lê o Excel gerado pelo copart_leilao.py. _versao força recarregar."""
    xls = pd.ExcelFile(arquivo)
    df = pd.read_excel(xls, sheet_name="Leilao")
    log = None
    if "Monitoramento_AoVivo" in xls.sheet_names:
        try:
            log = pd.read_excel(xls, sheet_name="Monitoramento_AoVivo")
        except Exception:
            log = None
    return df, log


# --------------------------------------------------------------- Sidebar
st.sidebar.title("Leilão Copart")
url = st.sidebar.text_input("Link do leilão (saleListResult)", value=URL_EXEMPLO)
saida = st.sidebar.text_input("Arquivo Excel", value=st.session_state.saida_atual)
st.session_state.saida_atual = saida
intervalo = st.sidebar.slider("Refresh do ao vivo (segundos)", 5, 120, 15)
max_paginas = st.sidebar.number_input("Máx. páginas na extração", 1, 50, 5)
auto = st.sidebar.checkbox("Atualizar tela automaticamente", value=True)
auto_seg = st.sidebar.slider("Atualizar tela a cada (s)", 5, 120, 15)

col_a, col_b = st.sidebar.columns(2)
with col_a:
    btn_extrair = st.button("Extrair agora", use_container_width=True)
with col_b:
    if proc_rodando():
        btn_parar = st.button("Parar", use_container_width=True)
        btn_iniciar = False
    else:
        btn_iniciar = st.button("Iniciar ao vivo", use_container_width=True)
        btn_parar = False

if btn_extrair:
    parar_processo()
    with st.spinner("Extraindo do Copart (navegador real, pode levar 1-2 min)..."):
        r = subprocess.run(
            [sys.executable, str(SCRIPT), "--url", url, "--saida", saida,
             "--headless", "--max-paginas", str(int(max_paginas))],
            capture_output=True, text=True, timeout=420,
        )
    st.session_state["_log_extracao"] = (r.stdout or "")[-3000:] + "\n" + (r.stderr or "")[-2000:]
    if r.returncode != 0:
        st.error("Extração falhou. Veja o detalhe abaixo.")
    else:
        st.success("Extração concluída.")
        st.cache_data.clear()

if btn_iniciar:
    parar_processo()
    log_path = BASE / "monitor_ao_vivo.log"
    fh = open(log_path, "a", encoding="utf-8")
    fh.write(f"\n===== inicio {datetime.now():%d/%m/%Y %H:%M:%S} url={url} =====\n")
    fh.flush()
    st.session_state.proc = subprocess.Popen(
        [sys.executable, str(SCRIPT), "--url", url, "--saida", saida,
         "--headless", "--max-paginas", str(int(max_paginas)),
         "--monitorar", "--intervalo", str(int(intervalo))],
        stdout=fh, stderr=subprocess.STDOUT,
    )
    st.success("Monitoramento ao vivo iniciado em segundo plano.")
    st.cache_data.clear()

if btn_parar:
    parar_processo()
    st.warning("Monitoramento parado. O Excel já está salvo.")

# ------------------------------------------------------------------ Corpo
st.title("Painel do Leilão — Copart")

if proc_rodando():
    st.success("● AO VIVO — monitorando o topo da lista (refresh automático no site).")
else:
    st.info("Parado. Clique em **Extrair agora** (foto do momento) ou **Iniciar ao vivo** (acompanha o pregão).")

if st.session_state.get("_log_extracao"):
    with st.expander("Ver saída da última extração"):
        st.code(st.session_state["_log_extracao"][-4000:])

arq = Path(saida)
if not arq.exists():
    arq_base = BASE / saida
    if arq_base.exists():
        arq = arq_base

if not arq.exists():
    st.warning("Nenhum Excel ainda. Cole o link e clique em **Extrair agora**.")
    st.stop()

versao = int(time.time() // max(auto_seg, 5))
try:
    df, log = carregar_excel(str(arq), versao)
except Exception as e:
    st.error(f"Não consegui ler o Excel (ele pode estar sendo gravado agora): {e}")
    st.stop()

st.caption(f"Arquivo: `{arq}` — atualizado em {datetime.fromtimestamp(arq.stat().st_mtime):%d/%m/%Y %H:%M:%S}")

# Métricas
total = len(df)
col_status = "Status" if "Status" in df.columns else None
col_arrem = "Valor Arrematado (R$)" if "Valor Arrematado (R$)" in df.columns else None
n_pregao = 0
n_arrem = 0
if col_status is not None:
    s = df[col_status].astype(str).str.upper()
    n_pregao = int(s.str.contains("PREGAO|AO VIVO", na=False).sum())
if col_arrem is not None:
    n_arrem = int(df[col_arrem].notna().sum())

m1, m2, m3 = st.columns(3)
m1.metric("Itens na lista", total)
m2.metric("Em pregão / ao vivo", n_pregao)
m3.metric("Com valor final", n_arrem)

# Destaque do topo (item em pregão)
st.subheader("Item do topo — em pregão agora")
topo = df.iloc[0].to_dict() if total else {}
c1, c2, c3, c4 = st.columns(4)
c1.metric("Lote", str(topo.get("Codigo / Lote", "-")))
c2.metric("Veículo", str(topo.get("Titulo (Marca + Modelo)", "-")))
c3.metric("Lance atual (R$)", str(topo.get("Lance Atual (R$)", "-")))
c4.metric("FIPE (R$)", str(topo.get("Valor FIPE (R$)", "-")))
st.write(
    f"Documento: **{topo.get('Documento', '-')}** · Categoria: **{topo.get('Categoria', '-')}** · "
    f"Condição: **{topo.get('Condicao', '-')}** · Pátio: **{topo.get('Patio Veiculo', '-')}** · "
    f"Lote/Vaga: **{topo.get('Lote/Vaga', '-')}**"
)

# Filtros + tabela
st.subheader("Todos os itens")
f1, f2, f3 = st.columns(3)
busca = f1.text_input("Buscar (marca, modelo, lote)")
marcas = sorted(d for d in df["Marca"].dropna().astype(str).unique()) if "Marca" in df.columns else []
f_marca = f2.multiselect("Marca", marcas)
status_vals = sorted(d for d in df["Status"].dropna().astype(str).unique()) if col_status else []
f_status = f3.multiselect("Status", status_vals)

vista = df.copy()
if busca:
    b = busca.lower()
    vista = vista[vista.apply(lambda r: b in str(r.values).lower(), axis=1)]
if f_marca and "Marca" in vista.columns:
    vista = vista[vista["Marca"].astype(str).isin(f_marca)]
if f_status and col_status:
    vista = vista[vista[col_status].astype(str).isin(f_status)]

st.dataframe(vista, use_container_width=True, height=450)

# Download
with open(arq, "rb") as fh:
    st.download_button("Baixar Excel", fh.read(), file_name=Path(arq).name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Log do ao vivo
if log is not None and len(log):
    st.subheader("Log do monitoramento")
    st.dataframe(log.tail(50), use_container_width=True)

# Auto-refresh da tela (recarrega o Excel)
if auto and not st.session_state.get("_stop_refresh", False):
    time.sleep(auto_seg)
    st.cache_data.clear()
    st.rerun()
