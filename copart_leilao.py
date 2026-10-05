#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Monitor de Leilão Copart -> Excel
=================================
Uso típico:

  pip install -r requirements.txt
  python -m playwright install chromium

  # 1) Extração completa + Excel (abre navegador, faça login se pedir, depois extrai sozinho):
  python copart_leilao.py --url "https://www.copart.com.br/saleListResult/10493/2026-10-07?..." --saida leilao_goiania_07-10-2026.xlsx

  # 2) Extração + monitoramento AO VIVO do item do topo (refresh até sair o valor):
  python copart_leilao.py --url "..." --saida leilao_goiania.xlsx --monitorar --intervalo 15

  # 3) Só gerar Excel de exemplo (sem acessar o site, para ver o formato):
  python copart_leilao.py --exemplo

Como funciona o "ao vivo":
- O primeiro item do topo da lista é o que está em pregão naquele momento.
- O script lê o topo, guarda o lance, dá refresh (page.reload) a cada N segundos.
- Quando o topo MUDA (o lote anterior saiu e o próximo subiu) OU o status vira
  VENDIDO/ARREMATADO/ADJUDICADO, ele registra o último lance visto como
  "Valor Arrematado" e salva o Excel na hora (não perde nada se fechar).
- Para parar: Ctrl+C. O Excel já estará salvo.

OBS: Copart usa proteção Incapsula/Imperva. Por isso usamos Playwright com
navegador de verdade (chromium). Se aparecer captcha/login, faça manualmente
na janela que abrir e depois o script continua sozinho.
"""

import argparse
import re
import sys
import time
from datetime import datetime
from pathlib import Path

URL_PADRAO = (
    "https://www.copart.com.br/saleListResult/10493/2026-10-07"
    "?location=Goi%C3%A2nia%20-%20GO&saleDate=1791374400000"
    "&from=%2FsalesListResult&auctioneer=Claudio%20Moreira%20Santos%20-%20Juceg%20071%2F19&yardNum=56"
)

COLUNAS = [
    "Ordem",
    "Codigo / Lote",
    "Ano Modelo",
    "Marca",
    "Modelo",
    "Titulo (Marca + Modelo)",
    "Documento",
    "Categoria",
    "Condicao",
    "Valor FIPE (R$)",
    "Patio Veiculo",
    "Data Leilao / Situacao",
    "Patio do Leilao",
    "Lote/Vaga",
    "Lance Atual (R$)",
    "Valor Arrematado (R$)",
    "Status",
    "Link Lote",
    "Observacao",
    "Atualizado Em",
]

STATUS_FINAL = ("vendido", "sold", "arremat", "adjudic", "finalizado", "encerrado", "conclu")


# ---------------------------------------------------------------- Excel

def salvar_excel(lotes: list, arquivo: str):
    """Salva/atualiza o Excel. Pode ser chamada a cada refresh sem medo."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Leilao"

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    thin = Side(style="thin", color="B0B0B0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for c, nome in enumerate(COLUNAS, start=1):
        cell = ws.cell(row=1, column=c, value=nome)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUNAS))}1"
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    for i, l in enumerate(lotes, start=2):
        valores = [
            l.get("ordem", i - 1),
            l.get("lote", ""),
            l.get("ano", ""),
            l.get("marca", ""),
            l.get("modelo", ""),
            l.get("titulo", ""),
            l.get("documento", ""),
            l.get("categoria", ""),
            l.get("condicao", ""),
            l.get("valor_fipe", ""),
            l.get("patio_veiculo", ""),
            l.get("data_leilao", ""),
            l.get("patio_leilao", ""),
            l.get("lote_vaga", ""),
            l.get("lance_atual", ""),
            l.get("valor_arrematado", ""),
            l.get("status", ""),
            l.get("link", ""),
            l.get("observacao", ""),
            l.get("atualizado_em", ""),
        ]
        for c, v in enumerate(valores, start=1):
            cell = ws.cell(row=i, column=c, value=v)
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if c in (10, 15, 16):
                cell.alignment = Alignment(horizontal="right", vertical="center")

    larguras = [7, 13, 11, 14, 16, 30, 14, 18, 12, 15, 16, 16, 15, 13, 15, 17, 18, 44, 30, 19]
    for idx, lg in enumerate(larguras, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = lg
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # Aba 2: log do ao vivo
    ws2 = wb.create_sheet("Monitoramento_AoVivo")
    ws2.append(["Data/Hora", "Evento", "Lote", "Titulo", "Lance/Valor (R$)", "Detalhe"])
    for c in ws2[1]:
        c.fill = PatternFill("solid", fgColor="385723")
        c.font = Font(bold=True, color="FFFFFF")
    ws2.freeze_panes = "A2"
    ws2.sheet_properties.pageSetUpPr.fitToPage = True
    for log in getattr(salvar_excel, "_logs", []):
        ws2.append(log)
    if ws2.max_row > 1:
        ws2.auto_filter.ref = f"A1:F{ws2.max_row}"
        ws2.column_dimensions["A"].width = 20
        ws2.column_dimensions["B"].width = 28
        ws2.column_dimensions["C"].width = 14
        ws2.column_dimensions["D"].width = 40
        ws2.column_dimensions["E"].width = 18
        ws2.column_dimensions["F"].width = 40

    Path(arquivo).parent.mkdir(parents=True, exist_ok=True)
    wb.save(arquivo)
    print(f"[excel] salvo: {arquivo} ({len(lotes)} itens)")


def log_evento(evento, lote="", titulo="", valor="", detalhe=""):
    if not hasattr(salvar_excel, "_logs"):
        salvar_excel._logs = []
    salvar_excel._logs.append([
        datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        evento, str(lote), str(titulo), str(valor), str(detalhe),
    ])


# ------------------------------------------------------- Extração

JS_EXTRAIR = r"""
() => {
  const tables = [...document.querySelectorAll('table')];
  // escolhe a tabela com mais linhas (é a grade de lotes, não menus)
  let best = null, bestRows = [];
  tables.forEach(t => {
    const rows = [...t.querySelectorAll('tbody tr')];
    if (rows.length > bestRows.length) { best = t; bestRows = rows; }
  });
  const out = [];
  bestRows.forEach(tr => {
    const tds = [...tr.querySelectorAll('td')].map(td => td.innerText.trim());
    if (!tds.length) return;
    const texto = tr.innerText || '';
    if (texto.length < 5) return;
    const a = tr.querySelector('a[href*="/lot/"]');
    // pega também o código clicável (link com o número dentro do td)
    let link = a ? a.href : '';
    if (!link) {
      const m = (texto.match(/(\d{7})/) || [])[1];
      if (m) link = 'https://www.copart.com.br/lot/' + m;
    }
    out.push({cells: tds, link: link, texto: texto.slice(0, 2000)});
  });
  let ths = [];
  if (best) ths = [...best.querySelectorAll('thead th')].map(th => th.innerText.trim());
  // total de registros ("Mostrando 1 a 20 de 27 registros")
  const bodyTxt = document.body.innerText || '';
  const mt = bodyTxt.match(/Mostrando\s+\d+\s+a\s+\d+\s+de\s+(\d+)\s+registros/i);
  return {headers: ths, rows: out, total: mt ? mt[1] : '', url: location.href, title: document.title};
}
"""


def _limpa(v: str) -> str:
    return re.sub(r"\s+", " ", (v or "").strip())


def _extrai_reais(texto: str) -> list:
    """Retorna TODOS os valores R$ encontrados, em ordem."""
    return re.findall(r"R\$\s?([\d\.\,]+)", texto or "")


def parse_lote(cells: list, link: str, texto: str, ordem: int, headers: list | None = None) -> dict:
    """Mapeia a linha usando os headers reais do Copart BR (quando disponíveis).

    Layout real (print da tela):
    [chk, Imagens, Código, Ano Modelo, Marca, Modelo, Documento, Categoria,
     Condição, Valor FIPE, Pátio Veículo, Data do Leilão, Pátio do Leilão,
     Lote/Vaga, Lance Atual, ...]
    """
    headers = headers or []
    joined = " | ".join(cells)

    def col_por_nome(*nomes):
        # 1) tenta igualdade exata do header (evita 'Modelo' casar com 'Ano Modelo')
        norm_headers = [re.sub(r"\s+", " ", (h or "").strip().lower()) for h in headers]
        for n in nomes:
            nl = n.strip().lower()
            for i, hl in enumerate(norm_headers):
                if hl == nl and i < len(cells):
                    return _limpa(cells[i])
        # 2) fallback: contém
        for n in nomes:
            for i, h in enumerate(headers):
                hl = (h or "").lower()
                if n in hl and i < len(cells):
                    return _limpa(cells[i])
        return ""

    if headers and len(cells) >= 8:
        codigo = col_por_nome("código", "codigo") or ""
        ano = col_por_nome("ano modelo", "ano") or ""
        marca = col_por_nome("marca") or ""
        modelo = col_por_nome("modelo") or ""
        documento = col_por_nome("documento") or ""
        categoria = col_por_nome("categoria") or ""
        condicao = col_por_nome("condi", "condiç") or ""
        fipe = col_por_nome("fipe") or ""
        patio_vei = col_por_nome("pátio ve", "patio ve") or ""
        data_leilao = col_por_nome("data do leil") or ""
        patio_lei = col_por_nome("pátio do leil", "patio do leil") or ""
        lote_vaga = col_por_nome("lote/vaga", "lote") or ""
        lance_cell = col_por_nome("lance atual", "lance") or ""
        # Código às vezes vem como "1143295 Favoritar"
        mcode = re.search(r"\b(\d{6,9})\b", codigo + " " + (link or ""))
        lote = mcode.group(1) if mcode else _limpa(codigo.split()[0] if codigo else "")
        # Lance: a célula tem "Lance inicial : R$ 33.200,00 BRL Dar Lance Agora"
        # Pega o ÚLTIMO R$ da célula (o do lance, não FIPE). FIPE está em outra coluna.
        reais = _extrai_reais(lance_cell)
        lance = reais[-1] if reais else ""
        if not lance:  # fallback: último R$ da linha que não seja FIPE
            todos = _extrai_reais(joined)
            lance = todos[-1] if todos else ""
        fipe_val = _extrai_reais(fipe)
        fipe = fipe_val[0] if fipe_val else _limpa(fipe)
        titulo = _limpa(f"{marca} {modelo} {ano}".strip())[:150]
        if len(titulo) < 4:
            titulo = _limpa(texto[:120].replace("\n", " "))
        status = _limpa(data_leilao[:60])  # "AO VIVO AGORA" ou data
    else:
        # fallback genérico (layout desconhecido)
        m = re.search(r"\b(\d{6,9})\b", (link or "") + " " + joined)
        lote = m.group(1) if m else ""
        todos = _extrai_reais(joined)
        lance = todos[-1] if todos else ""
        fipe = todos[0] if len(todos) > 1 else ""
        m3 = re.search(r"\b(19[8-9]\d|20[0-2]\d)\b", joined)
        ano = m3.group(0) if m3 else ""
        marca, modelo, documento, categoria = "", "", "", ""
        condicao, patio_vei, data_leilao, patio_lei, lote_vaga = "", "", "", "", ""
        codigo = lote
        titulo = _limpa((cells[1] if len(cells) > 1 else texto[:120]))[:150]
        status = _limpa((cells[-1] if cells else ""))[:60]

    return {
        "ordem": ordem,
        "lote": lote,
        "ano": ano,
        "marca": _limpa(marca),
        "modelo": _limpa(modelo),
        "titulo": titulo,
        "documento": _limpa(documento),
        "categoria": _limpa(categoria),
        "condicao": _limpa(condicao),
        "valor_fipe": _limpa(fipe),
        "patio_veiculo": _limpa(patio_vei),
        "data_leilao": _limpa(data_leilao),
        "patio_leilao": _limpa(patio_lei) or "Goiânia - GO",
        "lote_vaga": _limpa(lote_vaga),
        "lance_atual": _limpa(lance),
        "valor_arrematado": "",
        "status": status or _limpa(data_leilao),
        "link": link,
        "observacao": joined[:500],
        "atualizado_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
    }


def refinar_com_texto(lote_dict: dict):
    return lote_dict


def dispensar_cookies(page):
    """Fecha banner de cookies que cobre a grade."""
    for txt in ("Reject All", "Rejeitar", "Recusar", "Accept All Cookies"):
        try:
            el = page.query_selector(f'button:has-text("{txt}")')
            if el and el.is_visible():
                el.click()
                page.wait_for_timeout(1200)
                print(f"[cookies] clicado: {txt}")
                break
        except Exception:
            continue


def esperar_grade(page, timeout_ms=90000):
    """Espera a grade de lotes renderizar (Angular + live demoram). Faz scroll p/ induzir load."""
    fim = time.time() + timeout_ms / 1000
    tentativa = 0
    while time.time() < fim:
        tentativa += 1
        try:
            n = page.evaluate('document.querySelectorAll("table tbody tr").length')
            txt = page.evaluate('(document.body.innerText || "").slice(0, 4000)')
            if n and n >= 3:
                return True
            # scroll para forçar o Angular a bindar os dados
            if tentativa % 3 == 0:
                try:
                    page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                    page.wait_for_timeout(1200)
                    page.evaluate("window.scrollTo(0, 0)")
                    dispensar_cookies(page)
                except Exception:
                    pass
            if "Mostrando" in (txt or "") and n and n >= 1:
                return True
        except Exception:
            pass
        page.wait_for_timeout(2500)
    return False


def extrair_lotes_da_pagina(page) -> list:
    data = page.evaluate(JS_EXTRAIR)
    headers = data.get("headers", [])
    rows = data.get("rows", [])
    print(f"[extracao] headers: {headers}")
    print(f"[extracao] linhas: {len(rows)} total_registros_site={data.get('total')}")
    lotes = []
    for i, r in enumerate(rows, start=1):
        d = parse_lote(r.get("cells", []), r.get("link", ""), r.get("texto", ""), i, headers)
        lotes.append(d)
    lotes = [l for l in lotes if l.get("lote")]
    vistos, unicos = set(), []
    for l in lotes:
        chave = l.get("lote")
        if chave in vistos:
            continue
        vistos.add(chave)
        unicos.append(l)
    for i, l in enumerate(unicos, start=1):
        l["ordem"] = i
    return unicos


def coletar_todas_paginas(page, max_paginas=50) -> list:
    """Varre todas as páginas numeradas (1, 2, Próximo...) mesclando sem duplicar."""
    todos = []
    for p in range(1, max_paginas + 1):
        try:
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            page.wait_for_timeout(1500)
            page.evaluate("window.scrollTo(0, 0)")
            page.wait_for_timeout(800)
        except Exception:
            pass
        lotes = extrair_lotes_da_pagina(page)
        chaves = {l.get("lote") for l in todos}
        novos = [l for l in lotes if l.get("lote") not in chaves]
        todos.extend(novos)
        print(f"[paginacao] pág {p}: {len(lotes)} na tela, {len(novos)} novos, total {len(todos)}")
        if not lotes or not novos:
            break
        # tenta ir para a próxima página numerada (p+1) ou botão Próximo
        avancou = False
        for seletor in [
            f'a:has-text("{p + 1}")',
            'a:has-text("Próximo")', 'a:has-text("Next")',
            'button:has-text("Próximo")', 'button:has-text("Next")',
        ]:
            try:
                els = page.query_selector_all(seletor)
                for el in els:
                    try:
                        if el.is_visible() and el.is_enabled():
                            txt = (el.inner_text() or "").strip()
                            if txt in (str(p + 1), "Próximo", "Next", ">", "»"):
                                el.click()
                                page.wait_for_timeout(6000)
                                esperar_grade(page, timeout_ms=20000)
                                avancou = True
                                break
                    except Exception:
                        continue
                if avancou:
                    break
            except Exception:
                continue
        if not avancou:
            break
    for i, l in enumerate(todos, start=1):
        l["ordem"] = i
    return todos


# ------------------------------------------------------- Ao vivo

def lance_para_float(txt) -> float | None:
    if txt is None or txt == "":
        return None
    s = str(txt).strip().replace("R$", "").strip()
    # "12.345,67" -> 12345.67
    try:
        if "," in s:
            s = s.replace(".", "").replace(",", ".")
        return float(re.sub(r"[^\d\.]", "", s) or "nan")
    except Exception:
        return None


def eh_status_final(status: str) -> bool:
    s = (status or "").lower()
    return any(k in s for k in STATUS_FINAL)


def monitorar_ao_vivo(page, lotes: list, arquivo: str, intervalo: int = 15):
    """Loop de monitoramento do item do topo com refresh.
    Regra pedida: o 1º da lista = item em pregão. Fica dando refresh até
    o valor sair; quando sair, registra o último lance como arrematado.
    """
    print("\n" + "=" * 70)
    print("MODO AO VIVO: monitorando o TOPO da lista.")
    print(f"Refresh a cada {intervalo}s. Ctrl+C para parar (o Excel já vai salvando).")
    print("=" * 70 + "\n")

    # índice do lote que está no topo / em pregão
    # começa pelo primeiro da lista extraída
    em_pregao_lote = lotes[0]["lote"] if lotes else ""
    ultimo_lance_visto = lotes[0].get("lance_atual", "") if lotes else ""
    por_lote = {l["lote"]: l for l in lotes if l.get("lote")}
    # se lotes sem número (layout diferente), usa link como chave
    if not por_lote:
        por_lote = {l.get("link", str(i)): l for i, l in enumerate(lotes)}

    log_evento("INICIO_MONITORAMENTO", em_pregao_lote,
               lotes[0].get("titulo", "") if lotes else "", ultimo_lance_visto,
               f"Topo inicial. Total de itens: {len(lotes)}")
    salvar_excel(lotes, arquivo)

    rodada = 0
    try:
        while True:
            rodada += 1
            try:
                page.reload(wait_until="domcontentloaded")
            except Exception:
                pass
            dispensar_cookies(page)
            esperar_grade(page, timeout_ms=25000)

            atuais = extrair_lotes_da_pagina(page)
            if not atuais:
                print(f"[{rodada}] página sem itens (bloqueio/captcha?). Aguardando {intervalo}s...")
                try:
                    page.screenshot(path=f"debug_rodada_{rodada}.png")
                except Exception:
                    pass
                time.sleep(intervalo)
                continue

            topo = atuais[0]
            topo_lote = topo.get("lote") or topo.get("link")
            topo_lance = topo.get("lance_atual", "")
            topo_status = topo.get("status", "")

            # atualiza lance do item em pregão se o topo ainda é o mesmo
            chave_pregao = em_pregao_lote or list(por_lote.keys())[0] if por_lote else topo_lote
            # normaliza: se estamos usando link como chave
            if chave_pregao not in por_lote and topo_lote in por_lote:
                chave_pregao = topo_lote

            if topo_lote == (em_pregao_lote or topo_lote):
                # mesmo item continua no topo -> atualiza lance
                if topo_lance and topo_lance != ultimo_lance_visto:
                    print(f"[{rodada}] Lote {topo.get('lote')} | lance: {ultimo_lance_visto} -> {topo_lance}")
                    ultimo_lance_visto = topo_lance
                    alvo = por_lote.get(chave_pregao)
                    if alvo is not None:
                        alvo["lance_atual"] = topo_lance
                        alvo["status"] = "EM PREGÃO (ao vivo)"
                        alvo["atualizado_em"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                    log_evento("LANCE_ATUALIZADO", topo.get("lote"), topo.get("titulo"), topo_lance)
                else:
                    print(f"[{rodada}] Lote {topo.get('lote')} no topo, lance={topo_lance or '(ainda sem lance)'} | status={topo_status}")

                # verifica se finalizou SEM trocar o topo (status mudou)
                if eh_status_final(topo_status) or eh_status_final(topo_lance):
                    alvo = por_lote.get(chave_pregao)
                    if alvo is not None and not alvo.get("valor_arrematado"):
                        final = topo_lance or ultimo_lance_visto
                        alvo["valor_arrematado"] = final
                        alvo["status"] = "ARREMATADO"
                        alvo["atualizado_em"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                        print(f"*** ARREMATADO lote {alvo.get('lote')} por R$ {final} ***")
                        log_evento("ARREMATADO", alvo.get("lote"), alvo.get("titulo"), final, topo_status)
                        salvar_excel(lotes, arquivo)
                        # avança para o próximo da fila
                        avanca_para_proximo(lotes, por_lote, arquivo)
                        if por_lote:
                            em_pregao_lote = proximo_em_aberto(lotes)
                            ultimo_lance_visto = ""
                        else:
                            break
                salvar_excel(lotes, arquivo)
            else:
                # TOPO MUDOU => o anterior saiu do pregão => registra valor final
                anterior = por_lote.get(em_pregao_lote) or por_lote.get(chave_pregao)
                if anterior is not None and not anterior.get("valor_arrematado"):
                    anterior["valor_arrematado"] = ultimo_lance_visto
                    anterior["status"] = "ARREMATADO (topo avançou)"
                    anterior["atualizado_em"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                    print(f"*** Lote {anterior.get('lote')} saiu do topo. Valor final R$ {ultimo_lance_visto} ***")
                    log_evento("ARREMATADO_TOPO_AVANCOU", anterior.get("lote"),
                               anterior.get("titulo"), ultimo_lance_visto,
                               f"Novo topo: {topo.get('lote')} - {topo.get('titulo')}")
                # incorpora itens novos que apareceram
                for a in atuais:
                    if a.get("lote") and a["lote"] not in por_lote:
                        a["ordem"] = len(lotes) + 1
                        lotes.append(a)
                        por_lote[a["lote"]] = a
                salvar_excel(lotes, arquivo)
                em_pregao_lote = topo.get("lote") or topo.get("link")
                ultimo_lance_visto = topo_lance
                print(f"[{rodada}] Novo item em pregão: {topo.get('lote')} - {topo.get('titulo')} (lance {topo_lance})")
                log_evento("NOVO_TOPO", topo.get("lote"), topo.get("titulo"), topo_lance)

            # checa término: todos com valor arrematado
            if lotes and all(l.get("valor_arrematado") for l in lotes if l.get("lote")):
                print("Todos os lotes com valor final. Encerrando.")
                break

            time.sleep(intervalo)
    except KeyboardInterrupt:
        print("\nInterrompido pelo usuário. Salvando Excel...")
        salvar_excel(lotes, arquivo)
        print(f"Arquivo final: {arquivo}")


def avanca_para_proximo(lotes, por_lote, arquivo):
    pass  # (placeholder p/ extensão futura)


def proximo_em_aberto(lotes):
    for l in lotes:
        if not l.get("valor_arrematado"):
            return l.get("lote") or l.get("link")
    return ""


# ------------------------------------------------------- Main

def abrir_navegador(headless: bool):
    from playwright.sync_api import sync_playwright
    pw = sync_playwright().start()
    browser = pw.chromium.launch(headless=headless)
    context = browser.new_context(
        locale="pt-BR",
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36",
        viewport={"width": 1366, "height": 900},
    )
    page = context.new_page()
    return pw, browser, page


def gerar_exemplo(arquivo: str):
    exemplos = [
        {"ordem": 1, "lote": "1143295", "ano": "2025", "marca": "FIAT", "modelo": "STRADA",
         "titulo": "FIAT STRADA 2025", "documento": "RECUPERÁVEL", "categoria": "Picapes Pequenas",
         "condicao": "COLISÃO", "valor_fipe": "83.295,00", "patio_veiculo": "CUIABÁ - MT",
         "data_leilao": "AO VIVO AGORA", "patio_leilao": "Goiânia - GO", "lote_vaga": "82/Recebimento",
         "lance_atual": "33.200,00",
         "valor_arrematado": "35.500,00", "status": "ARREMATADO", "link": "https://www.copart.com.br/lot/1143295",
         "observacao": "Exemplo realista", "atualizado_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S")},
        {"ordem": 2, "lote": "1132291", "ano": "2024", "marca": "SCANIA", "modelo": "P 320",
         "titulo": "SCANIA P 320 2024", "documento": "IRRECUPERÁVEL", "categoria": "Caminhões e Rebocadores",
         "condicao": "COLISÃO", "valor_fipe": "683.016,00", "patio_veiculo": "GOIÂNIA - GO",
         "data_leilao": "AO VIVO AGORA", "patio_leilao": "Goiânia - GO", "lote_vaga": "83/C05",
         "lance_atual": "34.200,00",
         "valor_arrematado": "", "status": "EM PREGÃO (ao vivo)", "link": "https://www.copart.com.br/lot/1132291",
         "observacao": "Item atual do topo", "atualizado_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S")},
    ]
    log_evento("EXEMPLO_GERADO", "1143295", "FIAT STRADA 2025", "35.500,00", "Arquivo de demonstração do formato")
    log_evento("LANCE_ATUALIZADO", "1132291", "SCANIA P 320 2024", "34.200,00", "Item do topo em pregão")
    salvar_excel(exemplos, arquivo)
    print(f"Exemplo gerado: {arquivo}. Abra no Excel para ver o formato.")


def main():
    ap = argparse.ArgumentParser(description="Copart -> Excel + monitoramento ao vivo do topo")
    ap.add_argument("--url", default=URL_PADRAO, help="Link do saleListResult do leilão")
    ap.add_argument("--saida", default="leilao_goiania_07-10-2026.xlsx", help="Arquivo Excel de saída")
    ap.add_argument("--monitorar", action="store_true", help="Ativa loop ao vivo com refresh no topo")
    ap.add_argument("--intervalo", type=int, default=15, help="Segundos entre refreshs (padrão 15)")
    ap.add_argument("--headless", action="store_true", help="Roda navegador oculto (padrão abre janela p/ login/captcha manual)")
    ap.add_argument("--max-paginas", type=int, default=20, help="Máx. páginas a varrer na extração")
    ap.add_argument("--exemplo", action="store_true", help="Só gera Excel de exemplo, sem acessar o site")
    args = ap.parse_args()

    if args.exemplo:
        gerar_exemplo(args.saida)
        return

    try:
        from playwright.sync_api import sync_playwright  # noqa
    except ImportError:
        print("Falta o Playwright. Rode:\n  pip install -r requirements.txt\n  python -m playwright install chromium")
        sys.exit(1)

    pw, browser, page = abrir_navegador(headless=args.headless)
    try:
        print(f"Abrindo:\n{args.url}\n")
        if not args.headless:
            print("Se pedir login/captcha, resolva NA JANELA que abriu. O script espera 25s...")
        page.goto(args.url, wait_until="domcontentloaded", timeout=60000)
        page.wait_for_timeout(8000)
        dispensar_cookies(page)
        if not esperar_grade(page, timeout_ms=90000):
            print("[aviso] grade demorou — recarregando 1x e tentando de novo...");
            try:
                page.reload(wait_until="domcontentloaded")
                page.wait_for_timeout(8000)
                dispensar_cookies(page)
            except Exception:
                pass
            if not esperar_grade(page, timeout_ms=90000):
                print("[aviso] mesmo após reload a grade não veio (bloqueio/Copart lento).");
        else:
            print("[ok] grade carregada.")
        page.wait_for_timeout(3000)
        try:
            page.screenshot(path="debug_pagina_inicial.png")
        except Exception:
            pass

        lotes = coletar_todas_paginas(page, max_paginas=args.max_paginas)
        if not lotes:
            # salva debug para ajuste de seletores
            try:
                html = page.content()
                Path("debug_pagina.html").write_text(html, encoding="utf-8")
                page.screenshot(path="debug_sem_itens.png")
            except Exception:
                pass
            print("NENHUM item extraído. Salvei debug_pagina.html + debug_sem_itens.png.")
            print("Me mande esses arquivos (ou print da tela) que ajusto os seletores.")
            return

        print(f"Extraídos {len(lotes)} itens. Salvando {args.saida}...")
        log_evento("EXTRACAO_INICIAL", "", "", "", f"{len(lotes)} itens extraídos de {args.url}")
        salvar_excel(lotes, args.saida)

        if args.monitorar:
            monitorar_ao_vivo(page, lotes, args.saida, intervalo=args.intervalo)
        else:
            print("\nPronto! Para monitorar o AO VIVO no dia do leilão, rode com --monitorar:")
            print(f'  python copart_leilao.py --url "{args.url}" --saida "{args.saida}" --monitorar --intervalo 15')
    finally:
        try:
            browser.close()
            pw.stop()
        except Exception:
            pass


if __name__ == "__main__":
    main()
