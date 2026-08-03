"""
KPI de Manutenção — Cliente Contrato PMOC (Refrigeração Aços)
-----------------------------------------------------
Varre recursivamente os relatórios em PDF (gerados pelo app Produttivo)
dentro de "REFRIGERAÇÃO AÇOS - RELATÓRIOS/ANO 2026" no Nextcloud, classifica
cada um como preventiva ou corretiva, extrai data e duração, e gera um
relatório .docx com KPIs por categoria e por equipamento.

Premissas validadas com dados reais (ver histórico do projeto):
- Estrutura: ANO 2026/<Categoria>/<Equipamento>/<arquivo>.pdf
- Tipo: definido pelas primeiras linhas do texto do PDF, não pelo nome do
  arquivo (o nome do arquivo pode estar errado — já vimos casos reais).
  Regra (confirmada com o Rafael em 30/07/2026):
    * título começa com "PMOC"                -> preventiva
    * título começa com "Manutenção Corretiva" -> corretiva (formulário
      novo, em uso desde jul/2026, só pra corretivas)
    * título começa com "Check list" -> vencimento_filtro (checklist de
      troca de filtro de bebedouro, não é atendimento de equipamento — ver
      extrair_datas_filtro)
    * título começa com "Ordem de Serviço":
        - categoria CHILLER'S -> preventiva (chillers ainda não têm
          template PMOC próprio; até criarem um, o time usa "Ordem de
          Serviço" pras visitas de rotina)
        - qualquer outra categoria -> corretiva (chamado avulso)
  Cuidado: essa regra pode classificar como preventiva alguma corretiva
  de chiller registrada ANTES do formulário "Manutenção Corretiva"
  existir (pré-julho/2026) — não dá pra distinguir com o dado disponível.
- Data: campo "Em: dd/mm/aaaa hh:mm" no topo do PDF (fallback: data no
  nome do arquivo, pro caso de PDF escaneado sem texto extraível).
- Duração: só é calculada para corretivas, via campos "Início do
  Trabalho" / "Fim do Trabalho". PMOC não tem duração real de serviço
  no PDF (só tempo de preenchimento do checklist no app, que não
  representa trabalho em campo).

Depois de gerar o .docx, publica os atendimentos na tabela
kpi_manutencao_atendimentos do Supabase (mesmo projeto do gerador de
orçamento) via service role key — é assim que a aba nova do Next.js
(/kpi-manutencao) exibe os dados, sem o Next.js falar com o Nextcloud
diretamente e sem tocar em nada de Omie/cadastro de clientes.

USO
1. Preencher .env com NEXTCLOUD_URL, NEXTCLOUD_USER, NEXTCLOUD_APP_PASSWORD,
   SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY.
2. python kpi_manutencao.py
"""

import json
import os
import re
import shutil
import sys
from dataclasses import asdict, dataclass, field
from datetime import datetime, timedelta

import pdfplumber
import requests
from docx import Document
from docx.shared import Pt
from dotenv import load_dotenv
from webdav3.client import Client

# Terminal/arquivo de saída no Windows costuma usar cp1252, que não sabe
# codificar acentos nem símbolos (⚠, ✓ etc.) — força UTF-8 pra print() não
# quebrar quando a saída for redirecionada pra arquivo.
sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

load_dotenv()

NEXTCLOUD_URL = os.environ["NEXTCLOUD_URL"]
NEXTCLOUD_USER = os.environ["NEXTCLOUD_USER"]
NEXTCLOUD_APP_PASSWORD = os.environ["NEXTCLOUD_APP_PASSWORD"]
SUPABASE_URL = os.environ["SUPABASE_URL"]
SUPABASE_SERVICE_ROLE_KEY = os.environ["SUPABASE_SERVICE_ROLE_KEY"]

PASTA_RAIZ = "REFRIGERAÇÃO AÇOS - RELATÓRIOS/ANO 2026"
PASTA_TEMP = "temp_pdfs"
SAIDA_DOCX = f"kpi_manutencao_acos_{datetime.now():%Y%m%d}.docx"
CACHE_PATH = "cache_pdfs.json"
# Bump sempre que classificar_tipo/extrair_data/extrair_duracao_corretiva
# mudar de comportamento — invalida entradas de cache presas na regra antiga.
CACHE_VERSAO = 6
# Intervalo (em dias) considerado "em dia" pra cada periodicidade de PMOC,
# com uma folga pequena pra não marcar como atrasado por 1-2 dias de
# variação de agenda. Só usado pra categorias com PMOC próprio (hoje: AR
# CONDICIONADOS e BEBEDOUROS) — chillers ainda não têm template de PMOC,
# então continuam na regra simplificada de "mês atual".
INTERVALO_DIAS_POR_PERIODICIDADE = {"semanal": 10, "mensal": 35, "trimestral": 100}
# Periodicidade da troca de filtro de bebedouro (fixa, sem template de PMOC
# próprio) — usada só pra corrigir o bug abaixo, o cálculo real de
# conforme/vencendo/atrasado é feito no painel (Next.js).
PERIODICIDADE_FILTRO_DIAS = 90


@dataclass
class Atendimento:
    categoria: str
    equipamento: str
    arquivo: str
    tipo: str  # 'preventiva' | 'corretiva' | 'desconhecida'
    data: datetime | None
    duracao_minutos: float | None
    ultima_substituicao_filtro: datetime | None = None
    proxima_substituicao_filtro: datetime | None = None
    periodicidade: str | None = None  # 'semanal' | 'mensal' | 'trimestral' | None


def conectar_nextcloud() -> Client:
    options = {
        "webdav_hostname": f"{NEXTCLOUD_URL}/remote.php/dav/files/{NEXTCLOUD_USER}",
        "webdav_login": NEXTCLOUD_USER,
        "webdav_password": NEXTCLOUD_APP_PASSWORD,
    }
    return Client(options)


def listar_arvore(client: Client, pasta_raiz: str) -> list[dict]:
    """Uma chamada só (Depth: infinity) pra árvore inteira, em vez de uma
    chamada por pasta — o jeito anterior (uma requisição por pasta de
    equipamento, ~100+ pastas) é o motivo da varredura levar ~10min só pra
    listar; isso aqui deve levar segundos. Retorna cada item com o
    `path` já relativo à raiz do WebDAV (sem o prefixo /remote.php/...)."""
    prefixo_servidor = f"/remote.php/dav/files/{NEXTCLOUD_USER}/"

    def caminho_relativo(item: dict) -> str:
        caminho_servidor = item["path"]
        if caminho_servidor.startswith(prefixo_servidor):
            caminho = caminho_servidor[len(prefixo_servidor):]
        else:
            caminho = caminho_servidor.lstrip("/")
        return caminho.rstrip("/")

    itens_brutos = client.list(pasta_raiz, get_info=True, recursive=True)
    raiz_norm = pasta_raiz.strip("/")
    resultado = []
    for item in itens_brutos:
        caminho = caminho_relativo(item)
        if caminho == raiz_norm:
            continue
        resultado.append({
            "path": caminho,
            "etag": item.get("etag"),
            "isdir": bool(item.get("isdir")),
        })
    return resultado


def extrair_pdfs(itens: list[dict], pasta_raiz: str) -> list[dict]:
    """Filtra a árvore completa (listar_arvore) só pros PDFs."""
    return [
        {"path": item["path"], "etag": item["etag"]}
        for item in itens
        if not item["isdir"] and item["path"].lower().endswith(".pdf")
    ]


def extrair_estrutura(itens: list[dict], pasta_raiz: str) -> dict[str, list[str]]:
    """Filtra a árvore completa (listar_arvore) pra estrutura Categoria/
    Equipamento — dá o total de equipamentos "no contrato" por categoria,
    incluindo os que não tiveram nenhum atendimento ainda este ano."""
    raiz_norm = pasta_raiz.strip("/")
    resultado: dict[str, list[str]] = {}
    for item in itens:
        if not item["isdir"]:
            continue
        caminho = item["path"]
        if not caminho.startswith(raiz_norm + "/"):
            continue
        rel = caminho[len(raiz_norm) + 1:]
        partes = rel.split("/")
        # só pastas de 2 níveis (Categoria/Equipamento), relativo à raiz —
        # qualquer outra profundidade não interessa aqui.
        if len(partes) != 2:
            continue
        categoria, equipamento = partes
        resultado.setdefault(categoria, []).append(equipamento)
    return resultado


def atendimento_para_dict(a: Atendimento) -> dict:
    d = asdict(a)
    d["data"] = a.data.isoformat() if a.data else None
    d["ultima_substituicao_filtro"] = (
        a.ultima_substituicao_filtro.isoformat() if a.ultima_substituicao_filtro else None
    )
    d["proxima_substituicao_filtro"] = (
        a.proxima_substituicao_filtro.isoformat() if a.proxima_substituicao_filtro else None
    )
    return d


def dict_para_atendimento(d: dict) -> Atendimento:
    return Atendimento(
        categoria=d["categoria"],
        equipamento=d["equipamento"],
        arquivo=d["arquivo"],
        tipo=d["tipo"],
        data=datetime.fromisoformat(d["data"]) if d["data"] else None,
        duracao_minutos=d["duracao_minutos"],
        ultima_substituicao_filtro=(
            datetime.fromisoformat(d["ultima_substituicao_filtro"])
            if d.get("ultima_substituicao_filtro")
            else None
        ),
        proxima_substituicao_filtro=(
            datetime.fromisoformat(d["proxima_substituicao_filtro"])
            if d.get("proxima_substituicao_filtro")
            else None
        ),
        periodicidade=d.get("periodicidade"),
    )


def carregar_cache(caminho: str) -> dict:
    if not os.path.exists(caminho):
        return {}
    with open(caminho, "r", encoding="utf-8") as f:
        return json.load(f)


def salvar_cache(cache: dict, caminho: str) -> None:
    with open(caminho, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)


def categoria_e_equipamento(caminho_pdf: str, pasta_raiz: str) -> tuple[str, str]:
    rel = caminho_pdf.strip("/")
    raiz_norm = pasta_raiz.strip("/")
    if rel.startswith(raiz_norm):
        rel = rel[len(raiz_norm):].strip("/")
    partes = rel.split("/")
    categoria = partes[0] if len(partes) >= 1 else "Desconhecida"
    equipamento = partes[-2] if len(partes) >= 2 else "Raiz"
    return categoria, equipamento


def classificar_tipo(texto: str, categoria: str) -> str:
    # Muitos relatórios do Produttivo têm uma linha só com um ID numérico
    # interno antes do título (ex: "93", "128") — por isso olhamos as
    # primeiras linhas, não só a primeira.
    linhas = [l.strip() for l in texto.splitlines() if l.strip()]
    eh_ordem_servico = False
    for linha in linhas[:3]:
        low = linha.lower()
        if low.startswith("pmoc"):
            return "preventiva"
        if low.startswith("manutenção corretiva") or low.startswith("manutencao corretiva"):
            return "corretiva"
        if low.startswith("check list"):
            # "Check list de filtros bebedouros" — não é preventiva nem
            # corretiva de equipamento, é o registro de troca/vencimento do
            # filtro do bebedouro. Ver extrair_datas_filtro pras datas.
            return "vencimento_filtro"
        if low.startswith("ordem de serviço") or low.startswith("ordem de servico"):
            eh_ordem_servico = True

    if eh_ordem_servico:
        # Chillers ainda não têm template PMOC próprio: até criarem um,
        # "Ordem de Serviço" é o formulário usado pras visitas de rotina.
        if categoria.upper().startswith("CHILLER"):
            return "preventiva"
        return "corretiva"

    return "desconhecida"


def extrair_periodicidade(texto: str) -> str | None:
    """Só confiável quando o título é 'PMOC ... Mensal/Trimestral/Semanal'
    (AR CONDICIONADOS e BEBEDOUROS têm esse template). Chillers não têm
    template de PMOC próprio ainda, então isso sempre volta None pra eles."""
    linhas = [l.strip() for l in texto.splitlines() if l.strip()]
    for linha in linhas[:3]:
        low = linha.lower()
        if not low.startswith("pmoc"):
            continue
        if "trimestral" in low:
            return "trimestral"
        if "semanal" in low:
            return "semanal"
        if "mensal" in low:
            return "mensal"
    return None


def extrair_data(texto: str, nome_arquivo: str) -> datetime | None:
    m = re.search(r"Em:\s*(\d{2})/(\d{2})/(\d{4})", texto)
    if m:
        dia, mes, ano = (int(x) for x in m.groups())
        try:
            return datetime(ano, mes, dia)
        except ValueError:
            pass
    # fallback: data no nome do arquivo (PDF escaneado sem texto, por exemplo)
    m = re.search(r"(\d{2})[.\-](\d{2})[.\-](\d{2,4})", nome_arquivo)
    if m:
        dia, mes, ano = m.groups()
        ano_int = int(ano)
        if ano_int < 100:
            ano_int += 2000
        try:
            return datetime(ano_int, int(mes), int(dia))
        except ValueError:
            return None
    return None


def extrair_duracao_corretiva(texto: str, data_doc: datetime | None) -> float | None:
    m_inicio = re.search(
        r"Início do Trabalho.*?Horário de início\s*\n?\s*(\d{2}:\d{2}:\d{2})",
        texto,
        re.DOTALL,
    )
    m_fim = re.search(
        r"Fim do Trabalho.*?Horário\s*\n?\s*(\d{2}/\d{2}/\d{4})\s+(\d{2}:\d{2}:\d{2})",
        texto,
        re.DOTALL,
    )
    if not m_inicio or not m_fim:
        return None
    fim_dt = datetime.strptime(f"{m_fim.group(1)} {m_fim.group(2)}", "%d/%m/%Y %H:%M:%S")
    inicio_time = datetime.strptime(m_inicio.group(1), "%H:%M:%S").time()
    inicio_dt = datetime.combine(fim_dt.date(), inicio_time)
    if inicio_dt > fim_dt:
        inicio_dt -= timedelta(days=1)
    return (fim_dt - inicio_dt).total_seconds() / 60


def extrair_datas_filtro(texto: str) -> tuple[datetime | None, datetime | None]:
    """Só o relatório 'Check list de filtros bebedouros' tem esses campos.
    Retorna (última substituição, próxima substituição)."""
    m = re.search(
        r"Data da (?:ultima|última) substi\w*[cç][aã]o\s+Data da pr[oó]xima substi\w*[cç][aã]o"
        r"\s*\n?\s*(\d{2})/(\d{2})/(\d{4})[^\n]*?(\d{2})/(\d{2})/(\d{4})",
        texto,
    )
    if not m:
        return None, None
    d1, m1, a1, d2, m2, a2 = m.groups()
    try:
        ultima = datetime(int(a1), int(m1), int(d1))
    except ValueError:
        ultima = None
    try:
        proxima = datetime(int(a2), int(m2), int(d2))
    except ValueError:
        proxima = None
    return ultima, proxima


def processar_pdf(client: Client, caminho_remoto: str, pasta_raiz: str) -> Atendimento:
    categoria, equipamento = categoria_e_equipamento(caminho_remoto, pasta_raiz)
    nome_arquivo = caminho_remoto.rstrip("/").split("/")[-1]
    caminho_local = os.path.join(PASTA_TEMP, nome_arquivo)

    client.download_sync(remote_path=caminho_remoto, local_path=caminho_local)

    texto = ""
    try:
        with pdfplumber.open(caminho_local) as pdf:
            texto = "\n".join((p.extract_text() or "") for p in pdf.pages)
    finally:
        if os.path.exists(caminho_local):
            os.remove(caminho_local)

    tipo = classificar_tipo(texto, categoria)
    data = extrair_data(texto, nome_arquivo)
    duracao = extrair_duracao_corretiva(texto, data) if tipo == "corretiva" else None
    ultima_filtro, proxima_filtro = extrair_datas_filtro(texto)
    periodicidade = extrair_periodicidade(texto) if tipo == "preventiva" else None

    if tipo == "vencimento_filtro" and data and proxima_filtro and proxima_filtro.date() == data.date():
        # Bug do formulário Produttivo: quando o campo "próxima substituição"
        # não é preenchido pelo técnico, ele volta com a mesma data/hora do
        # preenchimento do checklist (não é uma substituição vencida hoje —
        # é a visita de hoje que gera este checklist). Nesse caso, a data
        # real de referência é a própria visita: última substituição = hoje,
        # próxima = hoje + periodicidade.
        ultima_filtro = datetime(data.year, data.month, data.day)
        proxima_filtro = ultima_filtro + timedelta(days=PERIODICIDADE_FILTRO_DIAS)

    return Atendimento(
        categoria=categoria,
        equipamento=equipamento,
        arquivo=nome_arquivo,
        tipo=tipo,
        data=data,
        duracao_minutos=duracao,
        ultima_substituicao_filtro=ultima_filtro,
        proxima_substituicao_filtro=proxima_filtro,
        periodicidade=periodicidade,
    )


def gerar_relatorio(atendimentos: list[Atendimento], caminho_saida: str) -> None:
    doc = Document()
    doc.add_heading("KPI de Manutenção — Cliente Contrato PMOC (Refrigeração Aços)", level=0)
    doc.add_paragraph(f"Gerado em {datetime.now():%d/%m/%Y %H:%M} — Ano de referência: 2026")

    validos = [a for a in atendimentos if a.tipo != "desconhecida"]
    nao_classificados = [a for a in atendimentos if a.tipo == "desconhecida"]
    preventivas = [a for a in validos if a.tipo == "preventiva"]
    corretivas = [a for a in validos if a.tipo == "corretiva"]
    horas_corretivas_total = sum(
        (a.duracao_minutos or 0) for a in corretivas
    ) / 60

    doc.add_heading("Resumo Geral", level=1)
    resumo = doc.add_paragraph()
    resumo.add_run(f"Total de atendimentos encontrados: {len(atendimentos)}\n")
    resumo.add_run(f"Preventivas: {len(preventivas)}\n")
    resumo.add_run(f"Corretivas: {len(corretivas)}\n")
    resumo.add_run(f"Total de horas corretivas: {horas_corretivas_total:.1f}h\n")
    if nao_classificados:
        resumo.add_run(
            f"⚠ {len(nao_classificados)} arquivo(s) não classificado(s) — ver apêndice ao final.\n"
        )

    categorias = sorted({a.categoria for a in validos})
    doc.add_heading("Detalhamento por Categoria e Equipamento", level=1)
    for categoria in categorias:
        doc.add_heading(categoria, level=2)
        do_cat = [a for a in validos if a.categoria == categoria]
        equipamentos = sorted({a.equipamento for a in do_cat})

        linhas = []
        for equipamento in equipamentos:
            do_equip = [a for a in do_cat if a.equipamento == equipamento]
            n_prev = sum(1 for a in do_equip if a.tipo == "preventiva")
            n_corr = sum(1 for a in do_equip if a.tipo == "corretiva")
            horas = sum((a.duracao_minutos or 0) for a in do_equip if a.tipo == "corretiva") / 60
            linhas.append((equipamento, n_prev, n_corr, horas))

        # equipamentos mais problemáticos (mais horas corretivas) primeiro
        linhas.sort(key=lambda r: r[3], reverse=True)

        tabela = doc.add_table(rows=1, cols=4)
        tabela.style = "Light Grid Accent 1"
        hdr = tabela.rows[0].cells
        hdr[0].text, hdr[1].text, hdr[2].text, hdr[3].text = (
            "Equipamento",
            "Preventivas",
            "Corretivas",
            "Horas Corretivas",
        )
        for equipamento, n_prev, n_corr, horas in linhas:
            row = tabela.add_row().cells
            row[0].text = equipamento
            row[1].text = str(n_prev)
            row[2].text = str(n_corr)
            row[3].text = f"{horas:.1f}h"

    if nao_classificados:
        doc.add_heading("Apêndice — Arquivos Não Classificados", level=1)
        doc.add_paragraph(
            "Estes arquivos não começam com 'PMOC' nem 'Ordem de Serviço' no "
            "texto extraído (pode ser PDF escaneado sem texto, ou um formato "
            "de relatório diferente). Revisar manualmente:"
        )
        for a in nao_classificados:
            doc.add_paragraph(f"{a.categoria} / {a.equipamento} / {a.arquivo}", style="List Bullet")

    doc.save(caminho_saida)


def enviar_para_supabase(pares: list[tuple[str, Atendimento]]) -> None:
    """Substitui todo o conteúdo de kpi_manutencao_atendimentos pelo
    resultado desta execução. Dataset é pequeno (algumas centenas de
    linhas) — mais simples que fazer upsert incremental linha a linha."""
    headers = {
        "apikey": SUPABASE_SERVICE_ROLE_KEY,
        "Authorization": f"Bearer {SUPABASE_SERVICE_ROLE_KEY}",
        "Content-Type": "application/json",
    }
    tabela_url = f"{SUPABASE_URL}/rest/v1/kpi_manutencao_atendimentos"

    resp = requests.delete(f"{tabela_url}?id=not.is.null", headers=headers)
    resp.raise_for_status()

    registros = [
        {
            "caminho": caminho,
            "categoria": a.categoria,
            "equipamento": a.equipamento,
            "arquivo": a.arquivo,
            "tipo": a.tipo,
            "data": a.data.date().isoformat() if a.data else None,
            "duracao_minutos": a.duracao_minutos,
        }
        for caminho, a in pares
    ]

    # insere em lotes pra não estourar o tamanho do payload
    tamanho_lote = 200
    for i in range(0, len(registros), tamanho_lote):
        lote = registros[i : i + tamanho_lote]
        resp = requests.post(tabela_url, headers=headers, json=lote)
        resp.raise_for_status()


def montar_resumo_equipamentos(
    estrutura: dict[str, list[str]],
    pares: list[tuple[str, Atendimento]],
    data_referencia: datetime | None = None,
) -> list[dict]:
    """Um registro por equipamento (pasta), mesmo os que não tiveram
    nenhum atendimento ainda — é o que dá o "total no contrato" e permite
    calcular pendente = total - realizado.

    `data_referencia` define o "mês atual" usado pra calcular
    realizado/pendente — por padrão é hoje, mas pode ser simulado (ex:
    --mes-referencia 2026-08-01) só pra pré-visualizar o painel."""
    agora = data_referencia or datetime.now()
    por_equip: dict[tuple[str, str], dict] = {}

    def registro_vazio(categoria: str, equipamento: str) -> dict:
        return {
            "categoria": categoria,
            "equipamento": equipamento,
            "ultima_preventiva": None,
            "periodicidade": None,
            "ultima_substituicao_filtro": None,
            "proxima_substituicao_filtro": None,
            # data do checklist de filtro mais recente já visto — usada só
            # internamente pra saber qual "próxima substituição" é a mais
            # atual (a de um checklist antigo pode já estar desatualizada).
            "_data_checklist_filtro": None,
            # fallback pras categorias/equipamentos sem periodicidade
            # conhecida (hoje: chillers) — "teve preventiva neste mês?"
            "_realizado_mes_atual": False,
        }

    for categoria, equipamentos in estrutura.items():
        for equipamento in equipamentos:
            por_equip[(categoria, equipamento)] = registro_vazio(categoria, equipamento)

    for _, a in pares:
        chave = (a.categoria, a.equipamento)
        if chave not in por_equip:
            # atendimento cuja pasta não apareceu na listagem de estrutura
            # (não deveria acontecer, mas não descartamos o dado por isso)
            por_equip[chave] = registro_vazio(a.categoria, a.equipamento)
        info = por_equip[chave]
        if a.tipo == "preventiva" and a.data:
            if info["ultima_preventiva"] is None or a.data > info["ultima_preventiva"]:
                info["ultima_preventiva"] = a.data
                info["periodicidade"] = a.periodicidade
            if a.data.year == agora.year and a.data.month == agora.month:
                info["_realizado_mes_atual"] = True
        if a.proxima_substituicao_filtro and a.data:
            if info["_data_checklist_filtro"] is None or a.data > info["_data_checklist_filtro"]:
                info["_data_checklist_filtro"] = a.data
                info["ultima_substituicao_filtro"] = a.ultima_substituicao_filtro
                info["proxima_substituicao_filtro"] = a.proxima_substituicao_filtro

    # segunda passada: decide "em dia"/"dias de atraso" — pela periodicidade
    # real quando ela é conhecida (AR CONDICIONADOS/BEBEDOUROS), senão cai
    # de volta pra regra simplificada de "teve preventiva este mês"
    # (hoje é sempre o caso dos chillers, que não têm template de PMOC).
    for info in por_equip.values():
        periodicidade = info["periodicidade"]
        if periodicidade and periodicidade in INTERVALO_DIAS_POR_PERIODICIDADE:
            intervalo = INTERVALO_DIAS_POR_PERIODICIDADE[periodicidade]
            if info["ultima_preventiva"] is None:
                info["em_dia"] = False
                info["dias_atraso"] = None
            else:
                dias_desde = (agora - info["ultima_preventiva"]).days
                info["em_dia"] = dias_desde <= intervalo
                info["dias_atraso"] = max(0, dias_desde - intervalo)
        else:
            info["em_dia"] = info["_realizado_mes_atual"]
            info["dias_atraso"] = None
        del info["_realizado_mes_atual"]

    return list(por_equip.values())


def enviar_equipamentos_para_supabase(registros: list[dict]) -> None:
    headers = {
        "apikey": SUPABASE_SERVICE_ROLE_KEY,
        "Authorization": f"Bearer {SUPABASE_SERVICE_ROLE_KEY}",
        "Content-Type": "application/json",
    }
    tabela_url = f"{SUPABASE_URL}/rest/v1/kpi_manutencao_equipamentos"

    resp = requests.delete(f"{tabela_url}?id=not.is.null", headers=headers)
    resp.raise_for_status()

    payload = [
        {
            "categoria": r["categoria"],
            "equipamento": r["equipamento"],
            "em_dia": r["em_dia"],
            "dias_atraso": r["dias_atraso"],
            "periodicidade": r["periodicidade"],
            "ultima_preventiva": r["ultima_preventiva"].date().isoformat()
            if r["ultima_preventiva"]
            else None,
            "ultima_substituicao_filtro": r["ultima_substituicao_filtro"].date().isoformat()
            if r["ultima_substituicao_filtro"]
            else None,
            "proxima_substituicao_filtro": r["proxima_substituicao_filtro"].date().isoformat()
            if r["proxima_substituicao_filtro"]
            else None,
        }
        for r in registros
    ]

    tamanho_lote = 200
    for i in range(0, len(payload), tamanho_lote):
        lote = payload[i : i + tamanho_lote]
        resp = requests.post(tabela_url, headers=headers, json=lote)
        resp.raise_for_status()


def main():
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--mes-referencia",
        help="Simula outra data pra calcular realizado/pendente do mês "
        "(formato AAAA-MM-DD). Sem isso, usa a data de hoje de verdade — "
        "só usar pra pré-visualizar o painel, depois rodar de novo sem a "
        "flag pra restaurar os dados reais.",
    )
    args = parser.parse_args()
    data_referencia = (
        datetime.strptime(args.mes_referencia, "%Y-%m-%d") if args.mes_referencia else None
    )
    if data_referencia:
        print(f"⚠ SIMULANDO data de referência: {data_referencia:%d/%m/%Y} (não é a data real)")

    if os.path.exists(PASTA_TEMP):
        shutil.rmtree(PASTA_TEMP)
    os.makedirs(PASTA_TEMP, exist_ok=True)

    # Roda em lotes (ex: GitHub Actions com timeout curto): processa no
    # máximo N arquivos novos/alterados por execução e para — quem já está
    # em dia com o cache não conta pro limite (é só releitura, instantâneo).
    # Só publica no Supabase quando a lista INTEIRA foi percorrida nesta
    # execução; senão a gente arriscaria apagar o banco com dado pela
    # metade no meio da varredura. Com o cache persistido entre execuções
    # (ex: actions/cache), a próxima retoma de onde parou.
    max_novos = int(os.environ.get("KPI_MAX_NOVOS_POR_EXECUCAO", "0") or "0")

    client = conectar_nextcloud()
    cache_anterior = carregar_cache(CACHE_PATH)

    print(f"Listando '{PASTA_RAIZ}' (uma chamada recursiva só)...")
    arvore = listar_arvore(client, PASTA_RAIZ)
    pdfs = extrair_pdfs(arvore, PASTA_RAIZ)
    estrutura = extrair_estrutura(arvore, PASTA_RAIZ)
    print(f"{len(pdfs)} PDF(s) encontrado(s).")

    atendimentos: list[Atendimento] = []
    pares: list[tuple[str, Atendimento]] = []
    # Começa com o cache antigo inteiro — se a execução parar no meio do
    # lote, o que não foi visitado ainda continua salvo (não é substituído
    # por um cache_novo vazio/parcial).
    cache_novo: dict = dict(cache_anterior)
    reaproveitados = 0
    novos_processados = 0
    execucao_completa = True
    for i, item in enumerate(pdfs, 1):
        caminho, etag = item["path"], item["etag"]
        nome = caminho.rstrip("/").split("/")[-1]
        entrada_cache = cache_anterior.get(caminho)

        if (
            entrada_cache
            and entrada_cache.get("etag") == etag
            and entrada_cache.get("versao") == CACHE_VERSAO
        ):
            a = dict_para_atendimento(entrada_cache["dados"])
            atendimentos.append(a)
            pares.append((caminho, a))
            reaproveitados += 1
            continue

        if max_novos and novos_processados >= max_novos:
            execucao_completa = False
            print(
                f"\nLimite de {max_novos} novo(s)/alterado(s) atingido "
                f"({i - 1}/{len(pdfs)} arquivos vistos) — parando por aqui, "
                "a próxima execução continua de onde parou."
            )
            break

        print(f"  [{i}/{len(pdfs)}] {nome} (novo/alterado/regra atualizada)")
        try:
            a = processar_pdf(client, caminho, PASTA_RAIZ)
            atendimentos.append(a)
            pares.append((caminho, a))
            cache_novo[caminho] = {
                "etag": etag,
                "versao": CACHE_VERSAO,
                "dados": atendimento_para_dict(a),
            }
            novos_processados += 1
        except Exception as e:
            print(f"    [erro] não processado: {e}")

    salvar_cache(cache_novo, CACHE_PATH)
    shutil.rmtree(PASTA_TEMP, ignore_errors=True)

    print(f"\n{novos_processados} processado(s) agora, {reaproveitados} reaproveitado(s) do cache.")

    if not execucao_completa:
        print(
            "Execução parcial — não publica no Supabase ainda (lista não foi "
            "percorrida por inteiro). Rode de novo pra continuar o lote."
        )
        return

    print(f"Gerando relatório: {SAIDA_DOCX}")
    gerar_relatorio(atendimentos, SAIDA_DOCX)

    print("Publicando no Supabase...")
    enviar_para_supabase(pares)

    resumo_equipamentos = montar_resumo_equipamentos(estrutura, pares, data_referencia)
    enviar_equipamentos_para_supabase(resumo_equipamentos)

    print("Concluído.")


if __name__ == "__main__":
    main()
