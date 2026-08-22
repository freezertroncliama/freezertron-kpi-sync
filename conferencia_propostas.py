# -*- coding: utf-8 -*-
"""
Conferência diária: pasta de PDFs de propostas (Nextcloud, "ORÇAMENTOS OMIE")
× propostas_geradas (Supabase).

Achado em 22/08/2026 (sessão com o Rafael): dá pra imprimir o preview de uma
proposta (Ctrl+P direto no navegador) SEM clicar em "Concluir e Imprimir" —
isso gera um PDF que não bate com nada no banco, com um número que mais
tarde pode ser reaproveitado de verdade por outro cliente. Esse script acha
esses casos automaticamente, todo dia, sem precisar o Rafael pedir na mão.

Não mexe na Omie (só pasta × banco) — a comparação com a Omie já é feita sob
demanda pela tela /relatorio-comercial do app; repetir isso aqui todo dia só
aumentaria risco de rate limit da Omie sem necessidade.

Grava o resultado na tabela `conferencia_propostas` (migration
0015_conferencia_propostas.sql) e, se achar qualquer anomalia, tenta mandar
um aviso pelo WhatsApp (best-effort — se a Meta ainda não tiver um template
aprovado pra mensagem fora da janela de 24h, essa parte falha sozinha e só
fica registrado; a tela do app sempre mostra o resultado, com ou sem
WhatsApp).

Uso:
  python conferencia_propostas.py
"""

import os
import re
import sys
from datetime import datetime, timedelta

import requests
from dotenv import load_dotenv
from webdav3.client import Client

sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

load_dotenv()

NEXTCLOUD_URL = os.environ["NEXTCLOUD_URL"]
NEXTCLOUD_USER = os.environ["NEXTCLOUD_USER"]
NEXTCLOUD_APP_PASSWORD = os.environ["NEXTCLOUD_APP_PASSWORD"]
SUPABASE_URL = os.environ["SUPABASE_URL"]
SUPABASE_SERVICE_ROLE_KEY = os.environ["SUPABASE_SERVICE_ROLE_KEY"]

# Opcionais — sem eles o script roda igual, só pula o envio de WhatsApp.
WHATSAPP_ACCESS_TOKEN = os.environ.get("WHATSAPP_ACCESS_TOKEN")
WHATSAPP_PHONE_NUMBER_ID = os.environ.get("WHATSAPP_PHONE_NUMBER_ID")
WHATSAPP_DESTINO = os.environ.get("WHATSAPP_DESTINO")  # ex: "5535991240254"

PASTA = "ORÇAMENTOS OMIE"
# Só cobra PDF ausente de proposta recente — cobrar propostas antigas geraria
# ruído de coisa que já foi arquivada/impressa fora desse fluxo, sem
# relevância prática hoje.
DIAS_JANELA_SEM_PDF = 30

REGEX_COM_CLIENTE = re.compile(
    r"^Proposta N[ºo]\s*(\d+)\s*[—-]\s*(.+?)\s*[—-]\s*Freezertron\.pdf$", re.IGNORECASE
)
REGEX_SEM_CLIENTE = re.compile(r"^Proposta N[ºo]\s*(\d+)\s*[—-]\s*Freezertron\.pdf$", re.IGNORECASE)


def montar_client_webdav() -> Client:
    options = {
        "webdav_hostname": f"{NEXTCLOUD_URL}/remote.php/dav/files/{NEXTCLOUD_USER}",
        "webdav_login": NEXTCLOUD_USER,
        "webdav_password": NEXTCLOUD_APP_PASSWORD,
    }
    return Client(options)


def listar_pdfs(client: Client) -> list[str]:
    itens = client.list(PASTA, get_info=True)
    nomes = []
    for item in itens:
        caminho = item.get("path", "")
        nome = caminho.rstrip("/").split("/")[-1]
        if nome.strip().lower().endswith(".pdf"):
            nomes.append(nome)
    return nomes


def parsear_pdfs(nomes: list[str]) -> list[dict]:
    resultado = []
    for nome in nomes:
        m = REGEX_COM_CLIENTE.match(nome)
        if m:
            resultado.append({"numero": int(m.group(1)), "cliente": m.group(2).strip(), "arquivo": nome})
            continue
        m2 = REGEX_SEM_CLIENTE.match(nome)
        if m2:
            resultado.append({"numero": int(m2.group(1)), "cliente": "", "arquivo": nome})
    return resultado


def buscar_propostas() -> list[dict]:
    headers = {"apikey": SUPABASE_SERVICE_ROLE_KEY, "Authorization": f"Bearer {SUPABASE_SERVICE_ROLE_KEY}"}
    resp = requests.get(
        f"{SUPABASE_URL}/rest/v1/propostas_geradas",
        headers=headers,
        params={"select": "numero,created_at,cliente"},
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()


def nomes_batem(nome_pdf: str, nome_banco: str) -> bool:
    a = (nome_pdf or "").upper().strip()
    b = (nome_banco or "").upper().strip()
    if not a or not b:
        return False
    return a in b or b in a


def rodar_conferencia():
    client = montar_client_webdav()
    pdfs = parsear_pdfs(listar_pdfs(client))
    propostas = buscar_propostas()

    por_numero: dict[int, list[dict]] = {}
    for p in propostas:
        por_numero.setdefault(p["numero"], []).append(p)

    pdf_sem_proposta = []
    numero_conflitante = []
    numeros_com_pdf = set()

    for pdf in pdfs:
        numeros_com_pdf.add(pdf["numero"])
        candidatos = por_numero.get(pdf["numero"], [])
        if not candidatos:
            pdf_sem_proposta.append(pdf)
            continue
        if pdf["cliente"] and not any(
            nomes_batem(pdf["cliente"], c["cliente"].get("nome", "")) for c in candidatos
        ):
            numero_conflitante.append(
                {**pdf, "cliente_no_banco": candidatos[0]["cliente"].get("nome", "")}
            )

    limite = datetime.utcnow() - timedelta(days=DIAS_JANELA_SEM_PDF)
    proposta_sem_pdf = []
    for p in propostas:
        criado = datetime.fromisoformat(p["created_at"].replace("Z", "+00:00")).replace(tzinfo=None)
        if criado < limite:
            continue
        if p["numero"] not in numeros_com_pdf:
            proposta_sem_pdf.append(
                {
                    "numero": p["numero"],
                    "cliente": p["cliente"].get("nome", ""),
                    "created_at": p["created_at"],
                }
            )

    return pdf_sem_proposta, numero_conflitante, proposta_sem_pdf


def gravar_resultado(pdf_sem_proposta, numero_conflitante, proposta_sem_pdf):
    total = len(pdf_sem_proposta) + len(numero_conflitante) + len(proposta_sem_pdf)
    headers = {
        "apikey": SUPABASE_SERVICE_ROLE_KEY,
        "Authorization": f"Bearer {SUPABASE_SERVICE_ROLE_KEY}",
        "Content-Type": "application/json",
    }
    body = {
        "pdf_sem_proposta": pdf_sem_proposta,
        "numero_conflitante": numero_conflitante,
        "proposta_sem_pdf": proposta_sem_pdf,
        "total_anomalias": total,
    }
    resp = requests.post(
        f"{SUPABASE_URL}/rest/v1/conferencia_propostas", headers=headers, json=body, timeout=30
    )
    resp.raise_for_status()
    return total


def montar_mensagem(pdf_sem_proposta, numero_conflitante, proposta_sem_pdf) -> str:
    linhas = ["*Conferência de propostas — achados de hoje*"]
    if pdf_sem_proposta:
        linhas.append(f"\n📄 {len(pdf_sem_proposta)} PDF(s) impresso(s) sem proposta salva no banco:")
        for x in pdf_sem_proposta[:5]:
            linhas.append(f"  • Nº {x['numero']} — {x['arquivo']}")
    if numero_conflitante:
        linhas.append(f"\n⚠️ {len(numero_conflitante)} número(s) usado(s) por clientes diferentes:")
        for x in numero_conflitante[:5]:
            linhas.append(f"  • Nº {x['numero']}: PDF diz \"{x['cliente']}\", banco tem \"{x['cliente_no_banco']}\"")
    if proposta_sem_pdf:
        linhas.append(f"\n📁 {len(proposta_sem_pdf)} proposta(s) recente(s) sem PDF salvo na pasta:")
        for x in proposta_sem_pdf[:5]:
            linhas.append(f"  • Nº {x['numero']} — {x['cliente']}")
    linhas.append("\nDetalhes completos na tela /conferencia-propostas do sistema.")
    return "\n".join(linhas)


def enviar_whatsapp(texto: str):
    if not (WHATSAPP_ACCESS_TOKEN and WHATSAPP_PHONE_NUMBER_ID and WHATSAPP_DESTINO):
        print("WhatsApp não configurado (faltam variáveis) — pulando envio.")
        return
    resp = requests.post(
        f"https://graph.facebook.com/v21.0/{WHATSAPP_PHONE_NUMBER_ID}/messages",
        headers={"Authorization": f"Bearer {WHATSAPP_ACCESS_TOKEN}", "Content-Type": "application/json"},
        json={
            "messaging_product": "whatsapp",
            "to": WHATSAPP_DESTINO,
            "type": "text",
            "text": {"body": texto},
        },
        timeout=30,
    )
    if not resp.ok:
        print(f"Falha ao enviar WhatsApp ({resp.status_code}): {resp.text}")
    else:
        print("WhatsApp enviado com sucesso.")


def main():
    print("Conferência de propostas — iniciando...")
    pdf_sem_proposta, numero_conflitante, proposta_sem_pdf = rodar_conferencia()
    total = gravar_resultado(pdf_sem_proposta, numero_conflitante, proposta_sem_pdf)
    print(f"Concluído. Total de anomalias encontradas: {total}")

    if total > 0:
        try:
            enviar_whatsapp(montar_mensagem(pdf_sem_proposta, numero_conflitante, proposta_sem_pdf))
        except Exception as err:  # best-effort — nunca deve quebrar o script
            print(f"Erro ao tentar enviar WhatsApp: {err}")
    else:
        print("Nenhuma anomalia — não envia WhatsApp hoje.")


if __name__ == "__main__":
    main()
