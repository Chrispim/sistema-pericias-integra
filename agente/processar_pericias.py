"""
processar_pericias.py — Versão Otimizada
Integra Contabilidade & Perícias — Chrispim Gonçalves

Otimizações implementadas:
  1. Seleção de modelo por tarefa  (Haiku → classificar / Sonnet → gerar)
  2. Prompt Caching                (contexto fixo cacheado, -90% custo repetido)
  3. max_tokens controlado         (50 para classificar, 600 para roteiros)
  4. Batch API                     (-50% custo no processamento em lote)
  5. Prefill da resposta           (força JSON direto, elimina tokens de aquecimento)
  6. Saída estruturada             (sem narrativa, só o necessário)
"""

import json, os, time
from datetime import datetime

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")  # set via env var
EXCEL_PATH = r"G:\Meu Drive\Pessoal\Pericia Contabil\Relatorio_Pericias_Atualizado_Atual.xlsx"
OUTPUT_PATHS = [
    r"G:\Meu Drive\Pessoal\Pericia Contabil\agente_pericias\ultima_verificacao_intimacoes.json",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "ultima_verificacao_intimacoes.json"),
]

# Modelos — escolha por custo/capacidade
MODEL_CLASSIFICAR = "claude-haiku-4-5"    # $0.80/MTok input  — classificar, extrair
MODEL_GERAR       = "claude-sonnet-4-6"   # $3.00/MTok input  — roteiros, resumos

# ─────────────────────────────────────────────
# CONTEXTO FIXO CACHEADO (pago 1x, reutilizado)
# ─────────────────────────────────────────────
CONTEXTO_PERITO = """Você é assistente de Chrispim Gonçalves, perito contábil judicial.
Escritório: Integra Contabilidade & Perícias — Vitória/ES.
Tribunais: TJES (J=8, TT=08) e TRE-ES (J=6, TT=08).
Formato CNJ: NNNNNNN-DD.AAAA.J.TT.OOOO

Regras de urgência:
- URGENTE: laudo pendente, quesitos não respondidos, ≥3 notificações, prazo ≤5 dias
- ATENÇÃO: laudo entregue aguardando homologação, pendente de recebimento
- INFORMATIVO: demais movimentações

Ao classificar: responda APENAS em JSON válido, sem texto antes ou depois.
Ao gerar roteiro: seja objetivo, use lista numerada, máximo 6 itens."""

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def normalize_cnj(cnj: str) -> str:
    return cnj.replace("-", "").replace(".", "").replace(" ", "")

def ler_excel() -> list:
    planilha = []
    try:
        import openpyxl
        wb = openpyxl.load_workbook(EXCEL_PATH, read_only=True, data_only=True)
        ws = next(
            (wb[s] for s in wb.sheetnames if any(k in s for k in ["Fonte", "Laudo", "dados"])),
            wb.active,
        )
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i == 0:
                continue
            try:
                cnj = str(row[2]).strip() if row[2] else ""
                if cnj and cnj != "None":
                    planilha.append({
                        "cnj":     cnj,
                        "vara":    str(row[1]  or ""),
                        "autor":   str(row[3]  or ""),
                        "status":  str(row[8]  or ""),
                        "pendente":str(row[11] or ""),
                    })
            except Exception:
                pass
        wb.close()
        print(f"Excel lido: {len(planilha)} registros")
    except Exception as e:
        print(f"AVISO Excel: {e}")
    return planilha

def classificar_sem_api(email: dict, planilha_item: dict | None) -> str:
    """Classificação rápida por regras — zero tokens, zero custo."""
    if not planilha_item:
        return "sem_correspondencia"
    s = planilha_item["status"].lower()
    if "laudo pendente" in s or "quesita" in s or email["qtd"] >= 3:
        return "urgente"
    if "laudo entregue" in s or "pendente recebimento" in s:
        return "atencao"
    return "informativo"

# ─────────────────────────────────────────────
# CLAUDE API — OTIMIZADA
# ─────────────────────────────────────────────
def get_client():
    if not ANTHROPIC_API_KEY:
        return None
    try:
        import anthropic
        return anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    except ImportError:
        print("AVISO: anthropic não instalado. pip install anthropic")
        return None

def enriquecer_classificacao_haiku(client, processo: dict) -> dict:
    """
    Otimização 1: Haiku para classificação — barato e rápido.
    Otimização 2: Prompt caching no contexto fixo.
    Otimização 3: max_tokens=80 — só precisa do JSON de classificação.
    Otimização 5: Prefill '{"' força resposta JSON sem aquecimento.
    """
    prompt = (
        f"Processo: {processo['cnj']} | Tribunal: {processo['tribunal']}\n"
        f"Status atual: {processo.get('status','')}\n"
        f"Notificações: {processo['qtd']} | Última: {processo['ultima_data']}\n"
        f"Pendente: {processo.get('pendente','')}\n\n"
        f'Responda APENAS: {{"urgencia":"urgente|atencao|informativo","prazo_risco":true|false,"acao_imediata":"max 8 palavras"}}'
    )
    try:
        resp = client.messages.create(
            model=MODEL_CLASSIFICAR,
            max_tokens=80,                          # Otimização 3
            system=[
                {
                    "type": "text",
                    "text": CONTEXTO_PERITO,
                    "cache_control": {"type": "ephemeral"},  # Otimização 2
                }
            ],
            messages=[
                {"role": "user",      "content": prompt},
                {"role": "assistant", "content": '{"'},      # Otimização 5 — prefill
            ],
        )
        raw = '{"' + resp.content[0].text
        return json.loads(raw)
    except Exception as e:
        print(f"  Haiku erro ({processo['cnj'][:20]}): {e}")
        return {}

def gerar_roteiro_sonnet(client, processo: dict) -> str:
    """
    Otimização 1: Sonnet só para urgentes — onde o custo vale a pena.
    Otimização 2: Caching no contexto.
    Otimização 3: max_tokens=600 — roteiro cabe em 400 tokens facilmente.
    Otimização 6: Saída estruturada, sem introdução.
    """
    prompt = (
        f"Processo: {processo['cnj']} | {processo['tribunal']}\n"
        f"Autor: {processo.get('autor','')}\n"
        f"Vara: {processo.get('vara','')}\n"
        f"Status: {processo.get('status','')}\n"
        f"Notificações recentes: {processo['qtd']} (última: {processo['ultima_data']})\n"
        f"Pendências: {processo.get('pendente','')}\n\n"
        f"Gere roteiro de ação pericial em lista numerada. Seja direto, sem introdução."
    )
    try:
        resp = client.messages.create(
            model=MODEL_GERAR,
            max_tokens=600,                          # Otimização 3
            system=[
                {
                    "type": "text",
                    "text": CONTEXTO_PERITO,
                    "cache_control": {"type": "ephemeral"},  # Otimização 2
                }
            ],
            messages=[{"role": "user", "content": prompt}],
        )
        return resp.content[0].text.strip()
    except Exception as e:
        print(f"  Sonnet erro ({processo['cnj'][:20]}): {e}")
        return ""

def enriquecer_em_lote_batch(client, urgentes: list) -> dict:
    """
    Otimização 4: Batch API — 50% desconto, processa todos de uma vez.
    Retorna dict {cnj: roteiro}.
    Ideal para rodar no cron de madrugada.
    """
    if not urgentes:
        return {}

    requests = []
    for p in urgentes:
        prompt = (
            f"Processo: {p['cnj']} | {p['tribunal']}\n"
            f"Autor: {p.get('autor','')} | Vara: {p.get('vara','')}\n"
            f"Status: {p.get('status','')} | Notificações: {p['qtd']}\n"
            f"Pendências: {p.get('pendente','')}\n\n"
            f"Roteiro de ação pericial em lista numerada. Sem introdução."
        )
        requests.append({
            "custom_id": p["cnj"],
            "params": {
                "model": MODEL_GERAR,
                "max_tokens": 600,                   # Otimização 3
                "system": [
                    {
                        "type": "text",
                        "text": CONTEXTO_PERITO,
                        "cache_control": {"type": "ephemeral"},  # Otimização 2
                    }
                ],
                "messages": [{"role": "user", "content": prompt}],
            },
        })

    print(f"  Batch: enviando {len(requests)} processos urgentes...")
    try:
        batch = client.beta.messages.batches.create(requests=requests)
        print(f"  Batch ID: {batch.id} | Status: {batch.processing_status}")

        # Aguarda conclusão (max 10 min)
        for _ in range(60):
            time.sleep(10)
            batch = client.beta.messages.batches.retrieve(batch.id)
            if batch.processing_status == "ended":
                break
            print(f"  Aguardando batch... {batch.request_counts}")

        # Coleta resultados
        roteiros = {}
        for result in client.beta.messages.batches.results(batch.id):
            if result.result.type == "succeeded":
                roteiros[result.custom_id] = result.result.message.content[0].text.strip()
        print(f"  Batch concluído: {len(roteiros)} roteiros gerados")
        return roteiros
    except Exception as e:
        print(f"  Batch erro: {e}")
        return {}

# ─────────────────────────────────────────────
# PIPELINE PRINCIPAL
# ─────────────────────────────────────────────
def processar(usar_api: bool = True, usar_batch: bool = False):
    """
    usar_api=True  → enriquece com Haiku (classificação) + Sonnet (roteiros urgentes)
    usar_batch=True → usa Batch API para roteiros (-50% custo, mas demora ~5 min)
    """
    print(f"=== Integra Perícias — {datetime.now().strftime('%d/%m/%Y %H:%M')} ===")

    # 1. Dados de entrada (em produção: viriam do Gmail MCP)
    emails_data = [
        {"cnj": "0004037-40.2015.8.08.0004", "tribunal": "TJES",   "qtd": 1, "ultima_data": "2026-05-01"},
        {"cnj": "0003454-21.2016.8.08.0004", "tribunal": "TJES",   "qtd": 3, "ultima_data": "2026-05-01"},
        {"cnj": "5004093-24.2021.8.08.0021", "tribunal": "TJES",   "qtd": 3, "ultima_data": "2026-04-30"},
        {"cnj": "0000783-88.2017.8.08.0004", "tribunal": "TJES",   "qtd": 5, "ultima_data": "2026-04-30"},
        {"cnj": "0600276-54.2024.6.08.0019", "tribunal": "TRE-ES", "qtd": 5, "ultima_data": "2026-04-29"},
        {"cnj": "5001026-68.2022.8.08.0004", "tribunal": "TJES",   "qtd": 1, "ultima_data": "2026-04-29"},
        {"cnj": "5000931-38.2022.8.08.0004", "tribunal": "TJES",   "qtd": 2, "ultima_data": "2026-04-29"},
        {"cnj": "0002045-10.2016.8.08.0004", "tribunal": "TJES",   "qtd": 1, "ultima_data": "2026-04-27"},
        {"cnj": "0000932-26.2013.8.08.0004", "tribunal": "TJES",   "qtd": 1, "ultima_data": "2026-04-27"},
    ]

    # 2. Lê planilha Excel
    planilha = ler_excel()

    # 3. Enriquece cada email com dados da planilha
    processos = []
    for email in emails_data:
        found = next(
            (p for p in planilha if normalize_cnj(p["cnj"]) == normalize_cnj(email["cnj"])),
            None,
        )
        item = {
            **email,
            "autor":            found["autor"]   if found else "",
            "vara":             found["vara"]    if found else "",
            "status":           found["status"]  if found else "",
            "pendente":         found["pendente"]if found else "",
            "qtd_notificacoes": email["qtd"],
        }
        # Classificação rápida sem API (sempre roda — custo zero)
        item["urgencia_regra"] = classificar_sem_api(email, found)
        processos.append(item)

    # 4. Enriquecimento com API (opcional)
    client = get_client() if usar_api else None
    tokens_usados = {"haiku_input": 0, "haiku_output": 0, "sonnet_input": 0, "sonnet_output": 0}

    if client:
        print(f"\n[API] Classificando {len(processos)} processos com Haiku...")
        for p in processos:
            enrich = enriquecer_classificacao_haiku(client, p)  # ~50 tokens output
            if enrich:
                p["urgencia_api"]    = enrich.get("urgencia", p["urgencia_regra"])
                p["prazo_risco"]     = enrich.get("prazo_risco", False)
                p["acao_imediata"]   = enrich.get("acao_imediata", "")
            else:
                p["urgencia_api"]    = p["urgencia_regra"]

        # Urgentes: gera roteiros
        urgentes_api = [p for p in processos if p.get("urgencia_api") == "urgente"]
        print(f"[API] {len(urgentes_api)} processos urgentes — gerando roteiros...")

        if usar_batch and urgentes_api:
            # Batch API: -50% custo, ideal para cron noturno
            roteiros = enriquecer_em_lote_batch(client, urgentes_api)
            for p in urgentes_api:
                p["roteiro"] = roteiros.get(p["cnj"], "")
        else:
            # Síncrono: para uso interativo
            for p in urgentes_api:
                p["roteiro"] = gerar_roteiro_sonnet(client, p)  # max 600 tokens

    # 5. Agrupa por urgência (usa urgencia_api se disponível, senão urgencia_regra)
    urgentes      = [p for p in processos if p.get("urgencia_api", p["urgencia_regra"]) == "urgente"]
    atencao       = [p for p in processos if p.get("urgencia_api", p["urgencia_regra"]) == "atencao"]
    informativos  = [p for p in processos if p.get("urgencia_api", p["urgencia_regra"]) == "informativo"]
    sem_corresp   = [p["cnj"] for p in processos if p["urgencia_regra"] == "sem_correspondencia"]

    # 6. Resultado final
    resultado = {
        "data_verificacao":    datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
        "total_movimentacoes": sum(e["qtd"] for e in emails_data),
        "total_processos":     len(emails_data),
        "excel_lido":          len(planilha) > 0,
        "api_utilizada":       client is not None,
        "urgentes":            urgentes,
        "atencao":             atencao,
        "informativos":        informativos,
        "sem_correspondencia": sem_corresp,
    }

    # 7. Salva JSON
    for path in OUTPUT_PATHS:
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(path, "w", encoding="utf-8") as f:
                json.dump(resultado, f, ensure_ascii=False, indent=2)
            print(f"\nSalvo: {path}")
            break
        except Exception as e:
            print(f"Erro ao salvar {path}: {e}")

    print(f"\n{'='*50}")
    print(f"Urgentes: {len(urgentes)} | Atenção: {len(atencao)} | Info: {len(informativos)}")
    print(f"Sem correspondência: {len(sem_corresp)}")
    print("=== CONCLUIDO ===")
    return resultado


# ─────────────────────────────────────────────
# EXECUÇÃO
# ─────────────────────────────────────────────
if __name__ == "__main__":
    # usar_api=False  → roda sem API, só regras (igual versão anterior)
    # usar_api=True   → enriquece com Haiku + Sonnet
    # usar_batch=True → usa Batch API (-50% custo) — ideal para cron
    processar(
        usar_api=bool(ANTHROPIC_API_KEY),
        usar_batch=False,   # mude para True no cron noturno
    )
