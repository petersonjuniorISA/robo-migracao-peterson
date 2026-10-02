from __future__ import annotations

"""
ROBO MIGRACAO PETERSON - DEFINITIVO

Google Sheets + WhatsApp Business Desktop.

Regras da fila:
- P = Peterson
- U = vazio ou '-'
- O sem andamento anterior
- profissional e telefone precisam existir

Fluxo seguro:
- abre a conversa por deep-link
- confirma o contato
- preenche a mensagem por UI Automation (sem mover mouse)
- aciona o botao Enviar por UI Automation / teclado como fallback
- confirma pela aparicao do texto da mensagem na conversa
- SOMENTE depois atualiza O para exatamente 'Em andamento'
- tenta aplicar a etiqueta 'Peterson Migracao' via 'Adicionar a lista'

Observacao: a confirmacao nunca usa somente "campo vazio" como prova de envio,
porque isso gerou falso positivo em uma versao anterior.
"""

import argparse
import hashlib
import json
import re
import subprocess
import time
from datetime import datetime
from pathlib import Path
from typing import Any

BASE = Path(__file__).resolve().parent
CREDENTIALS_FILE = BASE / "credentials.json"
TOKEN_FILE = BASE / "token.json"
STATE_FILE = BASE / "estado_robo_peterson.json"
LOG_FILE = BASE / "robo_migracao_peterson.log"

SPREADSHEET_ID = "1vO_GkgBptSLoJKe7jkNttTJ5xSKkO3QiGD-hFl8Z5Ks"
SHEET_NAME = "Aprovados PHC"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

COL_A = 1
COL_M = 13
COL_N = 14
COL_O = 15
COL_P = 16
COL_T = 20
COL_U = 21

INTERVALO_ENTRE_CONTATOS = 2.0
ESPERA_DEEP_LINK = 1.2
TIMEOUT_ABRIR = 5.0
ESPERA_COLAR = 0.35
TIMEOUT_ENVIO = 7.0
ESPERA_MENU = 0.3
NOME_ETIQUETA = "Peterson Migração"

MENSAGEM_CADASTRADO = """Olá! Aqui é da ISA 💙

Estou entrando em contato para te avisar sobre uma novidade em relação ao atendimento do(a) {paciente}.

Esse atendimento passará por uma migração para o modelo ISA a partir do dia 15/10.

Até o dia 14/10, os atendimentos continuam acontecendo normalmente pela cooperativa. A partir de 15/10, o atendimento passa a ser realizado pela ISA.

Você consegue me confirmar, por favor, que recebeu essa mensagem e está ciente dessa mudança? Assim que me confirmar, podemos dar continuidade ao processo da sua migração. 💙"""

MENSAGEM_NAO_CADASTRADO = """Olá! Aqui é da ISA 💙

Estou entrando em contato sobre o atendimento do(a) {paciente}, que passará a ser realizado pela ISA a partir do dia 15/10.

Até o dia 14/10, os atendimentos continuam acontecendo normalmente pela cooperativa. A partir de 15/10, o atendimento passa a ser realizado pela ISA.

Para conseguirmos dar continuidade à sua migração, é necessário realizar o cadastro no App ISA Atende:

📱 iPhone:
https://apps.apple.com/br/app/isa-atende/id6449265549

📱 Android:
https://play.google.com/store/search?q=isa%20atende&c=apps

Após instalar o aplicativo, faça seu cadastro e login com seu e-mail e senha.

Também é necessário realizar o curso de Metas Internacionais antes de iniciar os atendimentos pela ISA:

https://isasaude.premia360.com/inscricao/6a0f3a23d3ab9

Você consegue me confirmar, por favor, que recebeu essa mensagem e que irá realizar o cadastro? Assim que me confirmar, podemos dar continuidade ao processo da sua migração. 💙"""


def agora() -> str:
    return datetime.now().strftime("%d/%m/%Y %H:%M:%S")


def log(msg: str) -> None:
    linha = f"[{agora()}] {msg}"
    print(linha)
    try:
        with LOG_FILE.open("a", encoding="utf-8") as f:
            f.write(linha + "\n")
    except OSError:
        pass


def texto(v: Any) -> str:
    return str(v if v is not None else "").strip()


def normalizar_telefone(v: Any) -> str:
    d = re.sub(r"\D", "", texto(v))
    if len(d) in (10, 11):
        d = "55" + d
    return d if d.startswith("55") and len(d) in (12, 13) else ""


def valor_col(row: list[Any], col: int) -> Any:
    i = col - 1
    return row[i] if i < len(row) else ""


def paciente_bloco(data: list[list[Any]], idx: int) -> str:
    for i in range(idx, 2, -1):
        value = texto(valor_col(data[i], COL_A))
        if value:
            return value
    return ""


def o_sem_andamento(v: Any) -> bool:
    o = texto(v).lower()
    return not any(
        t in o
        for t in (
            "contato",
            "em andamento",
            "aguardando",
            "resposta",
            "migrado",
            "revisão",
            "revisao",
        )
    )


def elegivel(row: list[Any]) -> bool:
    return (
        texto(valor_col(row, COL_P)).lower() == "peterson"
        and texto(valor_col(row, COL_U)).lower() in ("", "-")
        and o_sem_andamento(valor_col(row, COL_O))
        and bool(texto(valor_col(row, COL_M)))
        and bool(normalizar_telefone(valor_col(row, COL_N)))
    )


def montar_fila(data: list[list[Any]]) -> list[dict[str, Any]]:
    fila: list[dict[str, Any]] = []
    for idx in range(3, len(data)):
        row = data[idx]
        if not elegivel(row):
            continue
        fila.append(
            {
                "linha": idx + 1,
                "paciente": paciente_bloco(data, idx),
                "profissional": texto(valor_col(row, COL_M)),
                "telefone": normalizar_telefone(valor_col(row, COL_N)),
                "conta": texto(valor_col(row, COL_T)),
                "contato": texto(valor_col(row, COL_O)),
            }
        )
    return fila


def mensagem_para(item: dict[str, Any]) -> str:
    paciente = item["paciente"] or "seu atendimento"
    conta = item["conta"].lower()
    cadastrado = (
        "cadastrado" in conta
        and "não cadastrado" not in conta
        and "nao cadastrado" not in conta
    )
    if cadastrado:
        return MENSAGEM_CADASTRADO.format(paciente=paciente)
    return MENSAGEM_NAO_CADASTRADO.format(paciente=paciente)


# =========================
# Google Sheets
# =========================


def google_api():
    try:
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from googleapiclient.discovery import build
    except ImportError as exc:
        raise RuntimeError(
            "Bibliotecas Google ausentes. Rode o BAT do robo."
        ) from exc

    creds = None
    if TOKEN_FILE.exists():
        try:
            creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
        except Exception:
            creds = None

    if creds and creds.valid:
        pass
    elif creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        if not CREDENTIALS_FILE.exists():
            raise FileNotFoundError("credentials.json não encontrado.")
        flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_FILE), SCOPES)
        creds = flow.run_local_server(port=0, open_browser=True)

    TOKEN_FILE.write_text(creds.to_json(), encoding="utf-8")
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


def ler_planilha(api) -> list[list[Any]]:
    return (
        api.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{SHEET_NAME}'!A:U",
            valueRenderOption="UNFORMATTED_VALUE",
        ).execute().get("values", [])
    )


def atualizar_o_lote(api, linhas: list[int]) -> None:
    if not linhas:
        return
    data = [
        {"range": f"'{SHEET_NAME}'!O{linha}", "values": [["Em andamento"]]}
        for linha in linhas
    ]
    api.spreadsheets().values().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"valueInputOption": "USER_ENTERED", "data": data},
    ).execute()


def corrigir_o_antigo(api) -> int:
    data = (
        api.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{SHEET_NAME}'!O:O",
            valueRenderOption="UNFORMATTED_VALUE",
        ).execute().get("values", [])
    )
    updates = []
    for linha, row in enumerate(data, 1):
        value = texto(row[0] if row else "")
        if value.lower().startswith("em andamento -"):
            updates.append({"range": f"'{SHEET_NAME}'!O{linha}", "values": [["Em andamento"]]})
    if updates:
        api.spreadsheets().values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"valueInputOption": "USER_ENTERED", "data": updates},
        ).execute()
    return len(updates)


def status(api) -> dict[str, int]:
    data = ler_planilha(api)
    total = migrados = fila = 0
    for row in data[3:]:
        if texto(valor_col(row, COL_P)).lower() != "peterson":
            continue
        total += 1
        if texto(valor_col(row, COL_U)).lower() == "migrado":
            migrados += 1
        if elegivel(row):
            fila += 1
    return {"totalPeterson": total, "filaSegura": fila, "migrados": migrados}


# =========================
# Estado
# =========================


def carregar_estado() -> dict[str, Any]:
    if not STATE_FILE.exists():
        return {"confirmados": {}, "falhas": {}}
    try:
        data = json.loads(STATE_FILE.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            raise ValueError
        data.setdefault("confirmados", {})
        data.setdefault("falhas", {})
        return data
    except Exception:
        return {"confirmados": {}, "falhas": {}}


def salvar_estado(data: dict[str, Any]) -> None:
    tmp = STATE_FILE.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(STATE_FILE)


def msg_hash(msg: str) -> str:
    return hashlib.sha256(msg.encode("utf-8")).hexdigest()[:20]


# =========================
# WhatsApp UI Automation
# =========================


def importar_whatsapp():
    try:
        from pywinauto import Desktop
    except ImportError as exc:
        raise RuntimeError("pywinauto não instalado. Rode o BAT do robô.") from exc
    return Desktop


def janela_whatsapp(Desktop):
    for window in Desktop(backend="uia").windows():
        try:
            if "whatsapp" in window.window_text().lower():
                return window
        except Exception:
            pass
    raise RuntimeError("WhatsApp Business Desktop não encontrado.")


def info(control):
    try:
        el = control.element_info
        name = texto(getattr(el, "name", ""))
        aid = texto(getattr(el, "automation_id", ""))
        ctype = texto(getattr(el, "control_type", ""))
    except Exception:
        name = aid = ctype = ""
    try:
        wt = texto(control.window_text())
    except Exception:
        wt = ""
    try:
        r = control.rectangle()
        rect = (r.left, r.top, r.right, r.bottom)
    except Exception:
        rect = (0, 0, 0, 0)
    return name, aid, wt, ctype, rect


def visiveis(window):
    try:
        controls = window.descendants()
    except Exception:
        return []
    result = []
    for c in controls:
        try:
            if c.is_visible():
                result.append((c, *info(c)))
        except Exception:
            pass
    return result


def confirmar_conversa(window, profissional, telefone) -> bool:
    ultimos = telefone[-8:]
    nome = profissional.lower()
    for c, cname, aid, wt, ctype, rect in visiveis(window):
        alvo = f"{cname} {aid} {wt}".lower()
        if ultimos and ultimos in re.sub(r"\D", "", alvo):
            return True
        if nome and nome in alvo:
            return True
    return False


def abrir_chat(telefone):
    subprocess.Popen(
        ["cmd", "/c", "start", "", f"whatsapp://send?phone={telefone}"],
        shell=False,
    )


def campo_mensagem(window):
    wr = window.rectangle()
    wl, wt, wrt, wb = wr.left, wr.top, wr.right, wr.bottom
    width_total = max(1, wrt - wl)
    height_total = max(1, wb - wt)
    candidatos = []

    try:
        edits = window.descendants(control_type="Edit")
    except Exception:
        edits = []

    for c in edits:
        try:
            if not c.is_visible():
                continue
            r = c.rectangle()
            left, top, right, bottom = r.left, r.top, r.right, r.bottom
            cx = (left + right) / 2
            cy = (top + bottom) / 2
            width = right - left
            score = 0
            if cy > wt + height_total * 0.63:
                score += 10
            if cx > wl + width_total * 0.40:
                score += 8
            if width > width_total * 0.28:
                score += 6
            name, aid, value, ctype, _ = info(c)
            alvo = f"{name} {aid} {value}".lower()
            if any(x in alvo for x in ("mensagem", "message", "digite", "type", "escreva")):
                score += 12
            candidatos.append((score, bottom, right, c))
        except Exception:
            pass

    if not candidatos:
        return None
    candidatos.sort(key=lambda x: (x[0], x[1], x[2]))
    return candidatos[-1][3]


def botao_por_texto(window, termos):
    termos = tuple(t.lower() for t in termos)
    for c, name, aid, wt, ctype, rect in visiveis(window):
        if ctype.lower() not in ("button", "menuitem", "listitem", "checkbox", "text"):
            continue
        alvo = f"{name} {aid} {wt}".lower()
        if any(t in alvo for t in termos):
            return c
    return None


def botao_enviar(window):
    c = botao_por_texto(window, ("enviar", "send", "send message"))
    if c:
        return c

    # Fallback geométrico: botão no canto inferior direito da janela.
    wr = window.rectangle()
    candidatos = []
    for c, name, aid, wt, ctype, rect in visiveis(window):
        if ctype.lower() != "button":
            continue
        left, top, right, bottom = rect
        if right >= wr.right - 30 and bottom >= wr.bottom - 120:
            candidatos.append((bottom, right, c))
    if candidatos:
        candidatos.sort()
        return candidatos[-1][2]
    return None


def snapshot_textos_conversa(window):
    wr = window.rectangle()
    limiar = wr.left + 230
    out = []
    for c, name, aid, wt, ctype, rect in visiveis(window):
        if rect[0] > limiar and wt:
            out.append(wt)
    return "\n".join(out)


def texto_enviado_apareceu(window, mensagem) -> bool:
    primeira = texto(mensagem.splitlines()[0]).lower()
    if not primeira:
        return False
    snap = snapshot_textos_conversa(window).lower()
    return primeira[:30] in snap


def enviar_confirmado(window, mensagem):
    # A mensagem é inserida diretamente no controle. Não usa mouse.
    campo = campo_mensagem(window)
    if campo is None:
        raise RuntimeError("Caixa de mensagem não encontrada.")

    try:
        campo.set_edit_text(mensagem)
    except Exception as exc:
        raise RuntimeError(
            f"Não consegui inserir a mensagem diretamente no campo: {exc}"
        ) from exc

    time.sleep(ESPERA_COLAR)

    botao = botao_enviar(window)
    if botao is not None:
        acionado = False
        try:
            botao.invoke()
            acionado = True
        except Exception:
            pass
        if not acionado:
            try:
                botao.set_focus()
                botao.type_keys("{ENTER}", set_foreground=False)
                acionado = True
            except Exception:
                pass
        if not acionado:
            try:
                campo.set_focus()
                campo.type_keys("{ENTER}", set_foreground=False)
                acionado = True
            except Exception as exc:
                raise RuntimeError(f"Não consegui acionar Enviar: {exc}") from exc
    else:
        try:
            campo.set_focus()
            campo.type_keys("{ENTER}", set_foreground=False)
        except Exception as exc:
            raise RuntimeError(f"Botão Enviar não encontrado e Enter falhou: {exc}") from exc

    # CONFIRMAÇÃO FORTE: NÃO aceitamos apenas "campo vazio".
    limite = time.time() + TIMEOUT_ENVIO
    while time.time() < limite:
        time.sleep(0.45)
        if texto_enviado_apareceu(window, mensagem):
            return

    raise RuntimeError(
        "A mensagem não apareceu na conversa após o envio. "
        "O Google Sheets NÃO será atualizado."
    )


def abrir_e_confirmar(window, telefone, profissional):
    abrir_chat(telefone)
    time.sleep(ESPERA_DEEP_LINK)

    limite = time.time() + TIMEOUT_ABRIR
    while time.time() < limite:
        if confirmar_conversa(window, profissional, telefone):
            return
        time.sleep(0.3)

    raise RuntimeError(
        f"Não consegui confirmar a conversa de {profissional}."
    )


# =========================
# Etiqueta "Peterson Migração"
# =========================


def aplicar_etiqueta(window) -> tuple[bool, str]:
    # Usa UI Automation; não move mouse.
    menu = botao_por_texto(
        window,
        (
            "mais opções",
            "mais opcoes",
            "more options",
        ),
    )

    if menu is None:
        # Procura o botão de menu pela posição relativa no topo direito.
        wr = window.rectangle()
        candidatos = []
        for c, name, aid, wt, ctype, rect in visiveis(window):
            if ctype.lower() != "button":
                continue
            left, top, right, bottom = rect
            if right > wr.right - 80 and top < wr.top + 120:
                candidatos.append((right, bottom, c))
        if candidatos:
            candidatos.sort()
            menu = candidatos[-1][2]

    if menu is None:
        return False, "Menu de conversa não encontrado."

    try:
        menu.invoke()
    except Exception:
        try:
            menu.set_focus()
            menu.type_keys("{ENTER}", set_foreground=False)
        except Exception:
            return False, "Não consegui abrir o menu."

    time.sleep(ESPERA_MENU)

    # Na interface mostrada pelo usuário, a ação é "Adicionar à lista".
    item = botao_por_texto(
        window,
        (
            "adicionar à lista",
            "adicionar a lista",
            "add to list",
        ),
    )

    if item is None:
        return False, "Opção 'Adicionar à lista' não encontrada."

    try:
        item.invoke()
    except Exception:
        try:
            item.set_focus()
            item.type_keys("{ENTER}", set_foreground=False)
        except Exception:
            return False, "Não consegui abrir 'Adicionar à lista'."

    time.sleep(ESPERA_MENU)

    label = botao_por_texto(
        window,
        (NOME_ETIQUETA,),
    )

    if label is None:
        return False, f"Etiqueta '{NOME_ETIQUETA}' não encontrada."

    try:
        label.invoke()
    except Exception:
        try:
            label.set_focus()
            label.type_keys("{ENTER}", set_foreground=False)
        except Exception:
            return False, f"Não consegui selecionar '{NOME_ETIQUETA}'."

    time.sleep(ESPERA_MENU)

    return True, "Etiqueta aplicada."


def contatos_em_andamento(api):
    data = ler_planilha(api)
    result = []
    for idx in range(3, len(data)):
        row = data[idx]
        if texto(valor_col(row, COL_P)).lower() != "peterson":
            continue
        if texto(valor_col(row, COL_O)).lower() != "em andamento":
            continue
        prof = texto(valor_col(row, COL_M))
        fone = normalizar_telefone(valor_col(row, COL_N))
        if prof and fone:
            result.append({"linha": idx + 1, "profissional": prof, "telefone": fone})
    return result


def etiquetar(api, quantidade):
    Desktop = importar_whatsapp()
    window = janela_whatsapp(Desktop)
    itens = contatos_em_andamento(api)
    if not itens:
        print("✅ Nenhum contato em andamento.")
        return
    qtd = len(itens) if quantidade <= 0 else min(quantidade, len(itens))
    print(f"🏷️ Contatos para etiquetar: {qtd}")
    if input(f'Digite "INICIAR ETIQUETAS" para aplicar {NOME_ETIQUETA}: ').strip() != "INICIAR ETIQUETAS":
        print("Cancelado.")
        return
    ok = falhas = 0
    for i, item in enumerate(itens[:qtd], 1):
        try:
            abrir_e_confirmar(window, item["telefone"], item["profissional"])
            sucesso, motivo = aplicar_etiqueta(window)
            if sucesso:
                ok += 1
                log(f"[ETQ {i}/{qtd}] ✅ {item['profissional']}")
            else:
                falhas += 1
                log(f"[ETQ {i}/{qtd}] ⚠️ {item['profissional']} | {motivo}")
        except Exception as exc:
            falhas += 1
            log(f"[ETQ {i}/{qtd}] ❌ {item['profissional']} | {type(exc).__name__}: {exc}")
        if i < qtd:
            time.sleep(INTERVALO_ENTRE_CONTATOS)
    print(f"✅ Etiquetas aplicadas: {ok}/{qtd}")
    print(f"⚠️ Falhas: {falhas}")


# =========================
# Envio
# =========================


def enviar_lote(api, quantidade):
    Desktop = importar_whatsapp()
    window = janela_whatsapp(Desktop)
    data = ler_planilha(api)
    fila = montar_fila(data)
    if not fila:
        print("✅ Fila vazia.")
        return

    qtd = len(fila) if quantidade <= 0 else min(quantidade, len(fila))
    print("\n=" * 2 + "=" * 78)
    print("FILA — DEFINITIVO")
    print("=" * 80)
    for i, item in enumerate(fila[:qtd], 1):
        print(f"{i}. linha {item['linha']} | {item['profissional']}")

    if input(f'\nDigite "INICIAR PETERSON" para enviar para {qtd}: ').strip() != "INICIAR PETERSON":
        print("Cancelado.")
        return

    estado = carregar_estado()
    confirmados = estado["confirmados"]
    falhas = estado["falhas"]

    linhas_confirmadas = []
    enviados = etiquetas_ok = etiquetas_falha = ignorados = 0

    for pos, item in enumerate(fila[:qtd], 1):
        chave = str(item["linha"])
        msg = mensagem_para(item)
        h = msg_hash(msg)

        if isinstance(confirmados.get(chave), dict) and confirmados[chave].get("hash") == h:
            ignorados += 1
            log(f"[{pos}/{qtd}] ↷ Já confirmado | linha={item['linha']} | {item['profissional']}")
            continue

        try:
            log(f"[{pos}/{qtd}] Abrindo/enviando | {item['profissional']}")
            abrir_e_confirmar(window, item["telefone"], item["profissional"])
            enviar_confirmado(window, msg)

            # Estado local primeiro, para proteção contra duplicação.
            confirmados[chave] = {
                "telefone": item["telefone"],
                "profissional": item["profissional"],
                "hash": h,
                "status": "confirmado",
                "em": agora(),
            }
            falhas.pop(chave, None)
            salvar_estado(estado)

            linhas_confirmadas.append(item["linha"])
            enviados += 1
            log(f"[{pos}/{qtd}] ✅ ENVIO CONFIRMADO | linha={item['linha']}")

            # Etiqueta: se falhar, o envio continua válido e O será atualizado.
            sucesso_etq, motivo = aplicar_etiqueta(window)
            if sucesso_etq:
                etiquetas_ok += 1
                log(f"[{pos}/{qtd}] 🏷️ Etiqueta OK | {item['profissional']}")
            else:
                etiquetas_falha += 1
                log(f"[{pos}/{qtd}] ⚠️ Mensagem OK; etiqueta pendente | {item['profissional']} | {motivo}")

        except Exception as exc:
            falhas[chave] = {
                "telefone": item["telefone"],
                "profissional": item["profissional"],
                "hash": h,
                "status": "nao_confirmado",
                "erro": f"{type(exc).__name__}: {exc}",
                "em": agora(),
            }
            salvar_estado(estado)
            log(f"[{pos}/{qtd}] ❌ NÃO CONFIRMADO | {item['profissional']} | {type(exc).__name__}: {exc}")

        if pos < qtd:
            time.sleep(INTERVALO_ENTRE_CONTATOS)

    # Uma chamada ao Google, no final.
    atualizar_o_lote(api, linhas_confirmadas)

    print("\n" + "=" * 80)
    print(f"✅ ENVIOS CONFIRMADOS: {enviados}/{qtd}")
    print(f"🏷️ ETIQUETAS OK: {etiquetas_ok}/{enviados}")
    print(f"⚠️ ETIQUETAS PENDENTES: {etiquetas_falha}")
    print(f"↷ JÁ CONFIRMADOS: {ignorados}")
    print("✅ O atualizado somente para 'Em andamento' após confirmação.")
    print("=" * 80)


def resetar_linha(api, linha: int) -> None:
    # Uso pontual para desfazer um falso positivo de envio.
    api.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{SHEET_NAME}'!O{int(linha)}",
        valueInputOption="USER_ENTERED",
        body={"values": [["Não Iniciado"]]},
    ).execute()

    estado = carregar_estado()
    estado.get("confirmados", {}).pop(str(linha), None)
    estado.get("falhas", {}).pop(str(linha), None)
    salvar_estado(estado)


# =========================
# CLI
# =========================


def main():
    parser = argparse.ArgumentParser()
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--status", action="store_true")
    group.add_argument("--corrigir-o", action="store_true")
    group.add_argument("--enviar")
    group.add_argument("--etiquetar")
    group.add_argument("--resetar")
    args = parser.parse_args()

    api = google_api()

    if args.status:
        st = status(api)
        print("=" * 60)
        print("STATUS PETERSON")
        print("=" * 60)
        print(f"Total Peterson : {st['totalPeterson']}")
        print(f"Fila segura    : {st['filaSegura']}")
        print(f"Migrados       : {st['migrados']}")
        print("=" * 60)
        return

    if args.corrigir_o:
        n = corrigir_o_antigo(api)
        print(f"✅ Registros corrigidos: {n}")
        print("✅ Valor final: Em andamento")
        return

    qtd = 0 if args.enviar and args.enviar.lower() == "todos" else int(args.enviar) if args.enviar else 0
    if args.enviar:
        enviar_lote(api, qtd)
        return

    qtd = 0 if args.etiquetar and args.etiquetar.lower() == "todos" else int(args.etiquetar) if args.etiquetar else 0
    if args.etiquetar:
        etiquetar(api, qtd)
        return

    if args.resetar:
        linha = int(args.resetar)
        resetar_linha(api, linha)
        print(f"✅ Linha {linha} revertida para 'Não Iniciado' e histórico local removido.")


if __name__ == "__main__":
    main()
