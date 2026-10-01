
from __future__ import annotations

"""
ROBÔ MIGRAÇÃO PETERSON — DEFINITIVO / MODO TURBO

Fonte:
    Google Sheets -> aba "Aprovados PHC"

Fila:
    P = Peterson
    U = vazio ou "-"
    O sem andamento anterior
    ignora profissional/telefone inválidos

Fluxo:
    1) lê a fila uma única vez
    2) abre a conversa pelo deep-link do WhatsApp
    3) confirma a conversa com uma única leitura rápida da UI
    4) cola a mensagem
    5) Enter
    6) confirma o envio
    7) salva o estado local imediatamente
    8) tenta aplicar "Peterson Migração"
    9) no fim atualiza O em lote: "Em andamento"

Comandos:
    --status
    --corrigir-o
    --enviar 1
    --enviar 5
    --enviar todos
    --etiquetar 5
    --etiquetar todos

A credencial Google deve estar em:
    credentials.json
    token.json
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

COL_A_PACIENTE = 1
COL_M_PROFISSIONAL = 13
COL_N_TELEFONE = 14
COL_O_CONTATO = 15
COL_P_RESPONSAVEL = 16
COL_R_VALOR = 18
COL_S_ESCALA = 19
COL_T_CONTA = 20
COL_U_STATUS = 21

# TURBO: espera mínima para o WhatsApp processar a troca de conversa.
INTERVALO_ENTRE_ENVIO = 2.0
ESPERA_DEEP_LINK = 1.4
TIMEOUT_FALLBACK_CHAT = 5.0
ESPERA_COLAR = 0.35
ESPERA_ENVIO = 0.9
ESPERA_MENU = 0.25
ESPERA_ETIQUETA = 0.35

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


# ============================================================
# GERAL
# ============================================================

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

    if not d.startswith("55") or len(d) not in (12, 13):
        return ""

    return d


def valor_col(row: list[Any], col: int) -> Any:
    i = col - 1
    return row[i] if i < len(row) else ""


def hash_mensagem(mensagem: str) -> str:
    return hashlib.sha256(
        mensagem.encode("utf-8")
    ).hexdigest()[:20]


# ============================================================
# FILA / SHEETS
# ============================================================

def paciente_bloco(data: list[list[Any]], idx: int) -> str:
    for i in range(idx, 2, -1):
        paciente = texto(
            valor_col(data[i], COL_A_PACIENTE)
        )
        if paciente:
            return paciente
    return ""


def o_sem_andamento(valor: Any) -> bool:
    o = texto(valor).lower()

    return not any(
        termo in o
        for termo in (
            "contato",
            "em andamento",
            "aguardando",
            "resposta",
            "migrado",
            "revisão",
            "revisao",
        )
    )


def linha_elegivel(row: list[Any]) -> bool:
    responsavel = texto(
        valor_col(row, COL_P_RESPONSAVEL)
    ).lower()

    status_u = texto(
        valor_col(row, COL_U_STATUS)
    ).lower()

    profissional = texto(
        valor_col(row, COL_M_PROFISSIONAL)
    )

    fone = normalizar_telefone(
        valor_col(row, COL_N_TELEFONE)
    )

    if responsavel != "peterson":
        return False

    # U vazio ou "-" somente.
    if status_u not in ("", "-"):
        return False

    if not o_sem_andamento(
        valor_col(row, COL_O_CONTATO)
    ):
        return False

    if not profissional or not fone:
        return False

    return True


def montar_fila(data: list[list[Any]]) -> list[dict[str, Any]]:
    fila = []

    for idx in range(3, len(data)):
        row = data[idx]

        if not linha_elegivel(row):
            continue

        fila.append(
            {
                "linha": idx + 1,
                "paciente": paciente_bloco(data, idx),
                "profissional": texto(
                    valor_col(row, COL_M_PROFISSIONAL)
                ),
                "telefone": normalizar_telefone(
                    valor_col(row, COL_N_TELEFONE)
                ),
                "contato": texto(
                    valor_col(row, COL_O_CONTATO)
                ),
                "valor": valor_col(row, COL_R_VALOR),
                "escala": valor_col(row, COL_S_ESCALA),
                # T só decide a mensagem; não é exibido no nome.
                "conta": texto(
                    valor_col(row, COL_T_CONTA)
                ),
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
        return MENSAGEM_CADASTRADO.format(
            paciente=paciente
        )

    return MENSAGEM_NAO_CADASTRADO.format(
        paciente=paciente
    )


def google_sheets_api():
    try:
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from googleapiclient.discovery import build
    except ImportError as exc:
        raise RuntimeError(
            "Bibliotecas Google ausentes. Rode o BAT do robô."
        ) from exc

    creds = None

    if TOKEN_FILE.exists():
        try:
            creds = Credentials.from_authorized_user_file(
                str(TOKEN_FILE),
                SCOPES,
            )
        except Exception:
            creds = None

    if creds and creds.valid:
        pass
    elif creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        if not CREDENTIALS_FILE.exists():
            raise FileNotFoundError(
                "credentials.json não encontrado na pasta."
            )

        flow = InstalledAppFlow.from_client_secrets_file(
            str(CREDENTIALS_FILE),
            SCOPES,
        )
        creds = flow.run_local_server(
            port=0,
            open_browser=True,
        )

    TOKEN_FILE.write_text(
        creds.to_json(),
        encoding="utf-8",
    )

    return build(
        "sheets",
        "v4",
        credentials=creds,
        cache_discovery=False,
    )


def ler_planilha(api) -> list[list[Any]]:
    return (
        api.spreadsheets()
        .values()
        .get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{SHEET_NAME}'!A:U",
            valueRenderOption="UNFORMATTED_VALUE",
        )
        .execute()
        .get("values", [])
    )


def atualizar_o_lote(api, linhas: list[int]) -> None:
    if not linhas:
        return

    updates = [
        {
            "range": f"'{SHEET_NAME}'!O{linha}",
            "values": [["Em andamento"]],
        }
        for linha in linhas
    ]

    api.spreadsheets().values().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={
            "valueInputOption": "USER_ENTERED",
            "data": updates,
        },
    ).execute()


def corrigir_o_antigo(api) -> int:
    data = (
        api.spreadsheets()
        .values()
        .get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{SHEET_NAME}'!O:O",
            valueRenderOption="UNFORMATTED_VALUE",
        )
        .execute()
        .get("values", [])
    )

    updates = []

    for linha, row in enumerate(data, 1):
        valor = texto(row[0] if row else "")

        if valor.lower().startswith(
            "em andamento -"
        ):
            updates.append(
                {
                    "range": f"'{SHEET_NAME}'!O{linha}",
                    "values": [["Em andamento"]],
                }
            )

    if updates:
        api.spreadsheets().values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={
                "valueInputOption": "USER_ENTERED",
                "data": updates,
            },
        ).execute()

    return len(updates)


def status_sheets(api) -> dict[str, int]:
    data = ler_planilha(api)

    total = 0
    migrados = 0
    fila = 0

    for row in data[3:]:
        if texto(
            valor_col(row, COL_P_RESPONSAVEL)
        ).lower() != "peterson":
            continue

        total += 1

        status = texto(
            valor_col(row, COL_U_STATUS)
        ).lower()

        if status == "migrado":
            migrados += 1

        if linha_elegivel(row):
            fila += 1

    return {
        "totalPeterson": total,
        "filaSegura": fila,
        "migrados": migrados,
    }


# ============================================================
# ESTADO LOCAL / ANTI-DUPLICAÇÃO
# ============================================================

def estado_carregar() -> dict[str, Any]:
    if not STATE_FILE.exists():
        return {
            "confirmados": {},
            "falhas": {},
        }

    try:
        estado = json.loads(
            STATE_FILE.read_text(
                encoding="utf-8"
            )
        )

        if not isinstance(estado, dict):
            raise ValueError

        estado.setdefault("confirmados", {})
        estado.setdefault("falhas", {})

        return estado

    except Exception:
        return {
            "confirmados": {},
            "falhas": {},
        }


def estado_salvar(estado: dict[str, Any]) -> None:
    temp = STATE_FILE.with_suffix(".tmp")

    temp.write_text(
        json.dumps(
            estado,
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )

    temp.replace(STATE_FILE)


# ============================================================
# WHATSAPP UI
# ============================================================

def importar_whatsapp():
    try:
        from pywinauto import Desktop
        import pyautogui
        import pyperclip
    except ImportError as exc:
        raise RuntimeError(
            "Bibliotecas do WhatsApp ausentes. Rode o BAT."
        ) from exc

    return Desktop, pyautogui, pyperclip


def janela_whatsapp(Desktop):
    for window in Desktop(
        backend="uia"
    ).windows():
        try:
            if "whatsapp" in (
                window.window_text().lower()
            ):
                return window
        except Exception:
            pass

    raise RuntimeError(
        "WhatsApp Business Desktop não encontrado."
    )


def info_controle(control):
    try:
        info = control.element_info
        nome = texto(
            getattr(info, "name", "")
        )
        aid = texto(
            getattr(info, "automation_id", "")
        )
    except Exception:
        nome = ""
        aid = ""

    try:
        wt = texto(
            control.window_text()
        )
    except Exception:
        wt = ""

    try:
        rect = control.rectangle()
        rect_tuple = (
            rect.left,
            rect.top,
            rect.right,
            rect.bottom,
        )
    except Exception:
        rect_tuple = (0, 0, 0, 0)

    return nome, aid, wt, rect_tuple


def todos_visiveis(window):
    try:
        controles = window.descendants()
    except Exception:
        return []

    result = []

    for control in controles:
        try:
            if not control.is_visible():
                continue

            nome, aid, wt, rect = info_controle(
                control
            )

            result.append(
                (
                    control,
                    nome,
                    aid,
                    wt,
                    rect,
                )
            )
        except Exception:
            pass

    return result


def conversa_correta_aberta(
    window,
    profissional: str,
    telefone: str,
) -> bool:
    ultimos = re.sub(
        r"\D",
        "",
        telefone[-8:],
    )

    nome = profissional.lower()

    for control, ctrl_nome, aid, wt, rect in todos_visiveis(
        window
    ):
        alvo = (
            f"{ctrl_nome} {aid} {wt}"
        ).lower()

        if ultimos and ultimos in re.sub(
            r"\D",
            "",
            alvo,
        ):
            return True

        if nome and nome in alvo:
            return True

    return False


def abrir_chat(telefone: str) -> None:
    subprocess.Popen(
        [
            "cmd",
            "/c",
            "start",
            "",
            f"whatsapp://send?phone={telefone}",
        ],
        shell=False,
    )


def abrir_e_confirmar_chat(
    window,
    telefone: str,
    profissional: str,
    pyautogui,
    pyperclip,
) -> None:
    # Deep-link é o caminho rápido.
    abrir_chat(telefone)
    time.sleep(ESPERA_DEEP_LINK)

    if conversa_correta_aberta(
        window,
        profissional,
        telefone,
    ):
        return

    # Só usa a busca interna se o deep-link não tiver funcionado.
    try:
        for control, nome, aid, wt, rect in todos_visiveis(
            window
        ):
            try:
                tipo = str(
                    control.element_info.control_type
                    or ""
                ).lower()
            except Exception:
                tipo = ""

            if tipo != "edit":
                continue

            left, top, right, bottom = rect
            wr = window.rectangle()

            if (
                left
                < wr.left
                + (wr.right - wr.left) * 0.45
                and top
                < wr.top + 180
                and (right - left) > 150
            ):
                campo = control
                campo.click_input()
                campo.set_focus()

                pyperclip.copy(telefone)
                pyautogui.hotkey(
                    "ctrl",
                    "a",
                )
                pyautogui.hotkey(
                    "ctrl",
                    "v",
                )
                time.sleep(1.2)

                alvo = telefone[-8:]

                for c2, n2, a2, t2, r2 in todos_visiveis(
                    window
                ):
                    if alvo in re.sub(
                        r"\D",
                        "",
                        f"{n2} {t2}",
                    ):
                        try:
                            c2.click_input()
                            time.sleep(0.8)
                        except Exception:
                            pass

                        if conversa_correta_aberta(
                            window,
                            profissional,
                            telefone,
                        ):
                            return

                break

        raise RuntimeError
    except Exception as exc:
        raise RuntimeError(
            f"Não consegui abrir/confirmar a conversa de "
            f"{profissional}."
        ) from exc


def localizar_campo_mensagem(window):
    try:
        edits = window.descendants(
            control_type="Edit"
        )
    except Exception:
        edits = []

    wr = window.rectangle()

    win_left = wr.left
    win_top = wr.top
    win_right = wr.right
    win_bottom = wr.bottom

    largura_total = max(
        1,
        win_right - win_left,
    )
    altura_total = max(
        1,
        win_bottom - win_top,
    )

    candidatos = []

    for control in edits:
        try:
            if not control.is_visible():
                continue

            rect = control.rectangle()

            left = rect.left
            top = rect.top
            right = rect.right
            bottom = rect.bottom

            width = right - left
            height = bottom - top

            cx = (left + right) / 2
            cy = (top + bottom) / 2

            score = 0

            if cy > win_top + (
                altura_total * 0.65
            ):
                score += 8

            if cx > win_left + (
                largura_total * 0.40
            ):
                score += 7

            if width > largura_total * 0.30:
                score += 5

            try:
                info = control.element_info
                alvo = (
                    f"{getattr(info, 'name', '')} "
                    f"{getattr(info, 'automation_id', '')} "
                    f"{control.window_text()}"
                ).lower()
            except Exception:
                alvo = ""

            if any(
                x in alvo
                for x in (
                    "mensagem",
                    "message",
                    "digite",
                    "type",
                    "escreva",
                )
            ):
                score += 10

            candidatos.append(
                (
                    score,
                    bottom,
                    right,
                    control,
                )
            )

        except Exception:
            pass

    if not candidatos:
        return None

    candidatos.sort(
        key=lambda x: (
            x[0],
            x[1],
            x[2],
        )
    )

    return candidatos[-1][3]


def ler_campo_status(campo):
    try:
        return True, str(
            campo.get_value()
        )
    except Exception:
        pass

    try:
        return True, str(
            campo.window_text()
        )
    except Exception:
        pass

    return False, ""


def mensagem_aparece_na_conversa(
    window,
    mensagem: str,
) -> bool:
    marcador = texto(
        mensagem.splitlines()[0]
    )[:28].lower()

    if not marcador:
        return False

    for control, nome, aid, wt, rect in todos_visiveis(
        window
    ):
        try:
            tipo = str(
                control.element_info.control_type
                or ""
            ).lower()
        except Exception:
            tipo = ""

        if tipo == "edit":
            continue

        alvo = f"{nome} {aid} {wt}".lower()

        if marcador in alvo:
            if rect[0] > (
                window.rectangle().left
                + 180
            ):
                return True

    return False


def enviar_mensagem(
    window,
    telefone: str,
    profissional: str,
    mensagem: str,
    pyautogui,
    pyperclip,
) -> None:
    abrir_e_confirmar_chat(
        window,
        telefone,
        profissional,
        pyautogui,
        pyperclip,
    )

    campo = localizar_campo_mensagem(
        window
    )

    if campo is None:
        raise RuntimeError(
            "Caixa de mensagem não encontrada."
        )

    try:
        window.set_focus()
    except Exception:
        pass

    try:
        campo.set_focus()
    except Exception:
        pass

    pyperclip.copy(mensagem)

    # Sem click_input: evita mover o mouse e economiza tempo.
    pyautogui.hotkey(
        "ctrl",
        "a",
    )
    pyautogui.hotkey(
        "ctrl",
        "v",
    )

    time.sleep(
        ESPERA_COLAR
    )

    # O Enter é o método que já comprovadamente enviou
    # a mensagem no WhatsApp deste computador.
    pyautogui.press("enter")

    time.sleep(
        ESPERA_ENVIO
    )

    # Confirmação rápida: campo vazio após Enter.
    legivel, valor = ler_campo_status(
        campo
    )

    if legivel and valor.strip() == "":
        return

    # Fallback de confirmação: uma leitura da conversa.
    if mensagem_aparece_na_conversa(
        window,
        mensagem,
    ):
        return

    # Se a UI ficou lenta, uma segunda leitura curta.
    time.sleep(0.7)

    legivel, valor = ler_campo_status(
        campo
    )

    if legivel and valor.strip() == "":
        return

    if mensagem_aparece_na_conversa(
        window,
        mensagem,
    ):
        return

    raise RuntimeError(
        "Não foi possível confirmar o envio. "
        "O não será alterado."
    )


# ============================================================
# ETIQUETA
# ============================================================

def clicar_controle(control) -> bool:
    try:
        control.invoke()
        return True
    except Exception:
        pass

    try:
        control.click_input()
        return True
    except Exception:
        return False


def localizar_menu_conversa(window):
    botoes = []

    try:
        controls = window.descendants(
            control_type="Button"
        )
    except Exception:
        controls = []

    wr = window.rectangle()
    w_left = wr.left
    w_top = wr.top
    w_right = wr.right
    w_bottom = wr.bottom
    largura = max(
        1,
        w_right - w_left,
    )

    for control in controls:
        try:
            if not control.is_visible():
                continue

            nome, aid, wt, rect = info_controle(
                control
            )

            left, top, right, bottom = rect
            cx = (left + right) / 2
            cy = (top + bottom) / 2

            alvo = (
                f"{nome} {aid} {wt}"
            ).lower()

            score = 0

            if any(
                x in alvo
                for x in (
                    "mais opções",
                    "mais opcoes",
                    "more options",
                    "menu",
                )
            ):
                score += 30

            if cx > w_left + largura * 0.86:
                score += 12

            if cy < w_top + 130:
                score += 8

            if right >= w_right - 25:
                score += 5

            if score >= 18:
                botoes.append(
                    (
                        score,
                        bottom,
                        right,
                        control,
                    )
                )

        except Exception:
            pass

    if not botoes:
        return None

    botoes.sort(
        key=lambda x: (
            x[0],
            x[1],
            x[2],
        )
    )

    return botoes[-1][3]


def encontrar_controle_por_texto(
    window,
    termos,
    tipos=None,
):
    termos = tuple(
        t.lower()
        for t in termos
    )

    for control, nome, aid, wt, rect in todos_visiveis(
        window
    ):
        try:
            tipo = str(
                control.element_info.control_type
                or ""
            ).lower()
        except Exception:
            tipo = ""

        if tipos and tipo not in tipos:
            continue

        alvo = f"{nome} {aid} {wt}".lower()

        if any(
            termo in alvo
            for termo in termos
        ):
            return control

    return None


def aplicar_etiqueta(
    window,
    nome_etiqueta=NOME_ETIQUETA,
):
    # A etiqueta não interfere na confirmação do envio.
    # Se ela falhar, o contato continua marcado como enviado.
    menu = localizar_menu_conversa(
        window
    )

    if menu is None:
        return False, "Menu não encontrado."

    if not clicar_controle(menu):
        return False, "Menu não abriu."

    time.sleep(
        ESPERA_MENU
    )

    item = encontrar_controle_por_texto(
        window,
        (
            "etiquetar conversa",
            "etiquetar",
            "etiquetas",
            "labels",
            "label",
        ),
        tipos={
            "menuitem",
            "button",
            "text",
            "listitem",
        },
    )

    if item is None:
        return False, "Opção de etiqueta não encontrada."

    if not clicar_controle(item):
        return False, "Painel de etiquetas não abriu."

    time.sleep(
        ESPERA_MENU
    )

    etiqueta = encontrar_controle_por_texto(
        window,
        (nome_etiqueta,),
        tipos={
            "menuitem",
            "button",
            "text",
            "listitem",
            "checkbox",
        },
    )

    if etiqueta is None:
        return False, (
            f"Etiqueta '{nome_etiqueta}' não encontrada."
        )

    if not clicar_controle(etiqueta):
        return False, "Não consegui clicar na etiqueta."

    time.sleep(
        ESPERA_ETIQUETA
    )

    return True, "Etiqueta aplicada."


def contatos_para_etiquetar(api):
    data = ler_planilha(api)
    result = []

    for idx in range(3, len(data)):
        row = data[idx]

        if texto(
            valor_col(row, COL_P_RESPONSAVEL)
        ).lower() != "peterson":
            continue

        if texto(
            valor_col(row, COL_O_CONTATO)
        ).lower() != "em andamento":
            continue

        profissional = texto(
            valor_col(row, COL_M_PROFISSIONAL)
        )

        fone = normalizar_telefone(
            valor_col(row, COL_N_TELEFONE)
        )

        if profissional and fone:
            result.append(
                {
                    "linha": idx + 1,
                    "profissional": profissional,
                    "telefone": fone,
                }
            )

    return result


# ============================================================
# ENVIO EM LOTE
# ============================================================

def confirmar_lote(qtd):
    resposta = input(
        f'\nDigite "INICIAR PETERSON" para '
        f'enviar para {qtd}: '
    ).strip()

    return resposta == "INICIAR PETERSON"


def enviar_lote(
    api,
    quantidade: int,
) -> None:
    Desktop, pyautogui, pyperclip = (
        importar_whatsapp()
    )

    window = janela_whatsapp(
        Desktop
    )

    data = ler_planilha(api)
    fila = montar_fila(data)

    if not fila:
        print("✅ Fila vazia.")
        return

    qtd = (
        len(fila)
        if quantidade <= 0
        else min(
            quantidade,
            len(fila),
        )
    )

    print()
    print("=" * 80)
    print("FILA — MODO TURBO")
    print("=" * 80)

    for i, item in enumerate(
        fila[:qtd],
        1,
    ):
        print(
            f"{i}. linha {item['linha']} | "
            f"{item['profissional']}"
        )

    if not confirmar_lote(qtd):
        print(
            "Cancelado. Nenhuma mensagem enviada."
        )
        return

    estado = estado_carregar()
    confirmados = estado["confirmados"]
    falhas = estado["falhas"]

    linhas_para_atualizar = []
    enviados = 0
    etiqueta_ok = 0
    etiqueta_falha = 0
    ignorados = 0

    for posicao, item in enumerate(
        fila[:qtd],
        1,
    ):
        chave = str(item["linha"])
        msg = mensagem_para(item)
        hash_msg = hash_mensagem(msg)

        # Se já houve confirmação local da mesma mensagem,
        # não manda novamente.
        registro = confirmados.get(chave)

        if (
            isinstance(registro, dict)
            and registro.get("hash") == hash_msg
        ):
            ignorados += 1
            log(
                f"[{posicao}/{qtd}] ↷ "
                f"já confirmado | "
                f"linha={item['linha']} | "
                f"{item['profissional']}"
            )
            continue

        try:
            log(
                f"[{posicao}/{qtd}] "
                f"Enviando | "
                f"{item['profissional']}"
            )

            enviar_mensagem(
                window,
                item["telefone"],
                item["profissional"],
                msg,
                pyautogui,
                pyperclip,
            )

            # Salva imediatamente que o envio foi confirmado,
            # antes de qualquer etapa secundária.
            confirmados[chave] = {
                "telefone": item["telefone"],
                "profissional": item["profissional"],
                "hash": hash_msg,
                "status": "confirmado",
                "em": agora(),
            }

            falhas.pop(chave, None)
            estado_salvar(estado)

            enviados += 1
            linhas_para_atualizar.append(
                item["linha"]
            )

            log(
                f"[{posicao}/{qtd}] ✅ "
                f"ENVIO CONFIRMADO | "
                f"linha={item['linha']}"
            )

            # Etiqueta na mesma conversa, sem reabrir.
            ok, motivo = aplicar_etiqueta(
                window,
                NOME_ETIQUETA,
            )

            if ok:
                etiqueta_ok += 1
                log(
                    f"[{posicao}/{qtd}] 🏷️ "
                    f"Etiqueta aplicada | "
                    f"{item['profissional']}"
                )
            else:
                etiqueta_falha += 1
                log(
                    f"[{posicao}/{qtd}] ⚠️ "
                    f"Mensagem OK, etiqueta não aplicada | "
                    f"{item['profissional']} | {motivo}"
                )

        except Exception as exc:
            falhas[chave] = {
                "telefone": item["telefone"],
                "profissional": item["profissional"],
                "hash": hash_msg,
                "status": "nao_confirmado",
                "erro": (
                    f"{type(exc).__name__}: {exc}"
                ),
                "em": agora(),
            }

            estado_salvar(estado)

            log(
                f"[{posicao}/{qtd}] ❌ "
                f"Não confirmado | "
                f"{item['profissional']} | "
                f"{type(exc).__name__}: {exc}"
            )

        if posicao < qtd:
            time.sleep(
                INTERVALO_ENTRE_ENVIO
            )

    # Uma única chamada ao Google no fim.
    if linhas_para_atualizar:
        atualizar_o_lote(
            api,
            linhas_para_atualizar,
        )

    print()
    print("=" * 80)
    print(
        f"✅ ENVIOS CONFIRMADOS: "
        f"{enviados}/{qtd}"
    )
    print(
        f"🏷️ ETIQUETAS OK: "
        f"{etiqueta_ok}/{enviados}"
    )
    print(
        f"⚠️ ETIQUETAS PENDENTES: "
        f"{etiqueta_falha}"
    )
    print(
        f"↷ JÁ CONFIRMADOS: "
        f"{ignorados}"
    )
    print(
        "✅ O atualizado em lote como: "
        "Em andamento"
    )
    print("=" * 80)


def etiquetar_lote(
    api,
    quantidade: int,
):
    Desktop, pyautogui, pyperclip = (
        importar_whatsapp()
    )

    window = janela_whatsapp(
        Desktop
    )

    fila = contatos_para_etiquetar(
        api
    )

    if not fila:
        print(
            "✅ Nenhum contato com "
            "O='Em andamento'."
        )
        return

    qtd = (
        len(fila)
        if quantidade <= 0
        else min(
            quantidade,
            len(fila),
        )
    )

    print()
    print(
        f"🏷️ Contatos para etiquetar: {qtd}"
    )

    resposta = input(
        f'\nDigite "INICIAR ETIQUETAS" '
        f'para aplicar {NOME_ETIQUETA}: '
    ).strip()

    if resposta != "INICIAR ETIQUETAS":
        print("Cancelado.")
        return

    ok_total = 0
    falhas_total = 0

    for i, item in enumerate(
        fila[:qtd],
        1,
    ):
        try:
            abrir_e_confirmar_chat(
                window,
                item["telefone"],
                item["profissional"],
                pyautogui,
                pyperclip,
            )

            ok, motivo = aplicar_etiqueta(
                window,
                NOME_ETIQUETA,
            )

            if ok:
                ok_total += 1
                log(
                    f"[ETQ {i}/{qtd}] ✅ "
                    f"{item['profissional']}"
                )
            else:
                falhas_total += 1
                log(
                    f"[ETQ {i}/{qtd}] ⚠️ "
                    f"{item['profissional']} | "
                    f"{motivo}"
                )

        except Exception as exc:
            falhas_total += 1
            log(
                f"[ETQ {i}/{qtd}] ❌ "
                f"{item['profissional']} | "
                f"{type(exc).__name__}: {exc}"
            )

        if i < qtd:
            time.sleep(
                INTERVALO_ENTRE_ENVIO
            )

    print()
    print("=" * 70)
    print(
        f"🏷️ ETIQUETAS APLICADAS: "
        f"{ok_total}/{qtd}"
    )
    print(
        f"❌ FALHAS: {falhas_total}"
    )
    print("=" * 70)


def main():
    parser = argparse.ArgumentParser(
        description="Robô Migração Peterson"
    )

    grupo = parser.add_mutually_exclusive_group(
        required=True
    )

    grupo.add_argument(
        "--status",
        action="store_true",
    )

    grupo.add_argument(
        "--corrigir-o",
        action="store_true",
    )

    grupo.add_argument(
        "--enviar",
        metavar="QTD",
    )

    grupo.add_argument(
        "--etiquetar",
        metavar="QTD",
    )

    args = parser.parse_args()

    api = google_sheets_api()

    if args.status:
        st = status_sheets(api)

        print()
        print("=" * 60)
        print("STATUS PETERSON")
        print("=" * 60)
        print(
            f"Total Peterson : "
            f"{st['totalPeterson']}"
        )
        print(
            f"Fila segura    : "
            f"{st['filaSegura']}"
        )
        print(
            f"Migrados       : "
            f"{st['migrados']}"
        )
        print("=" * 60)
        return

    if args.corrigir_o:
        quantidade = corrigir_o_antigo(
            api
        )

        print()
        print(
            f"✅ Registros corrigidos: "
            f"{quantidade}"
        )
        print(
            "✅ Valor final: Em andamento"
        )
        return

    if args.enviar is not None:
        qtd = (
            0
            if args.enviar.lower()
            == "todos"
            else int(args.enviar)
        )

        enviar_lote(
            api,
            qtd,
        )
        return

    if args.etiquetar is not None:
        qtd = (
            0
            if args.etiquetar.lower()
            == "todos"
            else int(args.etiquetar)
        )

        etiquetar_lote(
            api,
            qtd,
        )


if __name__ == "__main__":
    main()
