import sqlite3
from calendar import monthrange
from datetime import datetime, date, time, timedelta
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
import shutil
import json

# Dependencias opcionais para exportacao.
# Instale com:
# pip install openpyxl reportlab
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    OPENPYXL_OK = True
except Exception:
    OPENPYXL_OK = False

try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    REPORTLAB_OK = True
except Exception:
    REPORTLAB_OK = False

# ============================================================
# REGISTRO DE PONTO MANUAL - CONTROLE PARALELO
# ------------------------------------------------------------
# Aplicativo local em Python + SQLite.
# Pode ser transformado em executavel com PyInstaller.
#
# Recursos:
# - Registro manual de ponto
# - Fechamento do dia 16 ao dia 15
# - Banco local SQLite
# - Exportacao para Excel
# - Exportacao para PDF
# - Feriados nacionais e estaduais de SP automaticos
# - Tela de configuracoes sem mexer no codigo
# - Backup manual do banco de dados
# ============================================================

APP_NAME = "RegistroPontoManual"
APP_DIR = Path.home() / APP_NAME
APP_DIR.mkdir(exist_ok=True)
DB_PATH = APP_DIR / "ponto_manual.db"
CONFIG_PATH = APP_DIR / "config.json"
BACKUP_DIR = APP_DIR / "backups"
BACKUP_DIR.mkdir(exist_ok=True)

CONFIG_PADRAO = {
    "entrada_padrao": "08:00",
    "saida_padrao": "18:00",
    "saida_padrao_sexta": "17:00",
    "jornada_minutos": 540,
    "jornada_sexta_minutos": 480,
    "intervalo_minimo_minutos": 60,
    "tolerancia_entrada_minutos": 5,
    "tolerancia_saida_minutos": 5,
    "limite_banco_seg_a_qui_minutos": 60,
    "limite_banco_sexta_minutos": 120,
    "fechamento_inicio_dia": 16,
    "fechamento_fim_dia": 15,
    "considerar_feriados_automaticos": True,
    "estado_feriados": "SP",
    "dias_uteis_sem_lancamento_geram_debito": True,
}


def parse_hora(valor: str):
    valor = normalizar_hora(valor)
    if not valor:
        return None
    try:
        return datetime.strptime(valor, "%H:%M").time()
    except ValueError:
        raise ValueError(f"Horario invalido: {valor}. Use HH:MM ou HHMM, exemplo: 08:00 ou 0800")


def normalizar_hora(valor: str):
    valor = (valor or "").strip().lower().replace(" ", "")
    if not valor:
        return ""

    valor = valor.replace("h", ":")
    if valor.endswith(":"):
        valor = valor[:-1]

    if ":" in valor:
        partes = valor.split(":")
        if len(partes) != 2 or not partes[0].isdigit() or not partes[1].isdigit():
            raise ValueError(f"Horario invalido: {valor}. Use HH:MM ou HHMM, exemplo: 08:00 ou 0800")
        hora = int(partes[0])
        minuto = int(partes[1])
    else:
        if not valor.isdigit() or len(valor) > 4:
            raise ValueError(f"Horario invalido: {valor}. Use HH:MM ou HHMM, exemplo: 08:00 ou 0800")
        if len(valor) <= 2:
            hora = int(valor)
            minuto = 0
        elif len(valor) == 3:
            hora = int(valor[0])
            minuto = int(valor[1:])
        else:
            hora = int(valor[:2])
            minuto = int(valor[2:])

    if hora > 23 or minuto > 59:
        raise ValueError(f"Horario invalido: {valor}. Use uma hora entre 00:00 e 23:59")
    return f"{hora:02d}:{minuto:02d}"


def jornada_minutos_para_data(d: date):
    if d.weekday() == 4:
        return int(CONFIG.get("jornada_sexta_minutos", 480))
    return int(CONFIG["jornada_minutos"])


def saida_padrao_para_data(d: date):
    if d.weekday() == 4:
        return CONFIG.get("saida_padrao_sexta", "17:00")
    return CONFIG["saida_padrao"]


def parse_data(valor: str):
    valor = (valor or "").strip()
    if not valor:
        raise ValueError("Data vazia.")
    formatos = ["%d/%m/%Y", "%Y-%m-%d"]
    for fmt in formatos:
        try:
            return datetime.strptime(valor, fmt).date()
        except ValueError:
            pass
    raise ValueError("Data invalida. Use DD/MM/AAAA.")


def data_br(d: date):
    return d.strftime("%d/%m/%Y")


def data_iso(d: date):
    return d.strftime("%Y-%m-%d")


def minutos_do_dia(t: time):
    return t.hour * 60 + t.minute


def hora_para_texto(minutos: int):
    sinal = "-" if minutos < 0 else ""
    minutos = abs(int(minutos))
    h = minutos // 60
    m = minutos % 60
    return f"{sinal}{h:02d}:{m:02d}"


def normalizar_booleano(valor):
    if isinstance(valor, bool):
        return valor
    if isinstance(valor, (int, float)):
        return bool(valor)
    if isinstance(valor, str):
        return valor.strip().lower() in {"1", "true", "t", "sim", "s", "yes", "y"}
    return bool(valor)


def validar_config(config):
    normalizada = CONFIG_PADRAO.copy()
    normalizada.update(config or {})

    normalizada["entrada_padrao"] = normalizar_hora(normalizada["entrada_padrao"])
    normalizada["saida_padrao"] = normalizar_hora(normalizada["saida_padrao"])
    normalizada["saida_padrao_sexta"] = normalizar_hora(normalizada["saida_padrao_sexta"])
    if not normalizada["entrada_padrao"]:
        raise ValueError("Configuracao invalida para entrada_padrao. Use HH:MM.")
    if not normalizada["saida_padrao"]:
        raise ValueError("Configuracao invalida para saida_padrao. Use HH:MM.")
    parse_hora(normalizada["entrada_padrao"])
    parse_hora(normalizada["saida_padrao"])
    parse_hora(normalizada["saida_padrao_sexta"])

    campos_inteiros_minimos = {
        "jornada_minutos": 1,
        "jornada_sexta_minutos": 1,
        "intervalo_minimo_minutos": 0,
        "tolerancia_entrada_minutos": 0,
        "tolerancia_saida_minutos": 0,
        "limite_banco_seg_a_qui_minutos": 0,
        "limite_banco_sexta_minutos": 0,
        "fechamento_inicio_dia": 1,
        "fechamento_fim_dia": 1,
    }

    for chave, minimo in campos_inteiros_minimos.items():
        try:
            valor = int(normalizada[chave])
        except Exception:
            raise ValueError(f"Configuracao invalida para {chave}. Use um numero inteiro.")
        if valor < minimo:
            raise ValueError(f"Configuracao invalida para {chave}. Valor minimo: {minimo}.")
        normalizada[chave] = valor

    for chave in ["fechamento_inicio_dia", "fechamento_fim_dia"]:
        if normalizada[chave] > 31:
            raise ValueError(f"Configuracao invalida para {chave}. Use um dia entre 1 e 31.")

    normalizada["estado_feriados"] = (str(normalizada.get("estado_feriados") or "SP").strip().upper() or "SP")
    normalizada["considerar_feriados_automaticos"] = normalizar_booleano(
        normalizada.get("considerar_feriados_automaticos", True)
    )
    normalizada["dias_uteis_sem_lancamento_geram_debito"] = normalizar_booleano(
        normalizada.get("dias_uteis_sem_lancamento_geram_debito", True)
    )
    return normalizada


def carregar_config():
    if not CONFIG_PATH.exists():
        salvar_config(CONFIG_PADRAO)
        return CONFIG_PADRAO.copy()

    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            dados = json.load(f)
        return validar_config(dados)
    except Exception:
        return CONFIG_PADRAO.copy()


def salvar_config(config):
    config = validar_config(config)
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=4, ensure_ascii=False)


CONFIG = carregar_config()


def conectar():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def criar_banco():
    with conectar() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS registros (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                data TEXT NOT NULL UNIQUE,
                entrada TEXT,
                saida_almoco TEXT,
                volta_almoco TEXT,
                saida TEXT,
                feriado INTEGER DEFAULT 0,
                observacao TEXT,
                criado_em TEXT DEFAULT CURRENT_TIMESTAMP,
                atualizado_em TEXT DEFAULT CURRENT_TIMESTAMP
            )
            """
        )
        conn.commit()


def dia_semana_nome(d: date):
    nomes = ["Segunda", "Terca", "Quarta", "Quinta", "Sexta", "Sabado", "Domingo"]
    return nomes[d.weekday()]


def eh_dia_util(d: date):
    return d.weekday() < 5


def pascoa(ano: int):
    """Calcula a data da Pascoa pelo algoritmo de Meeus/Jones/Butcher."""
    a = ano % 19
    b = ano // 100
    c = ano % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mes = (h + l - 7 * m + 114) // 31
    dia = ((h + l - 7 * m + 114) % 31) + 1
    return date(ano, mes, dia)


def feriados_automaticos(ano: int):
    """
    Feriados nacionais + estaduais de SP mais usados.
    Observacao: feriados municipais, pontos facultativos e regras internas
    devem ser tratados manualmente.
    """
    p = pascoa(ano)
    feriados = {
        date(ano, 1, 1): "Confraternizacao Universal",
        date(ano, 4, 21): "Tiradentes",
        date(ano, 5, 1): "Dia do Trabalho",
        date(ano, 9, 7): "Independencia do Brasil",
        date(ano, 10, 12): "Nossa Senhora Aparecida",
        date(ano, 11, 2): "Finados",
        date(ano, 11, 15): "Proclamacao da Republica",
        date(ano, 11, 20): "Consciencia Negra",
        date(ano, 12, 25): "Natal",
        p - timedelta(days=48): "Carnaval - segunda-feira",
        p - timedelta(days=47): "Carnaval - terca-feira",
        p - timedelta(days=46): "Quarta-feira de Cinzas",
        p - timedelta(days=2): "Sexta-feira Santa",
        p + timedelta(days=60): "Corpus Christi",
    }

    if CONFIG.get("estado_feriados", "SP").upper() == "SP":
        feriados[date(ano, 7, 9)] = "Revolucao Constitucionalista - SP"

    return feriados


def nome_feriado_auto(d: date):
    if not CONFIG.get("considerar_feriados_automaticos", True):
        return None
    return feriados_automaticos(d.year).get(d)


def eh_feriado_auto(d: date):
    return nome_feriado_auto(d) is not None


def calcular_registro(reg):
    d = parse_data(reg["data"]) if isinstance(reg["data"], str) else reg["data"]
    feriado_manual = int(reg.get("feriado") or 0) == 1
    feriado_auto_nome = nome_feriado_auto(d)
    feriado = feriado_manual or bool(feriado_auto_nome)

    entrada = parse_hora(reg.get("entrada"))
    saida_almoco = parse_hora(reg.get("saida_almoco"))
    volta_almoco = parse_hora(reg.get("volta_almoco"))
    saida = parse_hora(reg.get("saida"))

    avisos = []

    if feriado_auto_nome:
        avisos.append(f"Feriado automatico: {feriado_auto_nome}")
    if feriado_manual:
        avisos.append("Feriado manual marcado")

    sem_horarios = not any([entrada, saida_almoco, volta_almoco, saida])
    if sem_horarios:
        saldo = 0
        desconto = 0
        status = "Feriado" if feriado else "Sem lancamento"

        if d > date.today() and eh_dia_util(d) and not feriado:
            status = "Aguardando"
            avisos.append("Data futura ainda nao contabilizada.")
        elif (
            eh_dia_util(d)
            and not feriado
            and CONFIG.get("dias_uteis_sem_lancamento_geram_debito", True)
        ):
            saldo = -jornada_minutos_para_data(d)
            desconto = saldo
            status = "Pendente"
            avisos.append("Dia util sem lancamento contabilizado como debito.")

        return {
            "trabalhado_liquido": 0,
            "intervalo": 0,
            "saldo_total": saldo,
            "banco": 0,
            "extra": 0,
            "desconto": desconto,
            "status": status,
            "avisos": " | ".join(avisos),
            "feriado_calculado": feriado,
        }

    if not all([entrada, saida_almoco, volta_almoco, saida]):
        if d == date.today() and entrada and eh_dia_util(d) and not feriado:
            saida_almoco = saida_almoco or parse_hora("13:00")
            volta_almoco = volta_almoco or parse_hora("14:00")
            saida = saida or parse_hora(saida_padrao_para_data(d))
            avisos.append("Previsao do dia usando os horarios padrao restantes.")
        else:
            return {
                "trabalhado_liquido": 0,
                "intervalo": 0,
                "saldo_total": 0,
                "banco": 0,
                "extra": 0,
                "desconto": 0,
                "status": "Incompleto",
                "avisos": "Preencha entrada, saida almoco, volta almoco e saida.",
                "feriado_calculado": feriado,
            }

    entrada_min = minutos_do_dia(entrada)
    saida_almoco_min = minutos_do_dia(saida_almoco)
    volta_almoco_min = minutos_do_dia(volta_almoco)
    saida_min = minutos_do_dia(saida)

    if not (entrada_min < saida_almoco_min <= volta_almoco_min < saida_min):
        return {
            "trabalhado_liquido": 0,
            "intervalo": 0,
            "saldo_total": 0,
            "banco": 0,
            "extra": 0,
            "desconto": 0,
            "status": "Erro",
            "avisos": "Ordem dos horarios invalida.",
            "feriado_calculado": feriado,
        }

    intervalo = volta_almoco_min - saida_almoco_min
    manha = saida_almoco_min - entrada_min
    tarde = saida_min - volta_almoco_min
    trabalhado_liquido = manha + tarde

    saldo = trabalhado_liquido - jornada_minutos_para_data(d)

    entrada_padrao = minutos_do_dia(parse_hora(CONFIG["entrada_padrao"]))
    tol_entrada = int(CONFIG["tolerancia_entrada_minutos"])
    if entrada_padrao - tol_entrada <= entrada_min <= entrada_padrao + tol_entrada:
        diferenca_entrada = entrada_padrao - entrada_min
        if diferenca_entrada:
            saldo -= diferenca_entrada
            avisos.append(f"Tolerancia de entrada aplicada: {abs(diferenca_entrada)} min.")

    saida_padrao = minutos_do_dia(parse_hora(saida_padrao_para_data(d)))
    tol_saida = int(CONFIG["tolerancia_saida_minutos"])
    if saida_padrao - tol_saida <= saida_min <= saida_padrao + tol_saida:
        diferenca_saida = saida_min - saida_padrao
        if diferenca_saida:
            saldo -= diferenca_saida
            avisos.append(f"Tolerancia de saida aplicada: {abs(diferenca_saida)} min.")

    if intervalo < int(CONFIG["intervalo_minimo_minutos"]):
        avisos.append(
            f"Intervalo menor que 1h: {hora_para_texto(intervalo)}. "
            "A diferenca conta como tempo trabalhado."
        )

    if d.weekday() == 5:
        avisos.append("Sabado lancado. Folga fixa; tempo contabilizado no controle.")
    if d.weekday() == 6:
        avisos.append("Domingo lancado. A regra informada diz que domingo nunca trabalha.")

    # Feriado / sabado / domingo sem lancamento nao gera negativo.
    # Se trabalhar, o saldo vira o total trabalhado.
    if (not eh_dia_util(d) or feriado) and trabalhado_liquido > 0:
        saldo = trabalhado_liquido

    banco = 0
    extra = 0
    desconto = 0

    if saldo < 0:
        desconto = saldo
    elif saldo > 0:
        limite_banco = (
            int(CONFIG["limite_banco_sexta_minutos"])
            if d.weekday() == 4
            else int(CONFIG["limite_banco_seg_a_qui_minutos"])
        )
        banco = min(saldo, limite_banco)
        extra = max(0, saldo - limite_banco)

    if saldo == 0:
        status = "OK"
    elif saldo > 0 and extra == 0:
        status = "Banco positivo"
    elif saldo > 0 and extra > 0:
        status = "Banco + Extra"
    else:
        status = "Banco negativo"

    return {
        "trabalhado_liquido": trabalhado_liquido,
        "intervalo": intervalo,
        "saldo_total": saldo,
        "banco": banco,
        "extra": extra,
        "desconto": desconto,
        "status": status,
        "avisos": " | ".join(avisos),
        "feriado_calculado": feriado,
    }


def data_com_dia_no_mes(ano: int, mes: int, dia: int):
    ultimo_dia = monthrange(ano, mes)[1]
    return date(ano, mes, min(max(1, int(dia)), ultimo_dia))


def periodo_fechamento(ref: date):
    inicio_dia = int(CONFIG["fechamento_inicio_dia"])
    fim_dia = int(CONFIG["fechamento_fim_dia"])

    if ref.day >= inicio_dia:
        inicio = data_com_dia_no_mes(ref.year, ref.month, inicio_dia)
        if ref.month == 12:
            fim = data_com_dia_no_mes(ref.year + 1, 1, fim_dia)
        else:
            fim = data_com_dia_no_mes(ref.year, ref.month + 1, fim_dia)
    else:
        fim = data_com_dia_no_mes(ref.year, ref.month, fim_dia)
        if ref.month == 1:
            inicio = data_com_dia_no_mes(ref.year - 1, 12, inicio_dia)
        else:
            inicio = data_com_dia_no_mes(ref.year, ref.month - 1, inicio_dia)
    return inicio, fim


def salvar_registro(data_txt, entrada, saida_almoco, volta_almoco, saida, feriado, observacao):
    d = parse_data(data_txt)
    entrada = normalizar_hora(entrada)
    saida_almoco = normalizar_hora(saida_almoco)
    volta_almoco = normalizar_hora(volta_almoco)
    saida = normalizar_hora(saida)
    observacao = observacao or ""

    for h in [entrada, saida_almoco, volta_almoco, saida]:
        if h.strip():
            parse_hora(h)

    with conectar() as conn:
        conn.execute(
            """
            INSERT INTO registros (data, entrada, saida_almoco, volta_almoco, saida, feriado, observacao)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(data) DO UPDATE SET
                entrada=excluded.entrada,
                saida_almoco=excluded.saida_almoco,
                volta_almoco=excluded.volta_almoco,
                saida=excluded.saida,
                feriado=excluded.feriado,
                observacao=excluded.observacao,
                atualizado_em=CURRENT_TIMESTAMP
            """,
            (
                data_iso(d),
                entrada.strip(),
                saida_almoco.strip(),
                volta_almoco.strip(),
                saida.strip(),
                int(feriado),
                observacao.strip(),
            ),
        )
        conn.commit()


def excluir_registro(data_txt):
    d = parse_data(data_txt)
    with conectar() as conn:
        conn.execute("DELETE FROM registros WHERE data = ?", (data_iso(d),))
        conn.commit()


def buscar_registros(inicio: date, fim: date):
    with conectar() as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            """
            SELECT * FROM registros
            WHERE data BETWEEN ? AND ?
            ORDER BY data
            """,
            (data_iso(inicio), data_iso(fim)),
        ).fetchall()
        return [dict(r) for r in rows]


def buscar_por_data(d: date):
    with conectar() as conn:
        conn.row_factory = sqlite3.Row
        row = conn.execute("SELECT * FROM registros WHERE data = ?", (data_iso(d),)).fetchone()
        return dict(row) if row else None


def montar_linhas_fechamento(inicio: date, fim: date):
    registros = buscar_registros(inicio, fim)
    regs_por_data = {r["data"]: r for r in registros}
    linhas = []
    totais = {
        "trabalhado": 0,
        "saldo": 0,
        "banco": 0,
        "extra": 0,
        "negativo": 0,
    }

    atual = inicio
    while atual <= fim:
        iso = data_iso(atual)
        reg = regs_por_data.get(iso)

        if reg:
            calc = calcular_registro(reg)
            entrada = reg.get("entrada") or ""
            saida_almoco = reg.get("saida_almoco") or ""
            volta_almoco = reg.get("volta_almoco") or ""
            saida = reg.get("saida") or ""
            feriado = "Sim" if calc.get("feriado_calculado") else "Nao"
            obs = reg.get("observacao") or ""
        else:
            fake = {
                "data": data_iso(atual),
                "entrada": "",
                "saida_almoco": "",
                "volta_almoco": "",
                "saida": "",
                "feriado": 0,
            }
            calc = calcular_registro(fake)
            if not eh_dia_util(atual):
                calc["status"] = "Folga"
            entrada = saida_almoco = volta_almoco = saida = ""
            feriado = "Sim" if calc.get("feriado_calculado") else "Nao"
            obs = ""

        totais["trabalhado"] += calc["trabalhado_liquido"]
        totais["saldo"] += calc["saldo_total"]
        totais["banco"] += calc["banco"]
        totais["extra"] += calc["extra"]
        totais["negativo"] += calc["desconto"]

        linhas.append({
            "data": data_br(atual),
            "dia": dia_semana_nome(atual),
            "entrada": entrada,
            "saida_almoco": saida_almoco,
            "volta_almoco": volta_almoco,
            "saida": saida,
            "intervalo": hora_para_texto(calc["intervalo"]),
            "trabalhado": hora_para_texto(calc["trabalhado_liquido"]),
            "saldo": hora_para_texto(calc["saldo_total"]),
            "banco": hora_para_texto(calc["banco"]),
            "extra": hora_para_texto(calc["extra"]),
            "status": calc["status"],
            "feriado": feriado,
            "avisos": calc["avisos"],
            "observacao": obs,
        })

        atual += timedelta(days=1)

    return linhas, totais


def exportar_excel(caminho, inicio: date, fim: date):
    if not OPENPYXL_OK:
        raise RuntimeError("Biblioteca openpyxl nao instalada. Rode: pip install openpyxl")

    linhas, totais = montar_linhas_fechamento(inicio, fim)

    wb = Workbook()
    ws = wb.active
    ws.title = "Fechamento"

    ws["A1"] = "Relatorio de Ponto Manual"
    ws["A1"].font = Font(size=16, bold=True)
    ws["A2"] = f"Periodo: {data_br(inicio)} a {data_br(fim)}"
    ws["A3"] = (
        f"Saldo final: {hora_para_texto(totais['saldo'])} | "
        f"Banco: {hora_para_texto(totais['banco'])} | "
        f"Extra: {hora_para_texto(totais['extra'])} | "
        f"Negativo: {hora_para_texto(totais['negativo'])}"
    )

    cabecalho = [
        "Data", "Dia", "Entrada", "Saida Almoco", "Volta", "Saida", "Intervalo",
        "Trabalhado", "Saldo", "Banco", "Extra", "Status", "Feriado", "Avisos", "Observacao"
    ]
    ws.append([])
    ws.append(cabecalho)

    header_row = 5
    fill = PatternFill("solid", fgColor="D9EAF7")
    border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    for cell in ws[header_row]:
        cell.font = Font(bold=True)
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = border

    for l in linhas:
        ws.append([
            l["data"], l["dia"], l["entrada"], l["saida_almoco"], l["volta_almoco"], l["saida"],
            l["intervalo"], l["trabalhado"], l["saldo"], l["banco"], l["extra"], l["status"],
            l["feriado"], l["avisos"], l["observacao"]
        ])

    for row in ws.iter_rows(min_row=6):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="center")

    larguras = [12, 12, 10, 14, 10, 10, 12, 12, 10, 10, 10, 16, 10, 45, 35]
    for idx, largura in enumerate(larguras, start=1):
        ws.column_dimensions[chr(64 + idx)].width = largura

    ws.freeze_panes = "A6"
    wb.save(caminho)


def exportar_pdf(caminho, inicio: date, fim: date):
    if not REPORTLAB_OK:
        raise RuntimeError("Biblioteca reportlab nao instalada. Rode: pip install reportlab")

    linhas, totais = montar_linhas_fechamento(inicio, fim)

    doc = SimpleDocTemplate(
        str(caminho),
        pagesize=landscape(A4),
        rightMargin=20,
        leftMargin=20,
        topMargin=20,
        bottomMargin=20,
    )
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph("Relatorio de Ponto Manual", styles["Title"]))
    story.append(Paragraph(f"Periodo: {data_br(inicio)} a {data_br(fim)}", styles["Normal"]))
    story.append(Paragraph(
        f"Trabalhado: {hora_para_texto(totais['trabalhado'])} | "
        f"Saldo: {hora_para_texto(totais['saldo'])} | "
        f"Banco: {hora_para_texto(totais['banco'])} | "
        f"Extra: {hora_para_texto(totais['extra'])} | "
        f"Negativo: {hora_para_texto(totais['negativo'])}",
        styles["Normal"]
    ))
    story.append(Spacer(1, 12))

    data = [["Data", "Dia", "Ent.", "Alm.", "Volta", "Saida", "Trab.", "Saldo", "Banco", "Extra", "Status", "Feriado"]]
    for l in linhas:
        data.append([
            l["data"], l["dia"], l["entrada"], l["saida_almoco"], l["volta_almoco"], l["saida"],
            l["trabalhado"], l["saldo"], l["banco"], l["extra"], l["status"], l["feriado"]
        ])

    tabela = Table(data, repeatRows=1)
    tabela.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightblue),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))

    story.append(tabela)
    doc.build(story)


def criar_backup():
    if not DB_PATH.exists():
        raise RuntimeError("Banco de dados ainda nao existe.")
    agora = datetime.now().strftime("%Y%m%d_%H%M%S")
    destino = BACKUP_DIR / f"backup_ponto_manual_{agora}.db"
    shutil.copy2(DB_PATH, destino)
    return destino


class ConfigWindow(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.master = master
        self.title("Configuracoes")
        self.geometry("560x680")
        self.resizable(False, False)

        self.vars = {}
        self._montar()
        self.transient(master)
        self.grab_set()

    def _add_campo(self, frame, row, chave, label, largura=15):
        ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", pady=4)
        var = tk.StringVar(value=str(CONFIG.get(chave, "")))
        self.vars[chave] = var
        ttk.Entry(frame, textvariable=var, width=largura).grid(row=row, column=1, sticky="w", pady=4)

    def _montar(self):
        frame = ttk.Frame(self, padding=16)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Configuracoes do calculo", font=("Segoe UI", 15, "bold")).grid(
            row=0,
            column=0,
            columnspan=2,
            sticky="w",
            pady=(0, 12),
        )

        campos = [
            ("entrada_padrao", "Entrada padrao, ex: 08:00"),
            ("saida_padrao", "Saida padrao, ex: 18:00"),
            ("saida_padrao_sexta", "Saida padrao sexta, ex: 17:00"),
            ("jornada_minutos", "Jornada diaria em minutos"),
            ("jornada_sexta_minutos", "Jornada sexta em minutos"),
            ("intervalo_minimo_minutos", "Intervalo minimo em minutos"),
            ("tolerancia_entrada_minutos", "Tolerancia de entrada em minutos"),
            ("tolerancia_saida_minutos", "Tolerancia de saida em minutos"),
            ("limite_banco_seg_a_qui_minutos", "Banco seg. a qui. em minutos"),
            ("limite_banco_sexta_minutos", "Banco sexta em minutos"),
            ("fechamento_inicio_dia", "Dia inicial do fechamento"),
            ("fechamento_fim_dia", "Dia final do fechamento"),
            ("estado_feriados", "Estado dos feriados, ex: SP"),
        ]

        for i, (chave, label) in enumerate(campos, start=1):
            self._add_campo(frame, i, chave, label)

        self.feriados_var = tk.BooleanVar(value=bool(CONFIG.get("considerar_feriados_automaticos", True)))
        ttk.Checkbutton(
            frame,
            text="Considerar feriados automaticos",
            variable=self.feriados_var,
        ).grid(row=len(campos) + 1, column=0, columnspan=2, sticky="w", pady=(10, 2))

        self.debito_sem_lancamento_var = tk.BooleanVar(
            value=bool(CONFIG.get("dias_uteis_sem_lancamento_geram_debito", True))
        )
        ttk.Checkbutton(
            frame,
            text="Dia util sem lancamento entra como debito no fechamento",
            variable=self.debito_sem_lancamento_var,
        ).grid(row=len(campos) + 2, column=0, columnspan=2, sticky="w", pady=(2, 10))

        botoes = ttk.Frame(frame)
        botoes.grid(row=len(campos) + 3, column=0, columnspan=2, sticky="e", pady=20)
        ttk.Button(botoes, text="Salvar configuracoes", command=self.salvar).pack(side="left", padx=5)
        ttk.Button(botoes, text="Cancelar", command=self.destroy).pack(side="left", padx=5)

        aviso = (
            "Observacao: feriados municipais, pontos facultativos e regras internas especificas "
            "devem ser marcados manualmente no lancamento do dia, caso necessario."
        )
        ttk.Label(frame, text=aviso, wraplength=510, foreground="#555555").grid(
            row=len(campos) + 4,
            column=0,
            columnspan=2,
            sticky="w",
        )

    def salvar(self):
        global CONFIG
        novo = CONFIG.copy()

        try:
            for chave, var in self.vars.items():
                valor = var.get().strip()
                if chave in ["entrada_padrao", "saida_padrao", "saida_padrao_sexta"]:
                    novo[chave] = normalizar_hora(valor)
                elif chave == "estado_feriados":
                    novo[chave] = valor.upper()
                else:
                    novo[chave] = int(valor)

            novo["considerar_feriados_automaticos"] = bool(self.feriados_var.get())
            novo["dias_uteis_sem_lancamento_geram_debito"] = bool(self.debito_sem_lancamento_var.get())
            novo = validar_config(novo)
            salvar_config(novo)
            CONFIG = carregar_config()
            self.master.carregar_dia()
            self.master.atualizar_resumo()
            messagebox.showinfo("Configuracoes", "Configuracoes salvas com sucesso.")
            self.destroy()
        except Exception as e:
            messagebox.showerror("Erro", str(e))


class RegistroPontoApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Registro de Ponto Manual")
        self.geometry("1240x740")
        self.minsize(1100, 640)

        criar_banco()

        self.data_var = tk.StringVar(value=data_br(date.today()))
        self.entrada_var = tk.StringVar(value=CONFIG["entrada_padrao"])
        self.saida_almoco_var = tk.StringVar(value="13:00")
        self.volta_almoco_var = tk.StringVar(value="14:00")
        self.saida_var = tk.StringVar(value=CONFIG["saida_padrao"])
        self.feriado_var = tk.BooleanVar(value=False)
        self.obs_var = tk.StringVar()

        self.ref_var = tk.StringVar(value=data_br(date.today()))
        self.resumo_var = tk.StringVar(value="")
        self.periodo_atual = periodo_fechamento(date.today())

        self._montar_ui()
        self.carregar_dia()
        self.atualizar_resumo()

    def _montar_ui(self):
        style = ttk.Style()
        style.configure("Treeview", rowheight=28)
        style.configure("TButton", padding=6)
        style.configure("TLabel", padding=3)

        container = ttk.Frame(self, padding=12)
        container.pack(fill="both", expand=True)

        topo = ttk.Frame(container)
        topo.pack(fill="x")

        titulo = ttk.Label(topo, text="Registro de Ponto Manual", font=("Segoe UI", 18, "bold"))
        titulo.pack(side="left", anchor="w")

        ttk.Button(topo, text="Configuracoes", command=self.abrir_configuracoes).pack(side="right", padx=4)
        ttk.Button(topo, text="Backup", command=self.backup).pack(side="right", padx=4)
        ttk.Button(topo, text="Exportar PDF", command=self.exportar_pdf_ui).pack(side="right", padx=4)
        ttk.Button(topo, text="Exportar Excel", command=self.exportar_excel_ui).pack(side="right", padx=4)

        subtitulo = ttk.Label(
            container,
            text="Controle paralelo local com fechamento do dia 16 ao dia 15. Banco salvo em SQLite no computador.",
            font=("Segoe UI", 10),
        )
        subtitulo.pack(anchor="w", pady=(0, 10))

        form = ttk.LabelFrame(container, text="Lancamento do dia", padding=10)
        form.pack(fill="x")

        campos = [
            ("Data", self.data_var, 12),
            ("Entrada", self.entrada_var, 8),
            ("Saida almoco", self.saida_almoco_var, 8),
            ("Volta almoco", self.volta_almoco_var, 8),
            ("Saida", self.saida_var, 8),
        ]

        for i, (label, var, width) in enumerate(campos):
            ttk.Label(form, text=label).grid(row=0, column=i, sticky="w")
            ent = ttk.Entry(form, textvariable=var, width=width)
            ent.grid(row=1, column=i, padx=(0, 10), sticky="w")

        ttk.Checkbutton(form, text="Feriado manual", variable=self.feriado_var).grid(
            row=1,
            column=5,
            padx=(0, 10),
            sticky="w",
        )

        ttk.Label(form, text="Observacao").grid(row=0, column=6, sticky="w")
        ttk.Entry(form, textvariable=self.obs_var, width=35).grid(row=1, column=6, padx=(0, 10), sticky="we")

        botoes = ttk.Frame(form)
        botoes.grid(row=1, column=7, sticky="e")
        ttk.Button(botoes, text="Salvar", command=self.salvar).pack(side="left", padx=3)
        ttk.Button(botoes, text="Carregar", command=self.carregar_dia).pack(side="left", padx=3)
        ttk.Button(botoes, text="Excluir", command=self.excluir).pack(side="left", padx=3)
        ttk.Button(botoes, text="Limpar", command=self.limpar).pack(side="left", padx=3)

        form.columnconfigure(6, weight=1)

        fechamento = ttk.LabelFrame(container, text="Fechamento", padding=10)
        fechamento.pack(fill="x", pady=10)

        ttk.Label(fechamento, text="Data de referencia").pack(side="left")
        ttk.Entry(fechamento, textvariable=self.ref_var, width=12).pack(side="left", padx=6)
        ttk.Button(fechamento, text="Atualizar fechamento", command=self.atualizar_resumo).pack(side="left", padx=6)
        ttk.Button(fechamento, text="Usar hoje", command=self.usar_hoje).pack(side="left", padx=6)

        ttk.Label(fechamento, textvariable=self.resumo_var, font=("Segoe UI", 10, "bold")).pack(side="left", padx=20)

        tabela_frame = ttk.Frame(container)
        tabela_frame.pack(fill="both", expand=True)

        colunas = (
            "data", "dia", "entrada", "almoco", "volta", "saida", "intervalo", "trabalhado",
            "saldo", "banco", "extra", "status", "feriado", "avisos"
        )
        self.tree = ttk.Treeview(tabela_frame, columns=colunas, show="headings")

        cabecalhos = {
            "data": "Data",
            "dia": "Dia",
            "entrada": "Entrada",
            "almoco": "Saida Almoco",
            "volta": "Volta",
            "saida": "Saida",
            "intervalo": "Intervalo",
            "trabalhado": "Trabalhado",
            "saldo": "Saldo",
            "banco": "Banco",
            "extra": "Extra",
            "status": "Status",
            "feriado": "Feriado",
            "avisos": "Avisos",
        }

        larguras = {
            "data": 90,
            "dia": 80,
            "entrada": 75,
            "almoco": 95,
            "volta": 75,
            "saida": 70,
            "intervalo": 80,
            "trabalhado": 90,
            "saldo": 80,
            "banco": 80,
            "extra": 80,
            "status": 120,
            "feriado": 70,
            "avisos": 320,
        }

        for col in colunas:
            self.tree.heading(col, text=cabecalhos[col])
            self.tree.column(col, width=larguras[col], anchor="center")

        self.tree.column("avisos", anchor="w")

        yscroll = ttk.Scrollbar(tabela_frame, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(tabela_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")

        tabela_frame.rowconfigure(0, weight=1)
        tabela_frame.columnconfigure(0, weight=1)

        self.tree.bind("<Double-1>", self.duplo_clique)

        rodape = ttk.Label(
            container,
            text=f"Banco local: {DB_PATH} | Configuracoes: {CONFIG_PATH}",
            font=("Segoe UI", 8),
        )
        rodape.pack(anchor="w", pady=(6, 0))

    def preencher_formulario_padrao(self, d=None):
        d = d or date.today()
        self.data_var.set(data_br(d))
        self.entrada_var.set(CONFIG["entrada_padrao"])
        self.saida_almoco_var.set("13:00")
        self.volta_almoco_var.set("14:00")
        self.saida_var.set(saida_padrao_para_data(d))
        self.feriado_var.set(False)
        self.obs_var.set("")

    def limpar(self):
        self.preencher_formulario_padrao(date.today())

    def usar_hoje(self):
        hoje = data_br(date.today())
        self.ref_var.set(hoje)
        self.data_var.set(hoje)
        self.carregar_dia()
        self.atualizar_resumo()

    def salvar(self):
        try:
            salvar_registro(
                self.data_var.get(),
                self.entrada_var.get(),
                self.saida_almoco_var.get(),
                self.volta_almoco_var.get(),
                self.saida_var.get(),
                self.feriado_var.get(),
                self.obs_var.get(),
            )
            self.atualizar_resumo()
            messagebox.showinfo("Salvo", "Registro salvo com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro ao salvar", str(e))

    def excluir(self):
        try:
            if not messagebox.askyesno("Confirmar", "Deseja excluir o registro desta data?"):
                return
            excluir_registro(self.data_var.get())
            self.carregar_dia()
            self.atualizar_resumo()
            messagebox.showinfo("Excluido", "Registro excluido.")
        except Exception as e:
            messagebox.showerror("Erro ao excluir", str(e))

    def carregar_dia(self):
        try:
            d = parse_data(self.data_var.get())
            reg = buscar_por_data(d)
            if not reg:
                self.preencher_formulario_padrao(d)
                return
            self.data_var.set(data_br(parse_data(reg["data"])))
            self.entrada_var.set(reg.get("entrada") or "")
            self.saida_almoco_var.set(reg.get("saida_almoco") or "")
            self.volta_almoco_var.set(reg.get("volta_almoco") or "")
            self.saida_var.set(reg.get("saida") or "")
            self.feriado_var.set(bool(reg.get("feriado")))
            self.obs_var.set(reg.get("observacao") or "")
        except Exception as e:
            messagebox.showerror("Erro ao carregar", str(e))

    def duplo_clique(self, event):
        item = self.tree.focus()
        if not item:
            return
        valores = self.tree.item(item, "values")
        if not valores:
            return
        self.data_var.set(valores[0])
        self.carregar_dia()

    def abrir_configuracoes(self):
        ConfigWindow(self)

    def backup(self):
        try:
            destino = criar_backup()
            messagebox.showinfo("Backup criado", f"Backup salvo em:\n{destino}")
        except Exception as e:
            messagebox.showerror("Erro no backup", str(e))

    def exportar_excel_ui(self):
        try:
            inicio, fim = self.periodo_atual
            caminho = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel", "*.xlsx")],
                initialfile=f"fechamento_{data_iso(inicio)}_a_{data_iso(fim)}.xlsx",
            )
            if not caminho:
                return
            exportar_excel(caminho, inicio, fim)
            messagebox.showinfo("Exportado", f"Excel salvo em:\n{caminho}")
        except Exception as e:
            messagebox.showerror("Erro ao exportar Excel", str(e))

    def exportar_pdf_ui(self):
        try:
            inicio, fim = self.periodo_atual
            caminho = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF", "*.pdf")],
                initialfile=f"fechamento_{data_iso(inicio)}_a_{data_iso(fim)}.pdf",
            )
            if not caminho:
                return
            exportar_pdf(caminho, inicio, fim)
            messagebox.showinfo("Exportado", f"PDF salvo em:\n{caminho}")
        except Exception as e:
            messagebox.showerror("Erro ao exportar PDF", str(e))

    def atualizar_resumo(self):
        try:
            ref = parse_data(self.ref_var.get())
            inicio, fim = periodo_fechamento(ref)
            self.periodo_atual = (inicio, fim)

            linhas, totais = montar_linhas_fechamento(inicio, fim)

            for item in self.tree.get_children():
                self.tree.delete(item)

            for l in linhas:
                self.tree.insert(
                    "",
                    "end",
                    values=(
                        l["data"],
                        l["dia"],
                        l["entrada"],
                        l["saida_almoco"],
                        l["volta_almoco"],
                        l["saida"],
                        l["intervalo"],
                        l["trabalhado"],
                        l["saldo"],
                        l["banco"],
                        l["extra"],
                        l["status"],
                        l["feriado"],
                        l["avisos"],
                    ),
                )

            status_final = "FECHADO" if totais["saldo"] == 0 else ("SOBRANDO" if totais["saldo"] > 0 else "FALTANDO")
            self.resumo_var.set(
                f"Periodo: {data_br(inicio)} a {data_br(fim)} | "
                f"Trabalhado: {hora_para_texto(totais['trabalhado'])} | "
                f"Saldo: {hora_para_texto(totais['saldo'])} | "
                f"Banco: {hora_para_texto(totais['banco'])} | "
                f"Extra: {hora_para_texto(totais['extra'])} | "
                f"Negativo: {hora_para_texto(totais['negativo'])} | "
                f"Status: {status_final}"
            )
        except Exception as e:
            messagebox.showerror("Erro no resumo", str(e))


if __name__ == "__main__":
    app = RegistroPontoApp()
    app.mainloop()
