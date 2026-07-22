"""
logica.py — Núcleo de regras de negócio do painel de Produtividade (Ativos e MP)
================================================================================
Este módulo NÃO importa Streamlit. Ele contém só as funções puras de:
  • carregar e normalizar os dados extraídos do Salesforce
  • reconstruir a linha do tempo de cada ordem a partir de TODAS as WOLI
  • classificar cada ação em: Criada / Reagendada / Cancelada
  • atribuir cada ação à pessoa e ao dia
  • agregar por pessoa, operação, dia e tipo de serviço

Isso permite testar a lógica isoladamente (ver bloco __main__ no final).

--------------------------------------------------------------------------------
REGRA DE NEGÓCIO (validável com a diretoria)
--------------------------------------------------------------------------------
A ordem de serviço (OS) é um Case do tipo 'OS' -> tem 1 WorkOrder -> tem N WOLI.

 • CRIAR    : abrir a ordem. 1 evento por Case (OS). Atribuído a quem abriu o Case.
 • REAGENDAR: cada nova WOLI criada DEPOIS da primeira da ordem é 1 reagendamento.
              Atribuído a quem criou a nova WOLI (quem marcou o novo agendamento).
 • CANCELAR : cada WOLI que terminou com status 'Cancelado' é 1 cancelamento.
              Atribuído a quem fez a última alteração na WOLI (LastModifiedBy).

Reagendamento "de Vendas" = OS de INSTALAÇÃO. Os demais reagendamentos = "Diversos".

>>> Os campos de atribuição são CONFIGURÁVEIS logo abaixo. Se, no seu Salesforce,
    quem cria/altera os registros for um usuário de integração/automação (e não a
    pessoa), basta trocar o campo de origem aqui (ex.: usar um campo de dono/agente).
"""

from __future__ import annotations
import unicodedata
import pandas as pd

FUSO = "America/Sao_Paulo"

# ==========================================================================
# CONFIGURAÇÃO DE ATRIBUIÇÃO  (troque aqui se precisar)
# ==========================================================================
# De qual coluna sai a "pessoa" responsável por cada tipo de ação.
ATTR_CRIACAO      = "Case_CreatedBy"        # quem abriu a ordem (alternativa: WO_CreatedBy)
ATTR_REAGENDA     = "WOLI_CreatedBy"        # quem criou a nova WOLI (o novo agendamento)
ATTR_CANCELA      = "WOLI_LastModifiedBy"   # quem fez a última alteração da WOLI cancelada

# Datas correspondentes a cada ação
DATA_CRIACAO      = "Case_CreatedDate"
DATA_REAGENDA     = "WOLI_CreatedDate"
DATA_CANCELA      = "WOLI_LastModifiedDate"

# ==========================================================================
# EQUIPE  —  pessoa -> operação
# ==========================================================================
EQUIPE = {
    "LORENA CARDOSO SOUZA SILVEIRA":       "Ativo MP",
    "ALCIONE APARECIDA DE OLIVEIRA":       "Ativo MP",
    "ROSIANE SAMPAIO SILVA":               "Reagendamento Vendas + Diversos",
    "SANDY RAMOS OLIVEIRA DE JESUS":       "Reagendamento Diversos",
    "EMERSON GUSTAVO CINTRA SILVA":        "Massivo Compartilhado",
    "DEBORA RAQUEL TAVARES DE OLIVEIRA":   "Massivo Compartilhado",
    "GABRIELA MENEZES DA SILVA":           "Massivo Compartilhado",
}

# Ordem de exibição das operações
ORDEM_OPERACOES = [
    "Ativo MP",
    "Reagendamento Vendas + Diversos",
    "Reagendamento Diversos",
    "Massivo Compartilhado",
]

TIPOS_EVENTO = ["Criada", "Reagendada", "Cancelada"]

# Cancelamento no momento da criacao da ordem: a WOLI original (item unico) ja
# nasce cancelada - criada e cancelada praticamente juntas. E uma categoria a
# parte: NAO e uma criacao real nem um cancelamento pro-ativo (produtivo).
CANCELADA_CRIACAO = "Cancelada na criação"
# Janela para considerar "nasceu cancelada" (criada e cancelada no mesmo momento).
LIMITE_CANCELAMENTO_NA_CRIACAO = pd.Timedelta(minutes=5)

# Colunas que a planilha de dados deve conter (contrato com o extrator)
COLUNAS_ESPERADAS = [
    "WorkOrderId", "WorkOrderNumber", "WO_CreatedDate", "WO_CreatedBy",
    "WO_TipoServico", "WO_Status",
    "CaseId", "CaseNumber", "Case_Type", "Case_Status", "Case_TipoSolicitacao",
    "Case_CreatedBy", "Case_CreatedDate", "Cliente", "CodigoItem",
    "WOLI_Id", "LineItemNumber", "Numero_OS", "WOLI_Status",
    "WOLI_CreatedDate", "WOLI_CreatedBy", "WOLI_LastModifiedDate", "WOLI_LastModifiedBy",
    "Motivo_Insucesso", "Nome_Tecnico",
]


# ==========================================================================
# HELPERS
# ==========================================================================
def _sem_acento(texto: str) -> str:
    if texto is None:
        return ""
    txt = unicodedata.normalize("NFKD", str(texto))
    return "".join(c for c in txt if not unicodedata.combining(c))


def normaliza_nome(nome) -> str:
    """Padroniza um nome para casar com a EQUIPE (maiúsculas, sem acento, espaço único)."""
    if nome is None or (isinstance(nome, float) and pd.isna(nome)):
        return ""
    base = _sem_acento(nome).upper().strip()
    return " ".join(base.split())


# lookup normalizado -> (nome_canônico, operação)
_EQUIPE_NORM = {normaliza_nome(k): (k, v) for k, v in EQUIPE.items()}


def pessoa_para_operacao(nome) -> str | None:
    reg = _EQUIPE_NORM.get(normaliza_nome(nome))
    return reg[1] if reg else None


def nome_canonico(nome) -> str | None:
    reg = _EQUIPE_NORM.get(normaliza_nome(nome))
    return reg[0] if reg else None


def eh_instalacao(tipo_servico) -> bool:
    """OS de instalação de Vendas (base para 'Reagendamento de Vendas').
    Atenção: 'DESINSTALAÇÃO' e 'REINSTALAÇÃO' contêm 'INSTALA' mas NÃO são
    instalação de Vendas — entram em Diversos."""
    t = _sem_acento(tipo_servico).upper()
    if "DESINSTALA" in t or "REINSTALA" in t:
        return False
    return "INSTALA" in t


def eh_cancelado(status) -> bool:
    return "CANCEL" in _sem_acento(status).upper()


def _cancelada_na_criacao(r: dict) -> bool:
    """True se a WOLI foi criada e cancelada praticamente no mesmo momento
    (a ordem 'nasceu cancelada')."""
    cri = r.get("WOLI_CreatedDate")
    mod = r.get("WOLI_LastModifiedDate")
    if pd.isna(cri) or pd.isna(mod):
        return False
    return (mod - cri) <= LIMITE_CANCELAMENTO_NA_CRIACAO


def _to_dt_local(serie: pd.Series) -> pd.Series:
    """Converte timestamps do Salesforce (UTC) para o horário de São Paulo."""
    dt = pd.to_datetime(serie, errors="coerce", utc=True)
    try:
        return dt.dt.tz_convert(FUSO)
    except Exception:
        return dt


# ==========================================================================
# CARGA E NORMALIZAÇÃO
# ==========================================================================
def carregar_dados(caminho: str) -> pd.DataFrame:
    """Lê a planilha de WOLI extraída do Salesforce e normaliza tipos/datas."""
    df = pd.read_excel(caminho)
    return preparar_dataframe(df)


def preparar_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # Garante todas as colunas esperadas (as ausentes viram vazias)
    for col in COLUNAS_ESPERADAS:
        if col not in df.columns:
            df[col] = pd.NA

    # Datas -> horário local
    for col in ["WO_CreatedDate", "Case_CreatedDate",
                "WOLI_CreatedDate", "WOLI_LastModifiedDate"]:
        df[col] = _to_dt_local(df[col])

    # Flags auxiliares
    df["is_instalacao"] = df["WO_TipoServico"].apply(eh_instalacao)
    df["is_cancelada"] = df["WOLI_Status"].apply(eh_cancelado)
    df["Segmento_Servico"] = df["is_instalacao"].map({True: "Instalação (Vendas)",
                                                      False: "Diversos"})
    return df


# ==========================================================================
# CONSTRUÇÃO DA TABELA DE EVENTOS  (o coração do painel)
# ==========================================================================
def construir_eventos(df: pd.DataFrame) -> pd.DataFrame:
    """
    A partir de TODAS as WOLI, gera uma linha por AÇÃO (evento):
      Criada / Reagendada / Cancelada — com pessoa, data e contexto da ordem.
    """
    if df.empty:
        return pd.DataFrame(columns=_colunas_evento())

    eventos: list[dict] = []
    nascidas_canceladas: set = set()   # CaseId de ordens que já nascem canceladas

    # ---- REAGENDAMENTO e CANCELAMENTO: percorrendo as WOLI por ordem ----
    limite_original = pd.Timedelta(days=2)  # 1ª WOLI é criada junto com a ordem
    df_wo = df.dropna(subset=["WorkOrderId"]).sort_values(
        ["WorkOrderId", "WOLI_CreatedDate"]
    )
    for _wo, grupo in df_wo.groupby("WorkOrderId", sort=False):
        linhas = grupo.to_dict("records")
        # A 1ª WOLI é a "original" (agendamento inicial) só se foi criada junto
        # com a ordem. Se a 1ª que temos é bem posterior à criação da OS, a
        # original ficou fora da janela extraída -> então ela também é reagendamento.
        primeira_eh_original = True
        wo_criacao = linhas[0].get("WO_CreatedDate")
        woli_criacao = linhas[0].get("WOLI_CreatedDate")
        if pd.notna(wo_criacao) and pd.notna(woli_criacao):
            primeira_eh_original = (woli_criacao - wo_criacao) <= limite_original

        for i, r in enumerate(linhas):
            eh_original = (i == 0 and primeira_eh_original)
            is_canc = bool(r.get("is_cancelada"))
            # WOLI criada e cancelada no MESMO momento = "nasce cancelada".
            # Vale para a WOLI original (a própria ordem nasce cancelada) E para
            # itens posteriores (ex.: tentativa de reagendamento que já sai
            # cancelada, como um erro de integração). Nesses casos é UM único
            # evento "Cancelada na criação" — nunca reagendamento, nunca
            # cancelamento pró-ativo.
            nasce_cancelada = is_canc and _cancelada_na_criacao(r)
            if nasce_cancelada:
                pessoa = r.get("WOLI_CreatedBy")
                if normaliza_nome(pessoa) != "":
                    eventos.append(_monta_evento(r, CANCELADA_CRIACAO, pessoa,
                                                 r.get("WOLI_CreatedDate")))
                if eh_original:   # a ordem em si nasceu cancelada -> sem "Criada"
                    cid = r.get("CaseId")
                    if pd.notna(cid):
                        nascidas_canceladas.add(cid)
                continue

            # Reagendamento = WOLI que não é a original (novo agendamento)
            if not eh_original:
                pessoa = r.get(ATTR_REAGENDA)
                if normaliza_nome(pessoa) != "":
                    eventos.append(_monta_evento(r, "Reagendada", pessoa,
                                                 r.get(DATA_REAGENDA)))
            # Cancelamento pró-ativo (cancelada algum tempo depois de criada)
            if is_canc:
                pessoa = r.get(ATTR_CANCELA)
                if normaliza_nome(pessoa) != "":
                    eventos.append(_monta_evento(r, "Cancelada", pessoa,
                                                 r.get(DATA_CANCELA)))

    # ---- CRIAÇÃO: 1 evento por ordem, exceto as que já nascem canceladas ----
    base_ordem = (
        df.sort_values("WOLI_CreatedDate")
          .dropna(subset=["CaseId"])
          .drop_duplicates(subset=["CaseId"], keep="first")
    )
    for _, r in base_ordem.iterrows():
        if r.get("CaseId") in nascidas_canceladas:
            continue
        pessoa = r.get(ATTR_CRIACAO)
        if normaliza_nome(pessoa) == "":
            continue
        eventos.append(_monta_evento(r, "Criada", pessoa, r.get(DATA_CRIACAO)))

    ev = pd.DataFrame(eventos, columns=_colunas_evento())
    if ev.empty:
        return ev

    # Enriquecimento de pessoa/operação/tempo
    ev["Pessoa"] = ev["Pessoa_raw"].apply(lambda n: nome_canonico(n) or str(n))
    ev["Operacao"] = ev["Pessoa_raw"].apply(pessoa_para_operacao)
    ev["na_equipe"] = ev["Operacao"].notna()
    ev["Data_Evento"] = _to_dt_local(ev["Data_Evento"])
    ev = ev.dropna(subset=["Data_Evento"])
    ev["Dia"] = ev["Data_Evento"].dt.date
    ev["AnoMes"] = ev["Data_Evento"].dt.strftime("%Y-%m")
    ev["Operacao"] = ev["Operacao"].fillna("Outros / Fora da equipe")
    return ev


def _colunas_evento() -> list[str]:
    return [
        "Tipo_Evento", "Pessoa_raw", "Data_Evento",
        "Numero_OS", "LineItemNumber", "CaseNumber", "Cliente", "CodigoItem",
        "WO_TipoServico", "Segmento_Servico", "is_instalacao",
        "WOLI_Status", "WO_Status", "Case_Status",
        "Nome_Tecnico", "Motivo_Insucesso", "WorkOrderNumber", "WorkOrderId", "CaseId",
    ]


def _monta_evento(r: dict, tipo: str, pessoa, data) -> dict:
    return {
        "Tipo_Evento": tipo,
        "Pessoa_raw": pessoa,
        "Data_Evento": data,
        "Numero_OS": r.get("Numero_OS"),
        "LineItemNumber": r.get("LineItemNumber"),
        "CaseNumber": r.get("CaseNumber"),
        "Cliente": r.get("Cliente"),
        "CodigoItem": r.get("CodigoItem"),
        "WO_TipoServico": r.get("WO_TipoServico"),
        "Segmento_Servico": r.get("Segmento_Servico"),
        "is_instalacao": r.get("is_instalacao"),
        "WOLI_Status": r.get("WOLI_Status"),
        "WO_Status": r.get("WO_Status"),
        "Case_Status": r.get("Case_Status"),
        "Nome_Tecnico": r.get("Nome_Tecnico"),
        "Motivo_Insucesso": r.get("Motivo_Insucesso"),
        "WorkOrderNumber": r.get("WorkOrderNumber"),
        "WorkOrderId": r.get("WorkOrderId"),
        "CaseId": r.get("CaseId"),
    }


# ==========================================================================
# FILTROS E AGREGAÇÕES
# ==========================================================================
def meses_disponiveis(ev: pd.DataFrame) -> list[str]:
    if ev.empty:
        return []
    return sorted(ev["AnoMes"].dropna().unique().tolist())


def filtrar(ev: pd.DataFrame, ano_mes: str | None = None,
            operacoes: list[str] | None = None,
            pessoas: list[str] | None = None,
            apenas_equipe: bool = True) -> pd.DataFrame:
    out = ev
    if apenas_equipe:
        out = out[out["na_equipe"]]
    if ano_mes:
        out = out[out["AnoMes"] == ano_mes]
    if operacoes:
        out = out[out["Operacao"].isin(operacoes)]
    if pessoas:
        out = out[out["Pessoa"].isin(pessoas)]
    return out


def resumo_por_pessoa(ev: pd.DataFrame) -> pd.DataFrame:
    if ev.empty:
        return pd.DataFrame(columns=["Pessoa", "Operacao", *TIPOS_EVENTO, "Total"])
    piv = (ev.pivot_table(index=["Operacao", "Pessoa"], columns="Tipo_Evento",
                          values="Numero_OS", aggfunc="count", fill_value=0)
             .reset_index())
    for t in TIPOS_EVENTO:
        if t not in piv.columns:
            piv[t] = 0
    piv["Total"] = piv[TIPOS_EVENTO].sum(axis=1)
    piv = piv[["Pessoa", "Operacao", *TIPOS_EVENTO, "Total"]]
    return piv.sort_values("Total", ascending=False, ignore_index=True)


def resumo_por_operacao(ev: pd.DataFrame) -> pd.DataFrame:
    if ev.empty:
        return pd.DataFrame(columns=["Operacao", *TIPOS_EVENTO, "Total"])
    piv = (ev.pivot_table(index="Operacao", columns="Tipo_Evento",
                          values="Numero_OS", aggfunc="count", fill_value=0)
             .reset_index())
    for t in TIPOS_EVENTO:
        if t not in piv.columns:
            piv[t] = 0
    piv["Total"] = piv[TIPOS_EVENTO].sum(axis=1)
    piv = piv[["Operacao", *TIPOS_EVENTO, "Total"]]
    return piv.sort_values("Total", ascending=False, ignore_index=True)


def serie_diaria(ev: pd.DataFrame) -> pd.DataFrame:
    if ev.empty:
        return pd.DataFrame(columns=["Dia", *TIPOS_EVENTO])
    piv = (ev.pivot_table(index="Dia", columns="Tipo_Evento",
                          values="Numero_OS", aggfunc="count", fill_value=0)
             .reset_index())
    for t in TIPOS_EVENTO:
        if t not in piv.columns:
            piv[t] = 0
    piv = piv[["Dia", *TIPOS_EVENTO]]
    return piv.sort_values("Dia", ignore_index=True)


def reagendamentos_por_segmento(ev: pd.DataFrame) -> pd.DataFrame:
    """Instalação (Vendas) x Diversos, só para reagendamentos."""
    rea = ev[ev["Tipo_Evento"] == "Reagendada"]
    if rea.empty:
        return pd.DataFrame(columns=["Pessoa", "Instalação (Vendas)", "Diversos", "Total"])
    piv = (rea.pivot_table(index="Pessoa", columns="Segmento_Servico",
                           values="Numero_OS", aggfunc="count", fill_value=0)
              .reset_index())
    for c in ["Instalação (Vendas)", "Diversos"]:
        if c not in piv.columns:
            piv[c] = 0
    piv["Total"] = piv[["Instalação (Vendas)", "Diversos"]].sum(axis=1)
    return piv.sort_values("Total", ascending=False, ignore_index=True)


def canceladas_na_criacao_por_pessoa(ev: pd.DataFrame) -> pd.Series:
    """Contagem de 'Cancelada na criação' por pessoa (Series indexada por Pessoa)."""
    if ev.empty:
        return pd.Series(dtype="int64")
    m = ev[ev["Tipo_Evento"] == CANCELADA_CRIACAO]
    if m.empty:
        return pd.Series(dtype="int64")
    return m.groupby("Pessoa").size()


def kpis(ev: pd.DataFrame) -> dict:
    if ev.empty:
        return {"Criada": 0, "Reagendada": 0, "Cancelada": 0,
                "Cancelada_criacao": 0, "Total": 0, "Pessoas_ativas": 0}
    tc = ev["Tipo_Evento"]
    criada = int((tc == "Criada").sum())
    reag = int((tc == "Reagendada").sum())
    canc = int((tc == "Cancelada").sum())
    canc_cri = int((tc == CANCELADA_CRIACAO).sum())
    return {
        "Criada": criada, "Reagendada": reag, "Cancelada": canc,
        "Cancelada_criacao": canc_cri,
        # Total = ações produtivas (não inclui as ordens que já nascem canceladas)
        "Total": criada + reag + canc,
        "Pessoas_ativas": int(ev["Pessoa"].nunique()),
    }


# ==========================================================================
# TESTE ISOLADO
# ==========================================================================
if __name__ == "__main__":
    import datetime as _dt
    import random

    random.seed(7)
    pessoas = list(EQUIPE.keys())
    tipos_servico = ["Instalação", "Manutenção Preventiva", "Reparo", "Troca de Filtro"]
    status_finais = ["Executado com Sucesso", "Agendado", "Cancelado", "Reagendado"]

    linhas = []
    hoje = _dt.datetime(2026, 7, 21, 12, 0, tzinfo=_dt.timezone.utc)
    for wo in range(1, 121):
        criador = random.choice(pessoas)
        abertura = hoje - _dt.timedelta(days=random.randint(0, 40),
                                        hours=random.randint(0, 20))
        tipo_sv = random.choice(tipos_servico)
        n_wolis = random.randint(1, 4)  # 1 = só criada; >1 = houve reagendamento
        for i in range(n_wolis):
            criacao_woli = abertura + _dt.timedelta(days=i * random.randint(1, 5))
            ultima = criacao_woli + _dt.timedelta(hours=random.randint(0, 30))
            status = "Reagendado" if i < n_wolis - 1 else random.choice(status_finais)
            quem_mexeu = random.choice(pessoas)
            linhas.append({
                "WorkOrderId": f"WO{wo:04d}", "WorkOrderNumber": f"WO-{wo:04d}",
                "WO_CreatedDate": abertura.isoformat(), "WO_CreatedBy": criador,
                "WO_TipoServico": tipo_sv, "WO_Status": status,
                "CaseId": f"CASE{wo:04d}", "CaseNumber": f"{100000 + wo}",
                "Case_Type": "OS", "Case_Status": status, "Case_TipoSolicitacao": tipo_sv,
                "Case_CreatedBy": criador, "Case_CreatedDate": abertura.isoformat(),
                "Cliente": f"Cliente {wo}", "CodigoItem": f"ITEM{wo:05d}",
                "WOLI_Id": f"WOLI{wo:04d}{i}", "LineItemNumber": f"{wo:04d}-{i+1}",
                "Numero_OS": f"OS{200000 + wo*10 + i}", "WOLI_Status": status,
                "WOLI_CreatedDate": criacao_woli.isoformat(), "WOLI_CreatedBy": quem_mexeu,
                "WOLI_LastModifiedDate": ultima.isoformat(), "WOLI_LastModifiedBy": quem_mexeu,
                "Motivo_Insucesso": "" if status != "Cancelado" else "Cliente solicitou",
                "Nome_Tecnico": f"Tecnico {random.randint(1,15)}",
            })

    df = preparar_dataframe(pd.DataFrame(linhas))
    ev = construir_eventos(df)
    print("Total de WOLI :", len(df))
    print("Total eventos :", len(ev))
    print("Meses         :", meses_disponiveis(ev))
    print("KPIs (equipe, mês atual):")
    evf = filtrar(ev, ano_mes="2026-07")
    print("   ", kpis(evf))
    print("\nResumo por pessoa (mês atual):")
    print(resumo_por_pessoa(evf).to_string(index=False))
    print("\nResumo por operação:")
    print(resumo_por_operacao(evf).to_string(index=False))
    print("\nReagendamentos por segmento:")
    print(reagendamentos_por_segmento(evf).to_string(index=False))
    print("\nSérie diária (head):")
    print(serie_diaria(evf).head(8).to_string(index=False))
