"""
parse_excel.py — le a aba BASE_PCM do arquivo Excel do PCM e gera os arquivos
JSON que o painel (index.html) consome.

Uso:
    python scripts/parse_excel.py --file uploads/PCM_Indicadores.xlsm

Por padrao processa a ULTIMA data encontrada na planilha. Para processar uma
data especifica:
    python scripts/parse_excel.py --file uploads/PCM_Indicadores.xlsm --date 2026-09-03
Para reprocessar TODO o historico:
    python scripts/parse_excel.py --file uploads/PCM_Indicadores.xlsm --date all
"""
import argparse
import json
import os
import sys
from datetime import datetime

import pandas as pd

SHEET = "BASE_PCM"
SHEET_HORAS = "horas_disponiveis"

MONTH_ORDER = [
    "janeiro", "fevereiro", "março", "abril", "maio", "junho",
    "julho", "agosto", "setembro", "outubro", "novembro", "dezembro",
]

# A Automação só responde pelas 4 máquinas do robô (Mespack R1/R2/E/F).
# Elétrica e Mecânica respondem por TODAS as máquinas monitoradas.
AUTOMACAO_EQUIPS = {"MESPACK R1", "MESPACK R2", "MESPACK E", "MESPACK F"}

# Fallback para quando não há linha de horas_disponiveis para a data
# (ex.: datas fora do período coberto pela aba). 22h/máquina, igual ao
# padrão já usado no Excel para dias sem programação do PCP.
FALLBACK_HORAS_POR_EQUIP = 22.0
FALLBACK_TOTAL_EQUIPS = 13


def calc_disponibilidade(sub, hd_df, target_date):
    """Disponibilidade/Indisponibilidade geral e por departamento, usando as
    horas disponíveis REAIS da aba horas_disponiveis (programação do PCP),
    em vez de uma constante fixa de 24h."""
    if hd_df is not None:
        hd_dia = hd_df[hd_df["Data"].dt.date == target_date]
    else:
        hd_dia = None

    if hd_dia is not None and not hd_dia.empty:
        horas_disp_total = float(hd_dia["Minutos"].sum()) / 60.0
        horas_disp_automacao = float(
            hd_dia[hd_dia["Equipamento"].isin(AUTOMACAO_EQUIPS)]["Minutos"].sum()
        ) / 60.0
    else:
        horas_disp_total = FALLBACK_TOTAL_EQUIPS * FALLBACK_HORAS_POR_EQUIP
        horas_disp_automacao = len(AUTOMACAO_EQUIPS) * FALLBACK_HORAS_POR_EQUIP

    horas_paradas_total = float(sub["Tempo Parada (h decimal)"].sum())
    horas_paradas_dept = sub.groupby("Departamento")["Tempo Parada (h decimal)"].sum()

    def _monta(horas_disp, horas_paradas):
        if horas_disp and horas_disp > 0:
            disp_pct = max(0.0, (1 - horas_paradas / horas_disp)) * 100
            indisp_pct = 100 - disp_pct
        else:
            disp_pct = None
            indisp_pct = None
        return {
            "horas_disponiveis": round(horas_disp, 4),
            "horas_paradas": round(horas_paradas, 4),
            "disponibilidade_pct": round(disp_pct, 4) if disp_pct is not None else None,
            "indisponibilidade_pct": round(indisp_pct, 4) if indisp_pct is not None else None,
        }

    departamentos = {}
    for dept in ["Elétrica", "Mecânica", "Automação"]:
        horas_disp_dept = horas_disp_automacao if dept == "Automação" else horas_disp_total
        horas_paradas = float(horas_paradas_dept.get(dept, 0.0))
        departamentos[dept] = _monta(horas_disp_dept, horas_paradas)

    return {
        "geral": _monta(horas_disp_total, horas_paradas_total),
        "departamentos": departamentos,
    }


def calc_mensal(df, hd_df, target_date, meta_mtbf):
    """Acumulado do mês (do dia 1 até target_date) de TODOS os departamentos
    somados: falhas, horas paradas, horas disponíveis reais, MTBF mensal (min),
    disponibilidade e status contra a meta. Mesma conta da aba IMPORTAÇÃO NECTAR
    (H23:I29): MTBF = ((H disponíveis - H paradas) / falhas) * 60."""
    inicio = pd.Timestamp(target_date.year, target_date.month, 1)
    fim = pd.Timestamp(target_date) + pd.Timedelta(days=1)  # exclusivo

    mes = df[(df["Data Inicio"] >= inicio) & (df["Data Inicio"] < fim)]
    falhas = int(mes["Falhas"].sum())
    horas_paradas = float(mes["Tempo Parada (h decimal)"].sum())

    if hd_df is not None:
        hd_mes = hd_df[(hd_df["Data"] >= inicio) & (hd_df["Data"] < fim)]
        horas_disp = float(hd_mes["Minutos"].sum()) / 60.0
    else:
        horas_disp = 0.0

    mtbf_min = None
    disp_pct = None
    if horas_disp > 0:
        disp_pct = max(0.0, (1 - horas_paradas / horas_disp)) * 100
        if falhas > 0:
            mtbf_min = ((horas_disp - horas_paradas) / falhas) * 60

    if mtbf_min is None:
        status = None
    else:
        status = "atingida" if mtbf_min >= meta_mtbf else "abaixo"

    return {
        "mes": MONTH_ORDER[target_date.month - 1],
        "ano": target_date.year,
        "ate": target_date.isoformat(),
        "falhas": falhas,
        "horas_paradas": round(horas_paradas, 4),
        "horas_disponiveis": round(horas_disp, 4),
        "mtbf_min": round(mtbf_min, 4) if mtbf_min is not None else None,
        "disponibilidade_pct": round(disp_pct, 4) if disp_pct is not None else None,
        "indisponibilidade_pct": round(100 - disp_pct, 4) if disp_pct is not None else None,
        "meta_mtbf": meta_mtbf,
        "status": status,
    }


def horas_disponiveis_por_equip(hd_df, target_date):
    """dict {equipamento: horas disponíveis reais naquele dia}, com fallback de 22h
    para equipamentos sem linha na aba horas_disponiveis nessa data."""
    if hd_df is None:
        return {}
    hd_dia = hd_df[hd_df["Data"].dt.date == target_date]
    if hd_dia.empty:
        return {}
    return (hd_dia.groupby("Equipamento")["Minutos"].sum() / 60.0).to_dict()


def process_date(df, target_date, args, hd_df=None):
    sub = df[df["Data Inicio"].dt.date == target_date].copy()
    if sub.empty:
        return None

    hd_equip = horas_disponiveis_por_equip(hd_df, target_date)

    g = (
        sub.groupby("Equipamento")
        .agg(horas=("Tempo Parada (h decimal)", "sum"), falhas=("Falhas", "sum"))
        .reset_index()
        .sort_values("horas", ascending=False)
    )
    # detalhamento por turno de cada equipamento (usado no filtro "por turno" do painel)
    piv_equip_turno = sub.groupby(["Equipamento", "Turno"]).agg(
        horas=("Tempo Parada (h decimal)", "sum"), falhas=("Falhas", "sum")
    ).reset_index()
    equip = []
    for _, r in g.iterrows():
        por_turno = {}
        for turno in [1, 2, 3]:
            linha = piv_equip_turno[
                (piv_equip_turno["Equipamento"] == r["Equipamento"]) & (piv_equip_turno["Turno"] == turno)
            ]
            if len(linha):
                por_turno[str(turno)] = {
                    "horas": round(float(linha["horas"].values[0]), 4),
                    "falhas": int(linha["falhas"].values[0]),
                }
            else:
                por_turno[str(turno)] = {"horas": 0.0, "falhas": 0}
        equip.append({
            "name": r["Equipamento"],
            "horas": round(float(r["horas"]), 4),
            "horas_disponiveis": round(float(hd_equip.get(r["Equipamento"], FALLBACK_HORAS_POR_EQUIP)), 4),
            "falhas": int(r["falhas"]),
            "por_turno": por_turno,
        })

    piv = sub.groupby(["Departamento", "Turno"])["Tempo Parada (h decimal)"].sum().reset_index()
    turnos = {}
    for dept in sorted(sub["Departamento"].dropna().unique()):
        turnos[dept] = {}
        for turno in [1, 2, 3]:
            val = piv[(piv["Departamento"] == dept) & (piv["Turno"] == turno)]["Tempo Parada (h decimal)"]
            turnos[dept][str(turno)] = round(float(val.values[0]), 4) if len(val) else 0.0

    gc_source = sub.copy()
    # Componente depende de um VLOOKUP externo que pode estar quebrado (link externo no Excel).
    # Quando isso acontece, cai para a coluna MOTIVO como alternativa, pra não perder os dados de falha.
    gc_source["Componente"] = gc_source["Componente"].fillna(gc_source["MOTIVO"])
    gc = (
        gc_source.groupby(["TAG", "Componente"])
        .agg(falhas=("Falhas", "sum"), horas=("Tempo Parada (h decimal)", "sum"))
        .reset_index()
    )
    # equipamento de cada TAG (assume-se 1 TAG = 1 equipamento; usa o mais frequente por segurança)
    equipamento_por_tag = (
        gc_source.groupby("TAG")["Equipamento"]
        .agg(lambda s: s.mode().iloc[0] if not s.mode().empty else s.iloc[0])
    )
    piv_falha_turno = gc_source.groupby(["TAG", "Componente", "Turno"]).agg(
        horas=("Tempo Parada (h decimal)", "sum"), falhas=("Falhas", "sum")
    ).reset_index()

    def _monta_lista(df_ordenado):
        lista = []
        for _, r in df_ordenado.iterrows():
            por_turno = {}
            for turno in [1, 2, 3]:
                linha = piv_falha_turno[
                    (piv_falha_turno["TAG"] == r["TAG"]) & (piv_falha_turno["Componente"] == r["Componente"]) & (piv_falha_turno["Turno"] == turno)
                ]
                if len(linha):
                    por_turno[str(turno)] = {
                        "horas": round(float(linha["horas"].values[0]), 4),
                        "falhas": int(linha["falhas"].values[0]),
                    }
                else:
                    por_turno[str(turno)] = {"horas": 0.0, "falhas": 0}
            lista.append({
                "tag": str(r["TAG"]),
                "componente": r["Componente"],
                "equipamento": equipamento_por_tag.get(r["TAG"], ""),
                "falhas": int(r["falhas"]),
                "horas": round(float(r["horas"]), 4),
                "por_turno": por_turno,
            })
        return lista

    # gráfico "Top Falhas" — mostra quem mais falhou (por número de falhas)
    top_falhas = _monta_lista(gc.sort_values("falhas", ascending=False).head(12))
    # cards "Parada por Componentes" — mostra quem mais tempo ficou parado (por tempo)
    parada_componentes = _monta_lista(gc.sort_values("horas", ascending=False).head(15))

    mg = df.groupby("Mês")["Falhas"].sum()
    evolucao = [{"mes": m, "falhas": int(mg[m])} for m in MONTH_ORDER if m in mg.index]

    disponibilidade = calc_disponibilidade(sub, hd_df, target_date)

    return {
        "date": target_date.isoformat(),
        "updated_at": datetime.now().isoformat(timespec="minutes"),
        "meta_mtbf": args.meta_mtbf,
        "dias_parados": args.dias_parados,
        "equip": equip,
        "turnos": turnos,
        "top_falhas": top_falhas,
        "parada_componentes": parada_componentes,
        "evolucao_mensal": evolucao,
        "disponibilidade": disponibilidade,
        "mensal": calc_mensal(df, hd_df, target_date, args.meta_mtbf),
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", required=True, help="Caminho do arquivo .xlsm/.xlsx")
    ap.add_argument("--date", default=None, help="Data AAAA-MM-DD, 'all' para reprocessar todo o histórico, ou omitido para a última data da planilha.")
    ap.add_argument("--outdir", default="data", help="Pasta de saida dos JSON")
    ap.add_argument("--meta-mtbf", type=float, default=300, help="Meta MTBF mensal (minutos)")
    ap.add_argument("--dias-parados", type=int, default=0, help="KPI Dias Parados (definido manualmente)")
    args = ap.parse_args()

    if not os.path.exists(args.file):
        print(f"ERRO: arquivo nao encontrado: {args.file}", file=sys.stderr)
        sys.exit(1)

    df = pd.read_excel(args.file, sheet_name=SHEET)
    df["Data Inicio"] = pd.to_datetime(df["Data Inicio"], errors="coerce")

    # Linhas sem data (vazias ou com valor inválido) não têm como entrar em nenhum dia.
    # Em vez de quebrar o robô inteiro, ignora essas linhas e avisa quantas foram.
    sem_data = df["Data Inicio"].isna()
    if sem_data.any():
        linhas_excel = (df.index[sem_data] + 2).tolist()  # +2: cabeçalho + índice começa em 0
        print(
            f"AVISO: {int(sem_data.sum())} linha(s) sem 'Data Inicio' foram ignoradas "
            f"(linhas do Excel: {linhas_excel[:20]}{'...' if len(linhas_excel) > 20 else ''})",
            file=sys.stderr,
        )
        df = df[~sem_data].copy()

    if df.empty:
        print("ERRO: nenhuma linha com 'Data Inicio' válida encontrada na planilha.", file=sys.stderr)
        sys.exit(1)

    # Horas disponíveis reais por equipamento/dia (programação do PCP). Se a aba
    # não existir no arquivo (planilhas antigas), segue com o fallback de 22h.
    try:
        hd_df = pd.read_excel(args.file, sheet_name=SHEET_HORAS)
        hd_df["Data"] = pd.to_datetime(hd_df["Data"], errors="coerce")
        hd_df = hd_df.dropna(subset=["Data"])
    except Exception as e:
        print(f"AVISO: aba '{SHEET_HORAS}' não encontrada/legível ({e}); usando fallback de 22h/máquina.", file=sys.stderr)
        hd_df = None

    if args.date == "all":
        target_dates = sorted(df["Data Inicio"].dt.date.unique())
    elif args.date:
        target_dates = [pd.to_datetime(args.date).date()]
    else:
        target_dates = [df["Data Inicio"].max().date()]

    os.makedirs(args.outdir, exist_ok=True)
    history_path = os.path.join(args.outdir, "history.json")
    history = []
    if os.path.exists(history_path):
        with open(history_path, "r", encoding="utf-8") as f:
            history = json.load(f)

    last_out = None
    for target_date in target_dates:
        out = process_date(df, target_date, args, hd_df=hd_df)
        if out is None:
            print(f"AVISO: nenhum registro para {target_date}, pulando", file=sys.stderr)
            continue

        day_path = os.path.join(args.outdir, f"{target_date.isoformat()}.json")
        with open(day_path, "w", encoding="utf-8") as f:
            json.dump(out, f, ensure_ascii=False, indent=2)

        total_horas = round(sum(e["horas"] for e in out["equip"]), 4)
        history = [h for h in history if h["date"] != out["date"]]
        history.append({"date": out["date"], "horas": total_horas})

        last_out = out
        print(f"OK: {day_path} gerado ({len(out['equip'])} equipamentos, {total_horas}h paradas no dia)")

    history.sort(key=lambda h: h["date"])
    with open(history_path, "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=2)

    if last_out is not None:
        latest_path = os.path.join(args.outdir, "latest.json")
        with open(latest_path, "w", encoding="utf-8") as f:
            json.dump(last_out, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
