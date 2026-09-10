"""
parse_excel.py — le a aba BASE_PCM do arquivo Excel do PCM e gera os arquivos
JSON que o painel (index.html) consome.

Uso:
    python scripts/parse_excel.py --file uploads/PCM_Indicadores.xlsm

Por padrao processa a ULTIMA data encontrada na planilha. Para processar uma
data especifica:
    python scripts/parse_excel.py --file uploads/PCM_Indicadores.xlsm --date 2026-09-03
"""
import argparse
import json
import os
import sys
from datetime import datetime

import pandas as pd

SHEET = "BASE_PCM"

MONTH_ORDER = [
    "janeiro", "fevereiro", "março", "abril", "maio", "junho",
    "julho", "agosto", "setembro", "outubro", "novembro", "dezembro",
]


def process_date(df, target_date, args):
    sub = df[df["Data Inicio"].dt.date == target_date].copy()
    if sub.empty:
        return None

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
        .sort_values("falhas", ascending=False)
    )
    piv_falha_turno = gc_source.groupby(["TAG", "Componente", "Turno"]).agg(
        horas=("Tempo Parada (h decimal)", "sum"), falhas=("Falhas", "sum")
    ).reset_index()
    top_falhas = []
    for _, r in gc.head(12).iterrows():
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
        top_falhas.append({
            "tag": str(r["TAG"]),
            "componente": r["Componente"],
            "falhas": int(r["falhas"]),
            "horas": round(float(r["horas"]), 4),
            "por_turno": por_turno,
        })

    mg = df.groupby("Mês")["Falhas"].sum()
    evolucao = [{"mes": m, "falhas": int(mg[m])} for m in MONTH_ORDER if m in mg.index]

    return {
        "date": target_date.isoformat(),
        "updated_at": datetime.now().isoformat(timespec="minutes"),
        "meta_mtbf": args.meta_mtbf,
        "dias_parados": args.dias_parados,
        "equip": equip,
        "turnos": turnos,
        "top_falhas": top_falhas,
        "evolucao_mensal": evolucao,
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", required=True, help="Caminho do arquivo .xlsm/.xlsx")
    ap.add_argument("--date", default=None, help="Data AAAA-MM-DD, 'all' para reprocessar todo o histórico, ou omitido para a última data da planilha.")
    ap.add_argument("--outdir", default="data", help="Pasta de saida dos JSON")
    ap.add_argument("--meta-mtbf", type=float, default=300, help="Meta MTBF mensal (h)")
    ap.add_argument("--dias-parados", type=int, default=0, help="KPI Dias Parados (definido manualmente)")
    args = ap.parse_args()

    if not os.path.exists(args.file):
        print(f"ERRO: arquivo nao encontrado: {args.file}", file=sys.stderr)
        sys.exit(1)

    df = pd.read_excel(args.file, sheet_name=SHEET)
    df["Data Inicio"] = pd.to_datetime(df["Data Inicio"])

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
        out = process_date(df, target_date, args)
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


if __name__ == "__main__":
    main()

