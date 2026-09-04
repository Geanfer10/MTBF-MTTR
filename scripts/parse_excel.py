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


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", required=True, help="Caminho do arquivo .xlsm/.xlsx")
    ap.add_argument("--date", default=None, help="Data AAAA-MM-DD. Padrao: ultima data da planilha.")
    ap.add_argument("--outdir", default="data", help="Pasta de saida dos JSON")
    ap.add_argument("--meta-mtbf", type=float, default=300, help="Meta MTBF mensal (h)")
    ap.add_argument("--dias-parados", type=int, default=0, help="KPI Dias Parados (definido manualmente)")
    args = ap.parse_args()

    if not os.path.exists(args.file):
        print(f"ERRO: arquivo nao encontrado: {args.file}", file=sys.stderr)
        sys.exit(1)

    df = pd.read_excel(args.file, sheet_name=SHEET)
    df["Data Inicio"] = pd.to_datetime(df["Data Inicio"])

    if args.date:
        target_date = pd.to_datetime(args.date).date()
    else:
        target_date = df["Data Inicio"].max().date()

    sub = df[df["Data Inicio"].dt.date == target_date].copy()
    if sub.empty:
        print(f"ERRO: nenhum registro encontrado para {target_date}", file=sys.stderr)
        sys.exit(1)

    # ---- horas paradas e falhas por equipamento (só entram equipamentos que rodaram/pararam no dia) ----
    g = (
        sub.groupby("Equipamento")
        .agg(horas=("Tempo Parada (h decimal)", "sum"), falhas=("Falhas", "sum"))
        .reset_index()
        .sort_values("horas", ascending=False)
    )
    equip = [
        {"name": r["Equipamento"], "horas": round(float(r["horas"]), 4), "falhas": int(r["falhas"])}
        for _, r in g.iterrows()
    ]

    # ---- paradas por turno, agrupadas por departamento ----
    piv = sub.groupby(["Departamento", "Turno"])["Tempo Parada (h decimal)"].sum().reset_index()
    turnos = {}
    for dept in sorted(sub["Departamento"].dropna().unique()):
        turnos[dept] = {}
        for turno in [1, 2, 3]:
            val = piv[(piv["Departamento"] == dept) & (piv["Turno"] == turno)]["Tempo Parada (h decimal)"]
            turnos[dept][str(turno)] = round(float(val.values[0]), 4) if len(val) else 0.0

    # ---- top falhas por componente (TAG + Componente), usado no Top 10 e em Parada por Componentes ----
    gc = (
        sub.groupby(["TAG", "Componente"])
        .agg(falhas=("Falhas", "sum"), horas=("Tempo Parada (h decimal)", "sum"))
        .reset_index()
        .sort_values("falhas", ascending=False)
    )
    top_falhas = [
        {
            "tag": str(r["TAG"]),
            "componente": r["Componente"],
            "falhas": int(r["falhas"]),
            "horas": round(float(r["horas"]), 4),
        }
        for _, r in gc.head(12).iterrows()
    ]

    # ---- evolução mensal de falhas (usa a planilha inteira, não só o dia filtrado) ----
    mg = df.groupby("Mês")["Falhas"].sum()
    evolucao = [{"mes": m, "falhas": int(mg[m])} for m in MONTH_ORDER if m in mg.index]

    out = {
        "date": target_date.isoformat(),
        "updated_at": datetime.now().isoformat(timespec="minutes"),
        "meta_mtbf": args.meta_mtbf,
        "dias_parados": args.dias_parados,
        "equip": equip,
        "turnos": turnos,
        "top_falhas": top_falhas,
        "evolucao_mensal": evolucao,
    }

    os.makedirs(args.outdir, exist_ok=True)

    day_path = os.path.join(args.outdir, f"{target_date.isoformat()}.json")
    with open(day_path, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)

    latest_path = os.path.join(args.outdir, "latest.json")
    with open(latest_path, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)

    # ---- histórico acumulado (usado no gráfico de Tendência das Horas Paradas) ----
    history_path = os.path.join(args.outdir, "history.json")
    history = []
    if os.path.exists(history_path):
        with open(history_path, "r", encoding="utf-8") as f:
            history = json.load(f)

    total_horas = round(sum(e["horas"] for e in equip), 4)
    history = [h for h in history if h["date"] != out["date"]]
    history.append({"date": out["date"], "horas": total_horas})
    history.sort(key=lambda h: h["date"])

    with open(history_path, "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=2)

    print(f"OK: {day_path} gerado ({len(equip)} equipamentos, {total_horas}h paradas no dia)")


if __name__ == "__main__":
    main()
