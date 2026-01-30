from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import xlsxwriter


@dataclass(frozen=True)
class RmaEntry:
    recebimento: str
    cliente: str
    nf: str
    os: str
    triagem: str
    produto_enviado: str
    und: str
    plataforma: str
    codigo: str
    numero_serie: str
    status: str
    configuracao_avaria: str
    pedido_marketplace: str
    laudo_tecnico: str


def summarize_entries(entries: Iterable[RmaEntry]) -> tuple[list[tuple[str, int]], list[tuple[str, int]]]:
    pieces: Counter[str] = Counter()
    reasons: Counter[str] = Counter()

    for e in entries:
        produto = (e.produto_enviado or "").strip()
        if produto:
            pieces[produto] += 1

        avaria = (e.configuracao_avaria or "").strip()
        if produto and avaria:
            reason_key = f"{produto} ({avaria})"
        else:
            reason_key = produto or avaria

        reason_key = (reason_key or "").strip()
        if reason_key:
            reasons[reason_key] += 1

    pieces_sorted = sorted(pieces.items(), key=lambda x: (-x[1], x[0].casefold()))
    reasons_sorted = sorted(reasons.items(), key=lambda x: (-x[1], x[0].casefold()))
    return pieces_sorted, reasons_sorted


def export_to_excel(
    entries: list[RmaEntry],
    file_path: str | Path,
    *,
    title: str,
    periodo_mes: str,
    periodo_ano: str,
) -> Path:
    path = Path(file_path)
    path.parent.mkdir(parents=True, exist_ok=True)

    headers = [
        "RECEBIMENTO",
        "Cliente",
        "NF",
        "OS",
        "Triagem",
        "Produto enviado",
        "UND",
        "Plataforma",
        "Código",
        "Numero de serie",
        "Status",
        "Configuração/Avaria",
        "Pedido Marketplace",
        "LAUDO TÉCNICO",
    ]

    workbook = xlsxwriter.Workbook(str(path))

    fmt_title = workbook.add_format(
        {
            "bold": True,
            "font_size": 14,
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#D9D9D9",
            "border": 1,
        }
    )
    fmt_header = workbook.add_format(
        {
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#4472C4",
            "font_color": "#FFFFFF",
            "border": 1,
        }
    )
    fmt_header_laudo = workbook.add_format(
        {
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#4472C4",
            "font_color": "#FF0000",
            "border": 1,
        }
    )
    fmt_cell = workbook.add_format({"border": 1, "valign": "top"})
    fmt_cell_wrap = workbook.add_format({"border": 1, "valign": "top", "text_wrap": True})
    fmt_status_reparo = workbook.add_format({"border": 1, "valign": "top", "bg_color": "#C6EFCE", "font_color": "#006100"})
    fmt_status_reembolso = workbook.add_format({"border": 1, "valign": "top", "bg_color": "#FFC7CE", "font_color": "#9C0006"})

    ws = workbook.add_worksheet("RMA")
    ws.hide_gridlines(2)

    ws.merge_range(0, 0, 0, len(headers) - 1, title, fmt_title)
    ws.set_row(0, 24)

    for i, h in enumerate(headers):
        ws.write(1, i, h, fmt_header_laudo if h == "LAUDO TÉCNICO" else fmt_header)

    ws.freeze_panes(2, 0)

    col_widths = {
        0: 13,
        1: 22,
        2: 10,
        3: 10,
        4: 13,
        5: 26,
        6: 8,
        7: 16,
        8: 12,
        9: 18,
        10: 12,
        11: 34,
        12: 20,
        13: 44,
    }
    for col, w in col_widths.items():
        ws.set_column(col, col, w)

    for row_idx, e in enumerate(entries, start=2):
        row_values = [
            e.recebimento,
            e.cliente,
            e.nf,
            e.os,
            e.triagem,
            e.produto_enviado,
            e.und,
            e.plataforma,
            e.codigo,
            e.numero_serie,
            e.status,
            e.configuracao_avaria,
            e.pedido_marketplace,
            e.laudo_tecnico,
        ]
        for col_idx, v in enumerate(row_values):
            if col_idx == 10:
                status_lower = (v or "").lower().strip()
                if "reparo" in status_lower:
                    ws.write(row_idx, col_idx, v, fmt_status_reparo)
                elif "reembolso" in status_lower:
                    ws.write(row_idx, col_idx, v, fmt_status_reembolso)
                else:
                    ws.write(row_idx, col_idx, v, fmt_cell)
            else:
                use_wrap = col_idx in {11, 13}
                ws.write(row_idx, col_idx, v, fmt_cell_wrap if use_wrap else fmt_cell)

    pieces_sorted, reasons_sorted = summarize_entries(entries)

    ws2 = workbook.add_worksheet("Resumo")
    ws2.hide_gridlines(2)

    fmt_table_header = workbook.add_format(
        {
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#9DC3E6",
            "border": 1,
        }
    )
    fmt_table_header_qty = workbook.add_format(
        {
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#9DC3E6",
            "font_color": "#FF0000",
            "border": 1,
        }
    )
    fmt_table_cell = workbook.add_format({"border": 1, "bg_color": "#D9D9D9"})
    fmt_table_cell_center = workbook.add_format(
        {"border": 1, "bg_color": "#D9D9D9", "align": "center"}
    )
    fmt_total_label = workbook.add_format(
        {
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#000000",
            "font_color": "#FFFFFF",
            "border": 1,
        }
    )
    fmt_total_qty = workbook.add_format(
        {
            "bold": True,
            "align": "center",
            "valign": "vcenter",
            "bg_color": "#000000",
            "font_color": "#FF0000",
            "border": 1,
        }
    )

    ws2.set_column(0, 0, 44)
    ws2.set_column(1, 1, 14)
    ws2.set_column(3, 9, 18)

    pieces_start_row = 0
    ws2.write(pieces_start_row, 0, "PEÇAS DEFEITUOSAS", fmt_table_header)
    ws2.write(pieces_start_row, 1, "QUANTIDADE", fmt_table_header_qty)

    for i, (name, qty) in enumerate(pieces_sorted, start=1):
        ws2.write(pieces_start_row + i, 0, name, fmt_table_cell)
        ws2.write(pieces_start_row + i, 1, qty, fmt_table_cell_center)

    pieces_total_row = pieces_start_row + len(pieces_sorted) + 1
    ws2.write(pieces_total_row, 0, "TOTAL", fmt_total_label)
    ws2.write(pieces_total_row, 1, sum(q for _, q in pieces_sorted), fmt_total_qty)

    reasons_start_row = pieces_total_row + 3
    ws2.write(reasons_start_row, 0, "MOTIVOS DEFEITUOSOS", fmt_table_header)
    ws2.write(reasons_start_row, 1, "QUANTIDADE", fmt_table_header_qty)

    palette = [
        "#FFD966",
        "#F4B183",
        "#C6E0B4",
        "#9DC3E6",
        "#D9D2E9",
        "#F8CBAD",
        "#A9D18E",
        "#8FAADC",
        "#E2EFDA",
        "#C9C9C9",
    ]

    for i, (name, qty) in enumerate(reasons_sorted, start=1):
        color = palette[(i - 1) % len(palette)]
        fmt_reason = workbook.add_format({"border": 1, "bg_color": color})
        fmt_reason_qty = workbook.add_format({"border": 1, "bg_color": color, "align": "center"})
        ws2.write(reasons_start_row + i, 0, name, fmt_reason)
        ws2.write(reasons_start_row + i, 1, qty, fmt_reason_qty)

    reasons_total_row = reasons_start_row + len(reasons_sorted) + 1
    ws2.write(reasons_total_row, 0, "TOTAL", fmt_total_label)
    ws2.write(reasons_total_row, 1, sum(q for _, q in reasons_sorted), fmt_total_qty)

    if pieces_sorted:
        chart = workbook.add_chart({"type": "doughnut"})
        n = len(pieces_sorted)
        categories = ["Resumo", pieces_start_row + 1, 0, pieces_start_row + n, 0]
        values = ["Resumo", pieces_start_row + 1, 1, pieces_start_row + n, 1]
        points = [{"fill": {"color": palette[i % len(palette)]}} for i in range(n)]
        chart.add_series(
            {
                "categories": categories,
                "values": values,
                "data_labels": {"percentage": True},
                "points": points,
            }
        )
        chart.set_hole_size(60)
        chart.set_title({"name": f"PEÇAS DEFEITUOSAS\n{periodo_mes} - {periodo_ano}"})
        chart.set_legend({"position": "left"})
        chart.set_chartarea({"border": {"none": True}})

        ws2.insert_chart(1, 3, chart, {"x_scale": 1.4, "y_scale": 1.4})

    workbook.close()
    return path
