from io import BytesIO

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter


def parse_input(source) -> tuple[str, list[dict]]:
    """Parse an FBA Carton Detail workbook into structured carton data."""
    wb = load_workbook(source, data_only=True)
    ws = wb.active

    shipment_id = None
    cartons: list[dict] = []
    current = None

    for row in ws.iter_rows(min_row=1, values_only=True):
        label = row[0]
        if label == "FBA Carton Detail" and row[1]:
            shipment_id = str(row[1]).strip()
        elif label == "Carton#:":
            if current:
                cartons.append(current)
            current = {
                "carton_id": str(row[1]),
                "items": [],
                "weight": None,
                "length": None,
                "width": None,
                "height": None,
            }
        elif current and label == "Total":
            current["weight"] = _as_int(row[5])
            current["length"] = _as_float(row[6])
            current["width"] = _as_float(row[7])
            current["height"] = _as_float(row[8])
        elif (
            current
            and label
            and label
            not in ("Sku", "Carton#:", "Total", "FBA Carton Detail", "Total Ctns:")
            and row[3] is not None
        ):
            current["items"].append({"upc": str(row[1]), "qty": int(row[3])})

    if current:
        cartons.append(current)

    if not shipment_id:
        raise ValueError(
            "Could not find shipment ID. Expected 'FBA Carton Detail' in cell A1 "
            "with the shipment ID in cell B1."
        )
    if not cartons:
        raise ValueError("No cartons found. Expected rows starting with 'Carton#:'.")

    return shipment_id, cartons


def build_output(shipment_id: str, cartons: list[dict]) -> tuple[BytesIO, str]:
    """Build the formatted Box Contents workbook."""
    wb = Workbook()

    ws_contents = wb.active
    ws_contents.title = "Box Contents"
    ws_contents.append(["Box Number", "UPC", "Qty"])

    for box_number, carton in enumerate(cartons, 1):
        for item in carton["items"]:
            ws_contents.append([box_number, item["upc"], item["qty"]])

    for row_idx in range(2, ws_contents.max_row + 1):
        ws_contents.cell(row=row_idx, column=2).number_format = "@"

    _auto_fit_columns(ws_contents, min_widths={"B": 14})

    ws_dims = wb.create_sheet("Weights and Dimensions")
    ws_dims.append(
        ["Box Number", "Weight", "Carton Length", "Carton Width", "Carton Height"]
    )
    for box_number, carton in enumerate(cartons, 1):
        ws_dims.append(
            [
                box_number,
                carton["weight"],
                carton["length"],
                carton["width"],
                carton["height"],
            ]
        )

    _auto_fit_columns(ws_dims)

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    filename = f"{shipment_id}-Box Contents.xlsx"
    return buffer, filename


def format_shipment(source) -> tuple[BytesIO, str, dict]:
    """Parse input and return output bytes, filename, and summary stats."""
    shipment_id, cartons = parse_input(source)
    output_buffer, filename = build_output(shipment_id, cartons)

    total_units = sum(item["qty"] for carton in cartons for item in carton["items"])
    total_weight = sum(carton["weight"] or 0 for carton in cartons)

    summary = {
        "shipment_id": shipment_id,
        "carton_count": len(cartons),
        "line_count": sum(len(c["items"]) for c in cartons),
        "total_units": total_units,
        "total_weight": total_weight,
    }
    return output_buffer, filename, summary


def _auto_fit_columns(ws, min_widths: dict[str, float] | None = None) -> None:
    """Set column widths from content so values are visible when the file opens."""
    min_widths = min_widths or {}
    for col_idx in range(1, ws.max_column + 1):
        column_letter = get_column_letter(col_idx)
        max_length = 0
        for row in ws.iter_rows(
            min_row=1, max_row=ws.max_row, min_col=col_idx, max_col=col_idx
        ):
            for cell in row:
                if cell.value is not None:
                    max_length = max(max_length, len(str(cell.value)))
        width = max(max_length + 2, min_widths.get(column_letter, 0))
        ws.column_dimensions[column_letter].width = width


def _as_int(value):
    if value is None:
        return None
    return int(value)


def _as_float(value):
    if value is None:
        return None
    return float(value)
