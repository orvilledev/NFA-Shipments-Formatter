"""End-to-end verification against reference output."""
from openpyxl import load_workbook

from formatter import format_shipment

INPUT = r"c:\Users\Administrator\Downloads\FBA19FJ35VLB.xlsx"
OUTPUT = r"c:\Users\Administrator\Downloads\FBA19FJ35VLB-Box Contents.xlsx"


def read_cells(path_or_buf):
    wb = load_workbook(path_or_buf, data_only=False)
    result = {}
    for name in wb.sheetnames:
        ws = wb[name]
        result[name] = [
            [(c.value, type(c.value).__name__) for c in row if c.value is not None]
            for row in ws.iter_rows(values_only=False)
        ]
    return wb.sheetnames, result


with open(INPUT, "rb") as f:
    buf, filename, summary = format_shipment(f, INPUT)

assert filename == "FBA19FJ35VLB-Box Contents.xlsx"

gen_names, gen_data = read_cells(buf)
ref_names, ref_data = read_cells(OUTPUT)

assert gen_names == ref_names == ["Box Contents", "Weights and Dimensions"]

for sheet in ref_names:
    assert gen_data[sheet] == ref_data[sheet], f"Mismatch in {sheet}"

print("All checks passed.")
print(f"Filename: {filename}")
print(f"Summary: {summary}")
