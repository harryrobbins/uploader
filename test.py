import openpyxl

from app import validate

from pathlib import Path

import networkx as nx

if __name__ == "__main__":
    errors = validate(Path("./org_chart.xlsx"))
    print("file with no errors:", errors)

    print("------------------")

    errors = validate(Path("./file_with_sheet_name_error.xlsx"))
    print("file with sheets errors:", errors)

    print("------------------")

    errors = validate(Path("./file_with_headings_error.xlsx"))
    print("file with headings errors:", errors)

    print("------------------")

    errors = validate(Path("./file_with_cyclical_reference.xlsx"))
    print("file with cyclical errors:", errors)


    # print(G.edges)