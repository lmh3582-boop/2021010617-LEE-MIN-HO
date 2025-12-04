import os
import pandas as pd
import xlsxwriter


def main():
    csv_name = input("Enter SU2 CSV file name (e.g., 032.csv): ").strip()
    if not csv_name:
        print("No file name provided.")
        return

    if not os.path.exists(csv_name):
        print(f"Error: file '{csv_name}' not found.")
        return

    section = input("Enter section percentage for the title (e.g., 50): ").strip()
    if not section:
        section = "50"

    df = pd.read_csv(csv_name)

    required_cols = ["Points_0", "Pressure_Coefficient"]
    for col in required_cols:
        if col not in df.columns:
            print(f"Error: column '{col}' not found in CSV file.")
            return

    data = df[["Points_0", "Pressure_Coefficient"]].copy()
    data.rename(columns={"Points_0": "x_over_c",
                         "Pressure_Coefficient": "Cp"}, inplace=True)
    data.sort_values("x_over_c", inplace=True)
    data.reset_index(drop=True, inplace=True)

    output_name = f"Section_{section}_SU2.xlsx"
    workbook = xlsxwriter.Workbook(output_name)
    sheet = workbook.add_worksheet("Data")

    # write headers
    sheet.write(0, 0, "x/C")
    sheet.write(0, 1, "Cp")

    # write data
    for i, row in data.iterrows():
        sheet.write(i + 1, 0, float(row["x_over_c"]))
        sheet.write(i + 1, 1, float(row["Cp"]))

    # create chart
    chart = workbook.add_chart({"type": "scatter", "subtype": "straight"})
    last_row = len(data)

    chart.add_series({
        "name":       "SU2",
        "categories": ["Data", 1, 0, last_row, 0],  # x/C
        "values":     ["Data", 1, 1, last_row, 1],  # Cp
        "line":       {"color": "black", "width": 2},
        "marker":     {"type": "none"},
    })

    chart.set_title({"name": f"Section {section}%"})
    chart.set_x_axis({"name": "x/C"})
    chart.set_y_axis({"name": "Cp"})
    chart.set_legend({"position": "right"})

    sheet.insert_chart("E2", chart)

    workbook.close()
    print(f"Saved Excel file: {output_name}")


if __name__ == "__main__":
    main()
