import util.excel_util
import xlwings as xw

def test_read_excel_outputs():
    # Start Excel app (hidden so it doesn’t pop up)
    app = xw.App(visible=False)
    try:
        wb = app.books.add()
        sheet = wb.sheets[0]

        # Write test values
        sheet["A1"].value = 123.456789
        sheet["B1"].value = "Text"
        sheet["C1:C3"].value = [[1.111], [2.222], [3.333]]
        sheet["D1:E2"].value = [[10.555, 20.666], [30.777, 40.888]]

        # Define output cells to test
        output_cells = {
            "single_value": "A1",
            "string_value": "B1",
            "column_range": "C1:C3",
            "matrix_range": "D1:E2",
        }

        # Run the function
        results = util.excel_util.read_excel_outputs(sheet, output_cells, decimals=2)

        # Print results
        print("Results:")
        for key, val in results.items():
            print(f"{key}: {val}")

    finally:
        wb.close()
        app.quit()


if __name__ == "__main__":
    test_read_excel_outputs()