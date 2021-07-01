from openpyxl import load_workbook

dynamodb_response = [
    {
        "Metrics": "Completed Impression Rate (%)",
        "Last Full Week": "95.67",
        "Date": "2021-06-30",
        "Prior Full Week": "96.56",
        "Target / Benchmark": "90",
        "MTD": "97.93",
        "This Quarter": "98.18"
    },
    {
        "Metrics": "Fill rate | Ad only (%)",
        "Last Full Week": "82.15",
        "Date": "2021-06-30",
        "Prior Full Week": "85.12",
        "Target / Benchmark": "-",
        "MTD": "82.7",
        "This Quarter": "81.12"
    },
    {
        "Metrics": "Fill rate | Ad+Promo (%)",
        "Last Full Week": "90.56",
        "Date": "2021-06-30",
        "Prior Full Week": "91.45",
        "Target / Benchmark": "100",
        "MTD": "92.56",
        "This Quarter": "93.05"
    },
    {
        "Metrics": "Uses Impression Rate (%)",
        "Last Full Week": "63.51",
        "Date": "2021-06-30",
        "Prior Full Week": "66.48",
        "Target / Benchmark": "60",
        "MTD": "66.51",
        "This Quarter": "67.16"
    }
]


class ExcelAutomator(object):
    # Class Instantiation
    def __init__(self, file_name) -> None:
        self.file_name = file_name
        self.wb = load_workbook(file_name)

    # Get sheet from the workbook
    def get_sheet(self, sheet_name):
        ws = self.wb[sheet_name]
        return ws

    # DAI metrics updating in excel template file
    def updated_dai(self, ws, response) -> None:
        # Required Column Headers
        headers = {col_name.value: col_name.column_letter for col_name in list(ws.rows)[0][1:11]}

        # Metrics search and header column letter rearrangement
        cell = []
        for each in response:
            row = self.search_value_in_col_idx(ws, each['Metrics'])
            cell.append([(f"{headers[key]}{row}", key) for key in response[0].keys() if key not in ['Date', 'Metrics']])

        # Assign cell values to respective cells for all DAI metrics
        for each_row, data in zip(cell, response):
            ws[each_row[0][0]] = data[each_row[0][1]]
            ws[each_row[1][0]] = data[each_row[1][1]]
            ws[each_row[2][0]] = data[each_row[2][1]]
            ws[each_row[3][0]] = data[each_row[3][1]]
            ws[each_row[4][0]] = data[each_row[4][1]]

    # Search row number of each metrics for cell value generation
    @staticmethod
    def search_value_in_col_idx(ws, metrics_name, col_idx=1):
        for row in range(1, ws.max_row + 1):
            if ws[row][col_idx].value == metrics_name:
                return row
        print(f'{metrics_name} not found!!')

    # Save the changes in the workbook
    def save(self, file_name):
        self.wb.save(filename=file_name)


# Main Constructor
if __name__ == "__main__":
    filename = 'Team Metrics.xlsx'
    excel_automator = ExcelAutomator(filename)
    work_sheet = excel_automator.get_sheet('SlingTV')
    excel_automator.updated_dai(work_sheet, dynamodb_response)
    excel_automator.save(filename)

