class XlsxManager:
    def __init__(self, xlsx_file: str) -> None:
        self.file = pd.ExcelFile(xlsx_file)
        self.sheet_frames = {}
        for sheet in self.file.sheet_names:
            self.sheet_frames[sheet] = self.file.parse(sheet)

    @staticmethod
    def write_book(filename: str, sheets: dict) -> None:
        writer = pd.ExcelWriter(filename + ".xlsx", engine='xlsxwriter')
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name, startrow=1, header=False, index=False)
            worksheet = writer.sheets[name]
            workbook = writer.book
            # Get the dimensions of the dataframe.
            (max_row, max_col) = df.shape
            # Create a list of column headers, to use in add_table().
            column_settings = [{'header': column} for column in df.columns]
            # Add the Excel table structure. Pandas will add the data.
            worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings})
            # Make the columns wider for clarity.
            worksheet.set_column(0, max_col - 1, 12)
            if filename == "SprintBook" and name == "Tasks":
                XlsxManager.format_sprint(name, workbook, worksheet, max_row+1)
            elif filename == "ReportPortalSummary":
                XlsxManager.format_report_portal(workbook, worksheet, max_row+1)

        writer.save()

    @staticmethod
    def format_sprint(name: str, workbook, worksheet, max_row):
        format_r = workbook.add_format({'bg_color': '#FFC7CE',
                                        'font_color': '#9C0006'})

        format_y = workbook.add_format({'bg_color': '#FFEB9C',
                                        'font_color': '#9C6500'})

        format_g = workbook.add_format({'bg_color': '#C6EFCE',
                                        'font_color': '#006100'})
        if name == "Tasks":
            worksheet.conditional_format(
                f"G2:G{max_row}",
                {
                    "type": "text",
                    "criteria": "containing",
                    "value": "In Progress",
                    "format": format_y
                }
            )
            worksheet.conditional_format(
                f"G2:G{max_row}",
                {
                    "type": "text",
                    "criteria": "containing",
                    "value": "Done",
                    "format": format_g
                }
            )

    @staticmethod
    def format_report_portal(workbook, worksheet, max_row):
        format_r = workbook.add_format({'bg_color': '#FFC7CE',
                                        'font_color': '#9C0006'})

        format_y = workbook.add_format({'bg_color': '#FFEB9C',
                                        'font_color': '#9C6500'})

        format_g = workbook.add_format({'bg_color': '#C6EFCE',
                                        'font_color': '#006100'})

        worksheet.conditional_format(
            f"C2:C{max_row}",
            {
                "type": "text",
                "criteria": "containing",
                "value": "Failed",
                "format": format_r
            }
        )
        worksheet.conditional_format(
            f"C2:C{max_row}",
            {
                "type": "text",
                "criteria": "containing",
                "value": "Skipped",
                "format": format_y
            }
        )

    def get_sheet_names(self) -> list:
        return self.file.sheet_names
