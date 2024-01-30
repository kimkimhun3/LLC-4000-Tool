import sys
from win32com.client import DispatchEx
from win32com import client

if __name__ == "__main__":
    # Check if the correct number of command-line arguments is provided
    if len(sys.argv) != 2:
        print("Usage: python python.py <filePath>")
        sys.exit(1)

    # Retrieve the filePath from command-line arguments
    filePath = sys.argv[1]

    # Initialize Excel and open workbook
    xl = DispatchEx("Excel.Application")
    xl.Visible = True
    wb = xl.Workbooks.Add(filePath)

    chartArea = wb.Sheets(1).ChartObjects("Diagram 2").Chart
    chartArea.HasLegend = True
    for series in chartArea.SeriesCollection():
        # Check if series name ends with "-M" or "-F"
        if series.Name == "NOW-BR":
            series.Format.Fill.ForeColor.RGB = 255  # Red color
            series.AxisGroup = 2
            series.ChartType = client.constants.xlLine
        if series.Name.endswith(("-M", "-F")):
            non_zero_values = []  # List to hold non-zero values
            x_values = []  # List to hold corresponding X values
            for i, value in enumerate(series.Values):
                if value != 0:
                    non_zero_values.append(value)
                    x_values.append(i)  # Keep track of the original index

            # Clear existing data and set filtered values
            series.Values = non_zero_values
            series.XValues = x_values
    xl.Quit()

    # Save the workbook and close Excel
    # wb.Save()
