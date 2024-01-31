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
    chartArea2 = wb.Sheets(1).ChartObjects("Диаграмма 1").Chart
    chartArea.HasLegend = True
    chartArea.Legend.Font.Size = 9
    for series in chartArea.SeriesCollection():
        if series.Name == "NOW-BR":
            series.Format.Fill.ForeColor.RGB = 255  # Red color
            series.AxisGroup = 2
            series.ChartType = client.constants.xlLine
        # Check if series name ends with "-M" or "-F"
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
        if series.Name == "PLOST-M":
            series.Format.Fill.ForeColor.RGB = 255  # Red color
    
    def set_axis_labels(chartArea, xAxisLabel, yAxisLabel):
        # Set X-axis label
        x_axis = chartArea.Axes(client.constants.xlCategory, client.constants.xlPrimary)
        x_axis.HasTitle = True
        x_axis.AxisTitle.Text = xAxisLabel
        x_axis.AxisTitle.Font.Size = 8  
        # Set primary Y-axis label
        primary_y_axis = chartArea.Axes(client.constants.xlValue, client.constants.xlPrimary)
        primary_y_axis.HasTitle = True
        primary_y_axis.AxisTitle.Text = yAxisLabel
        primary_y_axis.AxisTitle.Font.Size = 8  

    #All Data
    x_axis = chartArea.Axes(client.constants.xlCategory, client.constants.xlPrimary)
    x_axis.HasTitle = True
    x_axis.AxisTitle.Text = "Number evt (unit)"
    x_axis.AxisTitle.Font.Size = 8  
    primary_y_axis = chartArea.Axes(client.constants.xlValue, client.constants.xlPrimary)
    primary_y_axis.HasTitle = True
    primary_y_axis.AxisTitle.Text = "Time (ms)"
    primary_y_axis.AxisTitle.Font.Size = 8  
    secondary_y_axis = chartArea.Axes(client.constants.xlValue, client.constants.xlSecondary)
    secondary_y_axis.HasTitle = True
    secondary_y_axis.AxisTitle.Text = "Bitrate (kbps)"
    secondary_y_axis.AxisTitle.Font.Size = 8  

    #PLOST
    primary_y_axis2 = chartArea2.Axes(client.constants.xlValue, client.constants.xlPrimary)
    primary_y_axis2.HasTitle = True
    primary_y_axis2.AxisTitle.Text = "Packet Lost"
    primary_y_axis2.AxisTitle.Font.Size = 8  
    x_axis2 = chartArea2.Axes(client.constants.xlCategory, client.constants.xlPrimary)
    x_axis2.HasTitle = True
    x_axis2.AxisTitle.Text = "Number evt (unit)"
    x_axis2.AxisTitle.Font.Size = 8

    # Chart Areas 3 to 5 - "Diagram 3" to "Diagram 5" JT, JT-A, JT-SDA
    for i in range(3, 6):
        chart_name = "Diagram " + str(i)
        chartArea = wb.Sheets(1).ChartObjects(chart_name).Chart
        set_axis_labels(chartArea, "Number evt (unit)", "Time (μs)")

    # Chart Areas 6 to 10 - "Diagram 6" to "Diagram 10" RTT, RTT-A, RTT-SDA, RTT-STRD, RTT-LTRD
    for i in range(6, 11):
        chart_name = "Diagram " + str(i)
        chartArea = wb.Sheets(1).ChartObjects(chart_name).Chart
        set_axis_labels(chartArea, "Number evt (unit)", "Time (ms)")

    
    xl.Quit()

    # Save the workbook and close Excel
    # wb.Save()
