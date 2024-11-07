import sys
from win32com.client import DispatchEx
from win32com import client
import os

if __name__ == "__main__":
    # Check if the correct number of command-line arguments is provided
    if len(sys.argv) != 2:
        print("Usage: python python.py <filePath>")
        sys.exit(1)
    # Retrieve the filePath from command-line arguments
    filePath = sys.argv[1]
    savePath = filePath.replace("/", "\\")
    base, ext = os.path.splitext(savePath)
    save_path = f"{base}_Modified{ext}"

    # Initialize Excel and open workbook
    xl = DispatchEx("Excel.Application")
    xl.Visible = True
    wb = xl.Workbooks.Add(filePath)

    chartArea = wb.Sheets(1).ChartObjects("Diagram 2").Chart
    chartArea2 = wb.Sheets(1).ChartObjects("Диаграмма 1").Chart
    chartArea.HasLegend = True
    # chartArea.Legend.Font.Size = 8
    # Function to adjust the plot area position for a given chart
    def adjust_plot_area_position(chart):
        # Adjust the font size of the tick labels based on chart name
        vertical_axis = chart.Axes(2)  # Access the vertical axis

        # Check if any series in the chart has negative values
        has_negative_values = False
        for series in chart.SeriesCollection():
            if any(value < 0 for value in series.Values):
                has_negative_values = True
                break

        # Adjust the font size of the tick labels based on the presence of negative values
        if has_negative_values:
            vertical_axis.TickLabels.Font.Size = 7.2  # Change the font size to 7.1 points if negative values are present
        else:
            vertical_axis.TickLabels.Font.Size = 7  # Change the font size to 7 points if all values are positive
        vertical_axis.TickLabels.NumberFormat = "00000000"

        # Adjust the font size of the primary Y-axis title based on the minimum scale
        primary_y_axis = chart.Axes(2, 1)
        min_y_value = primary_y_axis.MinimumScale
        if min_y_value < 0:
            primary_y_axis.AxisTitle.Font.Size = 5  # Set font size to 5 if negative values are present
        else:
            primary_y_axis.AxisTitle.Font.Size = 8  # Otherwise, set font size to 8

    def adjust_legend(chart):
        chart.HasLegend = True
        legend = chart.Legend
        legend.Left = chart.PlotArea.Left + chart.PlotArea.Width + 100  # Adjust as needed
        legend.Top = chart.PlotArea.Top
        legend.Width = 90  # Set the width of the legend (adjust as needed)
        legend.Height = chart.PlotArea.Height  # Match the height of the plot area
        legend.Font.Size = 9  # Set the font size of the legend (adjust as needed)

    def adjust_vertical_axis_font(chart):
        # Access the vertical axis
        vertical_axis0 = chart.Axes(1)  # 2 represents the vertical axis
        vertical_axis = chart.Axes(2)  # 2 represents the vertical axis
        vertical_axis0.TickLabels.Font.Size = 8  # Change the font size to 8 points
        vertical_axis.TickLabels.Font.Size = 8  # Change the font size to 8 points

    for series in chartArea.SeriesCollection():
        if series.Name == "NOW-BR":
            series.Format.Line.Weight = 1.1  # Set line weight to 1.5 points
            series.Format.Fill.ForeColor.RGB = 255  # Red color
            series.AxisGroup = 2
            series.ChartType = 4
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
    # Iterate through all sheets in the workbook
        
    def set_axis_labels(chartArea, xAxisLabel, yAxisLabel):
        # Set X-axis label
        x_axis = chartArea.Axes(1, 1)
        x_axis.HasTitle = True
        x_axis.AxisTitle.Text = xAxisLabel
        x_axis.AxisTitle.Font.Size = 8  
        # Set primary Y-axis label
        primary_y_axis = chartArea.Axes(2, 1)
        primary_y_axis.HasTitle = True
        primary_y_axis.AxisTitle.Text = yAxisLabel
        primary_y_axis.AxisTitle.Font.Size = 8  


    #All Data
    # x_axis = chartArea.Axes(1, 1)
    # x_axis.HasTitle = True
    # x_axis.AxisTitle.Text = "Number evt (unit)"
    # x_axis.AxisTitle.Font.Size = 8  
    # primary_y_axis = chartArea.Axes(2, 1)
    # primary_y_axis.HasTitle = True
    # primary_y_axis.AxisTitle.Text = "Time (ms)"
    # primary_y_axis.AxisTitle.Font.Size = 8  
    # secondary_y_axis = chartArea.Axes(2, 2)
    # secondary_y_axis.HasTitle = True
    # secondary_y_axis.AxisTitle.Text = "Bitrate (kbps)"
    # secondary_y_axis.AxisTitle.Font.Size =  9
    # secondary_y_axis.TickLabels.Font.Size = 11
    # secondary_y_axis.AxisTitle.Left = secondary_y_axis.AxisTitle.Left + 10
    # secondary_y_axis.TickLabels.NumberFormat = "00000000"
        
    # All Data
    x_axis = chartArea.Axes(1, 1)
    x_axis.HasTitle = True
    x_axis.AxisTitle.Text = "Number evt (unit)"
    x_axis.AxisTitle.Font.Size = 8  

    primary_y_axis = chartArea.Axes(2, 1)
    primary_y_axis.HasTitle = True
    primary_y_axis.AxisTitle.Text = "Time (ms)"

    # Check if primary Y-axis contains negative values
    min_y_value = primary_y_axis.MinimumScale

    if min_y_value < 0:
        primary_y_axis.AxisTitle.Font.Size = 5  # Set font size to 5 if negative values are present
    else:
        primary_y_axis.AxisTitle.Font.Size = 8  # Otherwise, set font size to 8

    secondary_y_axis = chartArea.Axes(2, 2)
    secondary_y_axis.HasTitle = True
    secondary_y_axis.AxisTitle.Text = "Bitrate (kbps)"
    secondary_y_axis.AxisTitle.Font.Size = 9
    secondary_y_axis.TickLabels.Font.Size = 11
    secondary_y_axis.AxisTitle.Left = secondary_y_axis.AxisTitle.Left + 10
    secondary_y_axis.TickLabels.NumberFormat = "00000000"
    secondary_y_axis.MinimumScale = 0

    #PLOST
    primary_y_axis2 = chartArea2.Axes(2, 1)
    # Adjust plot area position
    # plot_area = chartArea2.PlotArea
    # plot_area.Left = plot_area.Left + 23.8  # 1 centimeter in points (1 point = 1/28.3464567 centimeters)
    primary_y_axis2.HasTitle = True
    primary_y_axis2.AxisTitle.Text = "Packet Lost"
    primary_y_axis2.AxisTitle.Font.Size = 8
    x_axis2 = chartArea2.Axes(1, 1)
    x_axis2.HasTitle = True
    x_axis2.AxisTitle.Text = "Number evt (unit)"
    x_axis2.AxisTitle.Font.Size = 8


    for i in range(3, 6):
        chart_name = "Diagram " + str(i)
        chartArea = wb.Sheets(1).ChartObjects(chart_name).Chart
        set_axis_labels(chartArea, "Number evt (unit)", "Time (μs)")
        

    # Chart Areas 6 to 10 - "Diagram 6" to "Diagram 10" RTT, RTT-A, RTT-SDA, RTT-STRD, RTT-LTRD
    for i in range(6, 11):
        chart_name = "Diagram " + str(i)
        chartArea = wb.Sheets(1).ChartObjects(chart_name).Chart
        set_axis_labels(chartArea, "Number evt (unit)", "Time (ms)")


    for sheet in wb.Sheets:
        for chartObject in sheet.ChartObjects():
            chart = chartObject.Chart
            adjust_legend(chart)
            adjust_plot_area_position(chart)
            # adjust_vertical_axis_font(chart)
            # The above two function is conflict each other
    wb.SaveAs(save_path, FileFormat=51)
    # Close workbook and quit Excel
    wb.Close(SaveChanges=1)
    xl.Quit()
