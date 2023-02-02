import os
import win32com.client as win32

excel = win32.Dispatch("Excel.Application")
excel.Visible = True

workbook = excel.Workbooks.Add()
workbook.SaveAs(os.path.join(os.getcwd(), "myfile.xlsx"))

sheet1 = workbook.Worksheets("Sheet1")
sheet1.name = "ToDo List"
sheet1.Range("A:D").ColumnWidth = 30

cells = sheet1.Cells

cells(1, "A").Value = "Task"
cells(1, "A").Font.Bold = True
cells(1, "B").Value = "Description"
cells(1, "B").Font.Bold = True
cells(1, "C").Value = "Done"
cells(1, "C").Font.Bold = True
cells(1, "D").Value = "Time Needed"
cells(1, "D").Font.Bold = True


cells(2, "A").Value = "Python"
cells(2, "B").Value = "Write Python script"
cells(2, "C").Value = ""
cells(2, "D").Value = "3"

cells(3, "A").Value = "Dinner"
cells(3, "B").Value = "Cook dinner"
cells(3, "C").Value = ""
cells(3, "D").Value = "1"


cells(4, "A").Value = "App"
cells(4, "B").Value = "Finish debugging app"
cells(4, "C").Value = ""
cells(4, "D").Value = "4"

cells(5, "D").Value = "=SUM(D2:D4)"

ch = sheet1.Shapes.AddChart().Select()

excel.ActiveChart.SetSourceData(Source=sheet1.Range("D2:D4"), PlotBy=2)
excel.ActiveChart.ChartType = 5

workbook.SaveAs(os.path.join(os.getcwd(), "myfile.xlsx"))