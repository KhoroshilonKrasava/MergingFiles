using System;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.Linq;

class ExcelMerger
{
    [STAThread]
    static void Main()
    {
        // explorer
        using (OpenFileDialog openFileDialog = new OpenFileDialog())
        {
            openFileDialog.Filter = "Excel Files|*.xlsx";
            openFileDialog.Multiselect = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                var combinedWorkbook = new XLWorkbook();
                int sheetCounter = 1;

                foreach (string filePath in openFileDialog.FileNames)
                {
                    // book Excel
                    var workbook = new XLWorkbook(filePath);
                    var worksheet = workbook.Worksheets.First();
                    string sheetName = "ДО" + sheetCounter++;
                    // add new list
                    var newWorksheet = combinedWorkbook.Worksheets.Add(sheetName);
                    worksheet.RangeUsed().CopyTo(newWorksheet.Cell(1, 1));
                }

                // save new Excel
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel Files|*.xlsx";
                    saveFileDialog.Title = "Сохранить объединенный Excel файл";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        combinedWorkbook.SaveAs(saveFileDialog.FileName);
                        Console.WriteLine($"Объединенный файл сохранен: {saveFileDialog.FileName}");
                    }
                }
            }
        }
    }
}
