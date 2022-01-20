using Microsoft.Office.Interop.Excel;


namespace CSReport
{
    public class ExcelReport
    {
        // Properties

        private Application xlApp;

        // Constructor

        public ExcelReport()
        {
            xlApp = new Application();
            xlApp.Visible = true;
        }


        // Public Methods

        public void Create()
        {
            var wb = CreateNewWorkbook(1);
            var ws = GetWorksheet(wb, 1);
            ValueOutputTest(ws);
        }

        // Private Methods

        private Workbook? CreateNewWorkbook(int numberOfSheets)
        {
            int originalNumberOfSheets = xlApp.SheetsInNewWorkbook;
            xlApp.SheetsInNewWorkbook = numberOfSheets;
            Workbook wb = xlApp.Workbooks.Add();
            xlApp.SheetsInNewWorkbook = originalNumberOfSheets;
            return wb;
        }

        private Worksheet? GetWorksheet(Workbook? wb, int sheetIndex)
        {
            if (wb == null) return null;
            if (sheetIndex > 0 && sheetIndex <= wb.Sheets.Count)
                return wb.Sheets[sheetIndex];
            return null;
        }

        private void ValueOutputTest(Worksheet? ws)
        {
            if (ws == null) return;
            ws.Name = "CS Test";
            // Output Option 1 -> working
            ws.Cells[1, 1].Value = "Cell R1C1";
            // Output Option 2 -> working
            ws.Range[ws.Cells[1, 2], ws.Cells[1, 2]].Value = "Cells R1C2";
        }


    }
}
