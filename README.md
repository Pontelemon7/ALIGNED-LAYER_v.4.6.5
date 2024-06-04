# ALIGNED-LAYER_v.4.6.5

   Excel.Workbook wb = excapp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            Excel.Worksheet sheet = (Excel.Worksheet)wb.ActiveSheet;
            var wb = excapp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            var sheet = (Excel.Worksheet)wb.ActiveSheet;
            MakeHeader(sheet, header, title);
            MakeCaption(header.Length + 2, sheet, reportData);
