using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.Data;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Drawing;


namespace PMMP
{
    public static class WorkbookUtilities
    {


        public static void ReplicateRow(SheetData sheetData, int refRowIndex, int count)
        {
            Repository.Utility.WriteLog("ReplicateRow started", System.Diagnostics.EventLogEntryType.Information);
            IEnumerable<Row> rows = sheetData.Descendants<Row>().Where(r => r.RowIndex.Value > refRowIndex);

            foreach (Row row in rows)
                IncrementIndexes(row, count);

            Row refRow = GetRow(sheetData, refRowIndex);

            for (int i = 0; i < count; i++)
            {
                Row newRow = (Row)refRow.Clone();
                IncrementIndexes(newRow, i + 1);
                //sheetData.InsertAt(newRow, i + 1);
                sheetData.InsertAfter(newRow, GetRow(sheetData, refRowIndex + i));
            }
            Repository.Utility.WriteLog("ReplicateRow completed successfully", System.Diagnostics.EventLogEntryType.Information);
        }

        public static void LoadSheetData(SheetData sheetData, DataTable data, int rowIndex, int columnindex)
        {
            //Populate data
            Repository.Utility.WriteLog("LoadSheetData started", System.Diagnostics.EventLogEntryType.Information);


            int startRow = rowIndex + 1;
            for (int i = 0; i < data.Rows.Count; i++)
            {
                Row row = GetRow(sheetData, i + startRow);
                if (row == null)
                {
                    row = CreateContentRow(data.Rows[i], i + startRow, columnindex);
                    sheetData.AppendChild(row);
                }
                else
                {
                    PopulateRow(data.Rows[i], i + 2, row);
                }
            }
            //  // Position the chart on the worksheet using a TwoCellAnchor object.
            //    drawingsPart.WorksheetDrawing = new WorksheetDrawing();
            //    TwoCellAnchor twoCellAnchor = drawingsPart.WorksheetDrawing.AppendChild<TwoCellAnchor>(new TwoCellAnchor());
            //    twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(new ColumnId("1"),
            //        new RowId("2")
            //        ));
            //    twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(new ColumnId("5"),
            //        new RowId(data.Rows.Count.ToString())
            //        ));

            //    twoCellAnchor.Append(new ClientData());

            //    // Save the WorksheetDrawing object.
            //    drawingsPart.WorksheetDrawing.Save();
            Repository.Utility.WriteLog("LoadSheetData completed successfully", System.Diagnostics.EventLogEntryType.Information);
        }

        public static Row GetRow(SheetData sheetData, int rowIndex)
        {
            try
            {
                return sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            catch
            {
                return null;
            }
        }

        public static Cell GetCell(Row row, int columnIndex)
        {
            return row.Elements<Cell>().FirstOrDefault(c => string.Compare(c.CellReference.Value, GetColumnName(columnIndex) + row.RowIndex, true) == 0);
        }

        private static string GetColumnName(int columnIndex)
        {
            int dividend = columnIndex;
            string columnName = String.Empty;
            int modifier;

            while (dividend > 0)
            {
                modifier = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modifier).ToString() + columnName;
                dividend = (int)((dividend - modifier) / 26);
            }

            return columnName;
        }

        private static Row CreateContentRow(DataRow dataRow, int rowIndex, int columnindex)
        {
            Row row = new Row { RowIndex = (UInt32)rowIndex };

            PopulateRow(dataRow, rowIndex + 2, row);




            return row;
        }

        private static void PopulateRow(DataRow dataRow, int rowindex, Row row)
        {
            Repository.Utility.WriteLog("PopulateRow started", System.Diagnostics.EventLogEntryType.Information);
            Cell dataCell = GetCell(row, 1);
            if (dataCell.DataType != null && dataCell.DataType == CellValues.SharedString)
                dataCell.DataType = CellValues.String;
            //dataCell.CellValue.Text = dataRow["Task"].ToString().Split(":".ToCharArray())[0] + " " + ((DateTime)dataRow["Finish"]).ToShortDateString();
            //dataCell.CellFormula = new CellFormula(string.Format("=CONCATENATE(I{0},\":  \",TEXT(H{1},\"m/d\"))", rowindex, rowindex));

            //dataCell = GetCell(row, 4);
            //if (dataCell.DataType != null && dataCell.DataType == CellValues.SharedString)
            //    dataCell.DataType = CellValues.String;
            //if (dataRow["Task"] != System.DBNull.Value)
            //dataCell.CellValue.Text = dataRow["Task"].ToString().Split(":".ToCharArray())[0];

            dataCell = GetCell(row, 2);
            if (dataCell.DataType != null && dataCell.DataType == CellValues.SharedString)
                dataCell.DataType = CellValues.String;
            if (dataRow["Start"] != System.DBNull.Value && !string.IsNullOrEmpty(dataRow["Start"].ToString()))
                dataCell.CellValue.Text = ((DateTime)dataRow["Start"]).ToOADate().ToString();

            dataCell = GetCell(row, 3);
            if (dataCell.DataType != null && dataCell.DataType == CellValues.SharedString)
                dataCell.DataType = CellValues.String;
            //dataCell.CellFormula = new CellFormula(string.Format("=IF(F{0}-B{1} > 4,F{2}-B{3},5)", rowindex, rowindex, rowindex, rowindex));
            //if (dataRow["Duration"] != System.DBNull.Value)
            //    dataCell.CellValue.Text = Convert.ToInt32(dataRow["Duration"].ToString()) != 0 ? (Convert.ToInt32(dataRow["Duration"].ToString()) / 4800).ToString() : "0";
            //else
            //    dataCell.CellValue.Text = "";



            dataCell = GetCell(row, 4);
            if (dataCell.DataType != null && dataCell.DataType == CellValues.SharedString)
                dataCell.DataType = CellValues.String;
            else
                dataCell.CellValue.Text = "";

            if (dataRow["BaseLineStart"] != System.DBNull.Value && !string.IsNullOrEmpty(dataRow["BaseLineStart"].ToString()))
                dataCell.CellValue.Text = ((DateTime)dataRow["BaseLineStart"]).ToOADate().ToString();
            else
                dataCell.CellValue.Text = "";
            dataCell = GetCell(row, 5);
            if (dataCell.DataType != null && dataCell.DataType == CellValues.SharedString)
                dataCell.DataType = CellValues.String;
            //dataCell.CellFormula = new CellFormula(string.Format("=IF(H{0}-D{1} > 4,H{2}-D{3},5)", rowindex, rowindex, rowindex, rowindex));
            //else
            //    dataCell.CellValue.Text = "";
            //if (dataRow["BLDuration"] != System.DBNull.Value)
            //dataCell.CellValue.Text = Convert.ToInt32(dataRow["BLDuration"].ToString()) != 0 ? (Convert.ToInt32(dataRow["BLDuration"].ToString()) / 4800).ToString() : "0";
            //else
            //    dataCell.CellValue.Text = "";
            dataCell = GetCell(row, 6);
            if (dataCell.DataType != null && dataCell.DataType == CellValues.SharedString)
                dataCell.DataType = CellValues.String;
            else
                dataCell.CellValue.Text = "";

            if (dataRow["Finish"] != System.DBNull.Value && !string.IsNullOrEmpty(dataRow["Finish"].ToString()))
                dataCell.CellValue.Text = ((DateTime)dataRow["Finish"]).ToOADate().ToString();
            else
                dataCell.CellValue.Text = "";

            dataCell = GetCell(row, 7);
            if (dataCell.DataType != null && dataCell.DataType == CellValues.SharedString)
                dataCell.DataType = CellValues.String;
            else
                dataCell.CellValue.Text = "";

            if (dataRow["ID"] != System.DBNull.Value && !string.IsNullOrEmpty(dataRow["ID"].ToString()))
                dataCell.CellValue.Text = dataRow["ID"].ToString();
            else
                dataCell.CellValue.Text = "";
            dataCell = GetCell(row, 8);
            if (dataCell.DataType != null && dataCell.DataType == CellValues.SharedString)
                dataCell.DataType = CellValues.String;
            else
                dataCell.CellValue.Text = "";

            if (dataRow["BaseLineFinish"] != System.DBNull.Value && !string.IsNullOrEmpty(dataRow["BaseLineFinish"].ToString()))
                dataCell.CellValue.Text = ((DateTime)dataRow["BaseLineFinish"]).ToOADate().ToString();
            else
                dataCell.CellValue.Text = "";
            dataCell = GetCell(row, 9);
            if (dataCell.DataType != null && dataCell.DataType == CellValues.SharedString)
                dataCell.DataType = CellValues.String;
            else
                dataCell.CellValue.Text = "";

            if (dataRow["Task"] != System.DBNull.Value && !string.IsNullOrEmpty(dataRow["Task"].ToString()))
                dataCell.CellValue.Text = dataRow["Task"].ToString();
            else
                dataCell.CellValue.Text = "";



            Repository.Utility.WriteLog("PopulateRow complete successfully", System.Diagnostics.EventLogEntryType.Information);
        }

        private static Cell CreateCell(int columnIndex, int rowIndex, object cellValue)
        {
            Repository.Utility.WriteLog("CreateCell started", System.Diagnostics.EventLogEntryType.Information);
            Cell cell = new Cell();

            cell.CellReference = GetColumnName(columnIndex) + rowIndex;

            var value = cellValue.ToString();

            Decimal number;
            if (cellValue.GetType() == typeof(Decimal) || Decimal.TryParse(value, out number))
            {
                cell.DataType = CellValues.Number;
            }
            else if (cellValue.GetType() == typeof(DBNull))
            {
                cell.DataType = CellValues.String;
                value = "NULL";
            }
            else if (cellValue.GetType() == typeof(DateTime))
            {
                cell.StyleIndex = (UInt32Value)12U;
                value = (cellValue as DateTime?).Value.ToOADate().ToString();
            }
            else if (cellValue.GetType() == typeof(Boolean))
            {
                value = ((bool)cellValue) ? "1" : "0";
            }
            else
            {
                cell.DataType = CellValues.String;
            }

            cell.CellValue = new CellValue(value);
            Repository.Utility.WriteLog("CreateCell completed successfully", System.Diagnostics.EventLogEntryType.Information);
            return cell;
        }

        private static void IncrementIndexes(Row row, int increment)
        {
            Repository.Utility.WriteLog("IncrementIndexes started", System.Diagnostics.EventLogEntryType.Information);
            uint newRowIndex;
            newRowIndex = System.Convert.ToUInt32(row.RowIndex.Value + increment);

            foreach (Cell cell in row.Elements<Cell>())
            {

                string cellReference = cell.CellReference.Value;
                cell.CellReference = new StringValue(cellReference.Replace(row.RowIndex.Value.ToString(), newRowIndex.ToString()));
                string cellFormula = cell.CellFormula != null ? cell.CellFormula.Text : "";
                if (cellFormula != "")
                {
                    cell.CellValue.Remove();
                    cell.CellFormula = new CellFormula(string.Format("={0}", cellFormula.Replace(row.RowIndex.Value.ToString(), newRowIndex.ToString()), newRowIndex));
                }
            }

            row.RowIndex = new UInt32Value(newRowIndex);
            Repository.Utility.WriteLog("IncrementIndexes completed successfully", System.Diagnostics.EventLogEntryType.Information);
        }

        internal static void LoadGraphSheetData(SheetData sheetData, List<GraphDataGroup> data, int rowIndex, int columnIndex)
        {
            //Populate data
            Repository.Utility.WriteLog("LoadSheetData started", System.Diagnostics.EventLogEntryType.Information);
            int startRow = rowIndex + 1;
            for (int i = 0; i < data.Count; i++)
            {
                for (int j = 0; j < data[i].Data.Count; j++)
                {
                    Row row = GetRow(sheetData, j + startRow);
                    if (row == null)
                    {
                        row = CreateContentRow(data[i], j + startRow, columnIndex);
                        sheetData.AppendChild(row);
                    }
                    else
                        PopulateRow(data[i].Data[j], row, data[i].Type);
                }
            }
            Repository.Utility.WriteLog("LoadSheetData completed successfully", System.Diagnostics.EventLogEntryType.Information);
        }

        private static Row CreateContentRow(GraphDataGroup graphDataGroup, int rowIndex, int columnIndex)
        {
            Row row = new Row { RowIndex = (UInt32)rowIndex };

            foreach (GraphData data in graphDataGroup.Data)
            {
                PopulateRow(data, row, graphDataGroup.Type);
            }

            return row;
        }

        private static void PopulateRow(GraphData graphData, Row row, string type)
        {
            Repository.Utility.WriteLog("PopulateRow started", System.Diagnostics.EventLogEntryType.Information);

            Cell dataCell = GetCell(row, 1);
            if (dataCell.DataType != null && dataCell.DataType == CellValues.SharedString)
                dataCell.DataType = CellValues.String;
            dataCell.CellValue.Text = graphData.Title.ToString();
            switch (type)
            {
                case "CF":
                case "BES":
                case "CS":
                    Cell dataCell1 = GetCell(row, 2);
                    if (dataCell1.DataType != null && dataCell1.DataType == CellValues.SharedString)
                        dataCell1.DataType = CellValues.String;
                    dataCell1.CellValue.Text = graphData.Count.ToString();
                    break;
                case "BEF":
                case "FCF":
                case "FCS":
                    Cell dataCell2 = GetCell(row, 3);
                    if (dataCell2.DataType != null && dataCell2.DataType == CellValues.SharedString)
                        dataCell2.DataType = CellValues.String;
                    dataCell2.CellValue.Text = graphData.Count.ToString();
                    break;
                case "BEFS":
                case "DQF":
                case "DQ":
                    Cell dataCell3 = GetCell(row, 4);
                    if (dataCell3.DataType != null && dataCell3.DataType == CellValues.SharedString)
                        dataCell3.DataType = CellValues.String;
                    dataCell3.CellValue.Text = graphData.Count.ToString();
                    break;
                case "BEFF":
                case "FDQF":
                case "FDQ":
                    Cell dataCell4 = GetCell(row, 5);
                    if (dataCell4.DataType != null && dataCell4.DataType == CellValues.SharedString)
                        dataCell4.DataType = CellValues.String;
                    dataCell4.CellValue.Text = graphData.Count.ToString();
                    break;
                case "CDQF":
                case "CDQ":
                    Cell dataCell5 = GetCell(row, 6);
                    if (dataCell5.DataType != null && dataCell5.DataType == CellValues.SharedString)
                        dataCell5.DataType = CellValues.String;
                    dataCell5.CellValue.Text = graphData.Count.ToString();
                    break;
                case "FCDQF":
                case "FCDQ":
                    Cell dataCell6 = GetCell(row, 7);
                    if (dataCell6.DataType != null && dataCell6.DataType == CellValues.SharedString)
                        dataCell6.DataType = CellValues.String;
                    dataCell6.CellValue.Text = graphData.Count.ToString();
                    break;
            }
            Repository.Utility.WriteLog("PopulateRow complete successfully", System.Diagnostics.EventLogEntryType.Information);
        }
    }
}
