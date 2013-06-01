using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System.Data;

namespace PMMP
{
    public static class WorkbookUtilities
    {
        

        public static void ReplicateRow(SheetData sheetData, int refRowIndex, int count)
        {
            IEnumerable<Row> rows = sheetData.Descendants<Row>().Where(r => r.RowIndex.Value > refRowIndex);

            foreach (Row row in rows)
                IncrementIndexes(row, count);

            Row refRow = GetRow(sheetData, refRowIndex);

            for (int i = 0; i < count; i++)
            {
                Row newRow = (Row)refRow.Clone();
                IncrementIndexes(newRow, i + 1);

                sheetData.InsertAfter(newRow, GetRow(sheetData, refRowIndex + i));
            }
        }

        public static void LoadSheetData(SheetData sheetData, DataTable data, int rowIndex, int columnindex)
        {
            //Populate data
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
                    PopulateRow(data.Rows[i], i + 2, row);
            }
        }

        private static Row GetRow(SheetData sheetData, int rowIndex)
        {
            return sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        }

        private static Cell GetCell(Row row, int columnIndex)
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

            Cell dataCell = GetCell(row,1);
            if (dataCell.DataType != null && dataCell.DataType == CellValues.SharedString)
                dataCell.DataType = CellValues.String;
            dataCell.CellValue.Text = dataRow["Task"].ToString() + " " + ((DateTime)dataRow["Finish"]).ToShortDateString();
            dataCell.CellFormula = new CellFormula(string.Format("=CONCATENATE(D{0},\":  \",TEXT(B{1},\"m/d\"))",rowindex,rowindex));
            dataCell = GetCell(row, 2);
            if (dataCell.DataType != null && dataCell.DataType == CellValues.SharedString)
                dataCell.DataType = CellValues.String;
            dataCell.CellValue.Text = ((DateTime)dataRow["Finish"]).ToOADate().ToString();

            dataCell = GetCell(row, 3);
            if (dataCell.DataType != null && dataCell.DataType == CellValues.SharedString)
                dataCell.DataType = CellValues.String;
            dataCell.CellValue.Text = 10.ToString();

            dataCell = GetCell(row, 4);
            if (dataCell.DataType != null && dataCell.DataType == CellValues.SharedString)
                dataCell.DataType = CellValues.String;
            dataCell.CellValue.Text = dataRow["Task"].ToString();
                
            //int rowIndex = (int)row.RowIndex.Value;
            //for (int i = 0; i < dataRow.Table.Columns.Count; i++)
            //{
            //    int index = i + columnindex + 1;
            //    Cell dataCell = GetCell(row, index);
            //    if (dataCell == null)
            //    {
            //        dataCell = CreateCell(i + columnindex + 1, rowIndex, dataRow[i]);
            //        row.AppendChild(dataCell);
            //    }
            //    else
            //    {
            //        if (dataCell.DataType != null && dataCell.DataType == CellValues.SharedString)
            //            dataCell.DataType = CellValues.String;
            //        if (dataRow[i].GetType() == typeof(DateTime))
            //            dataCell.CellValue.Text = ((DateTime)dataRow[i]).ToOADate().ToString();
            //        else
            //            dataCell.CellValue.Text = dataRow[i].ToString();
            //    }
            //}
        }

        private static Cell CreateCell(int columnIndex, int rowIndex, object cellValue)
        {
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

            return cell;
        }

        private static void IncrementIndexes(Row row, int increment)
        {
            uint newRowIndex;
            newRowIndex = System.Convert.ToUInt32(row.RowIndex.Value + increment);

            foreach (Cell cell in row.Elements<Cell>())
            {
                string cellReference = cell.CellReference.Value;
                cell.CellReference = new StringValue(cellReference.Replace(row.RowIndex.Value.ToString(), newRowIndex.ToString()));
            }

            row.RowIndex = new UInt32Value(newRowIndex);
        }
    }
}
