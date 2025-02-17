using Aspose.Cells;
using System;
using System.Globalization;

namespace TemplateSpeAsposeCLI.Utils
{
    public static class UtilsCell
    {
        public static string GetCellStringValue(Worksheet worksheet, int rowIndex, int colIndex)
        {
            Cell cell = worksheet.Cells[rowIndex, colIndex];
            string cellValue = "";
            if (cell != null) cellValue = cell.StringValue;
            return cellValue;
        }

        public static int getCellIntValue(Worksheet worksheet, int rowIndex, int colIndex)
        {
            string cellValue = GetCellStringValue(worksheet, rowIndex, colIndex);
            try
            {
                int result = Int32.Parse(cellValue);
                return result;
            }
            catch (FormatException)
            {
                return 0;
            }
        }

        public static double getCellDoubleValue(Worksheet worksheet, int rowIndex, int colIndex)
        {
            string cellValue = GetCellStringValue(worksheet, rowIndex, colIndex);
            if (!string.IsNullOrEmpty(cellValue))
            {
                try
                {
                    return double.Parse(cellValue);
                }
                catch (FormatException)
                {
                    try
                    {
                        var pointCulture = new CultureInfo("en")
                        {
                            NumberFormat =
                        {
                            NumberDecimalSeparator = "."
                        }
                        };
                        double value = double.Parse(cellValue, NumberStyles.Float, pointCulture);
                        return value;
                    }
                    catch (Exception)
                    {
                        return 0;
                    }
                }
            }
            else return 0;

        }
        /// <summary>
        /// Find cell by its value
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="value"></param>
        /// <param name="previousCell"></param>
        /// <returns> foun cell </returns>
        public static Cell FindCellByValue(Worksheet worksheet, string value, Cell previousCell = null)
        {
            FindOptions opts = new FindOptions
            {
                LookInType = LookInType.Values,
                LookAtType = LookAtType.EntireContent
            };
            Cell cell = worksheet.Cells.Find(value, previousCell, opts);
            return cell;
        }
    }
}