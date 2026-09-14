using System;
using System.Collections.Generic;

namespace MonteCarlo.Excel
{
    public static class ExcelDataReader
    {
        // =========================================================
        // READ A SPECIFIC EXCEL RANGE
        // =========================================================

        public static double[] ReadNumericDataFromRange(
            dynamic range,
            out string rangeAddress)
        {
            if (range == null)
            {
                throw new Exception(
                    "No Excel range was selected.");
            }


            dynamic worksheet =
                range.Worksheet;


            string sheetName =
                Convert.ToString(
                    worksheet.Name)
                ?? "";


            string address =
                range.Address[
                    false,
                    false];


            rangeAddress =
                $"{sheetName}!{address}";


            object values =
                range.Value2;


            List<double> numericValues =
                new List<double>();


            // -----------------------------------------------------
            // SINGLE CELL
            // -----------------------------------------------------

            if (range.Cells.Count == 1)
            {
                TryAddNumericValue(
                    numericValues,
                    values);
            }


            // -----------------------------------------------------
            // MULTIPLE CELLS
            // -----------------------------------------------------

            else if (
                values is object[,]
                valueArray)
            {
                int rowLower =
                    valueArray.GetLowerBound(0);

                int rowUpper =
                    valueArray.GetUpperBound(0);

                int columnLower =
                    valueArray.GetLowerBound(1);

                int columnUpper =
                    valueArray.GetUpperBound(1);


                for (
                    int row = rowLower;
                    row <= rowUpper;
                    row++)
                {
                    for (
                        int column = columnLower;
                        column <= columnUpper;
                        column++)
                    {
                        object cellValue =
                            valueArray[
                                row,
                                column];


                        TryAddNumericValue(
                            numericValues,
                            cellValue);
                    }
                }
            }


            else
            {
                TryAddNumericValue(
                    numericValues,
                    values);
            }


            if (numericValues.Count < 5)
            {
                throw new Exception(
                    $"The selected range contains only " +
                    $"{numericValues.Count} valid numeric observations.\n\n" +
                    $"At least 5 are required.");
            }


            return
                numericValues.ToArray();
        }


        // =========================================================
        // ADD NUMERIC VALUE
        // =========================================================

        private static void TryAddNumericValue(
            List<double> values,
            object value)
        {
            if (value == null)
            {
                return;
            }


            if (value is double doubleValue)
            {
                if (!double.IsNaN(doubleValue) &&
                    !double.IsInfinity(doubleValue))
                {
                    values.Add(
                        doubleValue);
                }

                return;
            }


            if (value is float floatValue)
            {
                values.Add(
                    floatValue);

                return;
            }


            if (value is int intValue)
            {
                values.Add(
                    intValue);

                return;
            }


            if (value is long longValue)
            {
                values.Add(
                    longValue);

                return;
            }


            if (value is decimal decimalValue)
            {
                values.Add(
                    (double)decimalValue);

                return;
            }

            // Ignore text, blanks and Excel errors.
        }
    }
}