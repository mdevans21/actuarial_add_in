namespace ActuarialAddIn.Helpers;

/// <summary>
/// Helper functions for converting Excel data types to .NET types.
/// Excel-DNA may pass arrays as object[], object[,], or double[] depending on how they're entered.
/// </summary>
public static class ArrayHelpers
{
    /// <summary>
    /// Converts an Excel array (which may be object[], object[,], or double[]) to a double array.
    /// Handles inline arrays like {1,2,3} and cell range references.
    /// </summary>
    public static double[]? ToDoubleArray(object? input)
    {
        if (input == null)
            return null;

        // Already a double array
        if (input is double[] doubleArray)
            return doubleArray;

        // Single value
        if (input is double d)
            return new[] { d };

        // 1D object array (from inline {1,2,3} syntax)
        if (input is object[] objArray)
        {
            var result = new double[objArray.Length];
            for (int i = 0; i < objArray.Length; i++)
            {
                if (objArray[i] is double val)
                    result[i] = val;
                else if (objArray[i] is int intVal)
                    result[i] = intVal;
                else if (double.TryParse(objArray[i]?.ToString(), out double parsed))
                    result[i] = parsed;
                else
                    return null; // Invalid data
            }
            return result;
        }

        // 2D object array (from cell range or {1,2;3,4} syntax) - flatten to 1D
        if (input is object[,] obj2DArray)
        {
            int rows = obj2DArray.GetLength(0);
            int cols = obj2DArray.GetLength(1);
            var result = new double[rows * cols];
            int idx = 0;
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    var cell = obj2DArray[r, c];
                    if (cell is double val)
                        result[idx++] = val;
                    else if (cell is int intVal)
                        result[idx++] = intVal;
                    else if (double.TryParse(cell?.ToString(), out double parsed))
                        result[idx++] = parsed;
                    else
                        return null; // Invalid data
                }
            }
            return result;
        }

        // Try to parse as single value
        if (double.TryParse(input.ToString(), out double singleVal))
            return new[] { singleVal };

        return null;
    }
}
