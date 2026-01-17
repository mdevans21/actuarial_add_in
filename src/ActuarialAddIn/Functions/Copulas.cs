using ExcelDna.Integration;
using MathNet.Numerics.Distributions;
using MathNet.Numerics.LinearAlgebra;

namespace ActuarialAddIn.Functions;

public static class Copulas
{
    [ExcelFunction(Description = "Generate correlated random numbers using Student-t copula", Category = "Actuarial.Copulas")]
    public static object[,] ACT_STUDENT_T_COPULA(
        [ExcelArgument(Description = "Correlation matrix (n x n)")] double[,] correlationMatrix,
        [ExcelArgument(Description = "Degrees of freedom for the t-distribution")] double degreesOfFreedom,
        [ExcelArgument(Description = "Number of samples to generate")] int numSamples,
        [ExcelArgument(Description = "Random seed (optional, use 0 for random)")] int seed = 0)
    {
        int n = correlationMatrix.GetLength(0);
        if (correlationMatrix.GetLength(1) != n)
            return new object[,] { { "Error: Correlation matrix must be square" } };

        if (degreesOfFreedom <= 0)
            return new object[,] { { "Error: Degrees of freedom must be positive" } };

        if (numSamples <= 0)
            return new object[,] { { "Error: Number of samples must be positive" } };

        try
        {
            var random = seed == 0 ? new Random() : new Random(seed);
            var result = new object[numSamples, n];

            // Create correlation matrix
            var corrMatrix = Matrix<double>.Build.DenseOfArray(correlationMatrix);

            // Cholesky decomposition
            var cholesky = corrMatrix.Cholesky();
            var L = cholesky.Factor;

            // Chi-squared for t-distribution
            var chiSquared = new ChiSquared(degreesOfFreedom, new Random(random.Next()));

            for (int i = 0; i < numSamples; i++)
            {
                // Generate independent standard normal samples
                var z = Vector<double>.Build.Dense(n, _ => Normal.Sample(random, 0, 1));

                // Apply Cholesky factor to get correlated normals
                var correlatedNormals = L * z;

                // Generate chi-squared sample for t-distribution scaling
                double w = chiSquared.Sample();
                double scale = Math.Sqrt(degreesOfFreedom / w);

                // Convert to t-distribution and then to uniform via t CDF
                var tDist = new StudentT(0, 1, degreesOfFreedom);
                for (int j = 0; j < n; j++)
                {
                    double tValue = correlatedNormals[j] * scale;
                    result[i, j] = tDist.CumulativeDistribution(tValue);
                }
            }

            return result;
        }
        catch (Exception ex)
        {
            return new object[,] { { $"Error: {ex.Message}" } };
        }
    }

    [ExcelFunction(Description = "Generate a single row of correlated uniform random numbers using Student-t copula", Category = "Actuarial.Copulas")]
    public static object[] ACT_STUDENT_T_COPULA_SINGLE(
        [ExcelArgument(Description = "Correlation matrix (n x n)")] double[,] correlationMatrix,
        [ExcelArgument(Description = "Degrees of freedom for the t-distribution")] double degreesOfFreedom,
        [ExcelArgument(Description = "Random seed (optional, use 0 for random)")] int seed = 0)
    {
        var result = ACT_STUDENT_T_COPULA(correlationMatrix, degreesOfFreedom, 1, seed);

        if (result[0, 0] is string)
            return new object[] { result[0, 0] };

        int n = result.GetLength(1);
        var singleRow = new object[n];
        for (int i = 0; i < n; i++)
            singleRow[i] = result[0, i];

        return singleRow;
    }
}
