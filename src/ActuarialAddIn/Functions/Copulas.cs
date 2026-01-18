using ExcelDna.Integration;
using MathNet.Numerics.Distributions;
using MathNet.Numerics.LinearAlgebra;

namespace ActuarialAddIn.Functions;

/// <summary>
/// Copula functions for modeling multivariate dependence structures.
/// References:
/// - McNeil, A.J., Frey, R., and Embrechts, P. (2015). "Quantitative Risk Management: Concepts, Techniques and Tools." Princeton University Press.
/// - Joe, H. (2014). "Dependence Modeling with Copulas." CRC Press.
/// </summary>
public static class Copulas
{
    #region Gaussian Copula

    [ExcelFunction(Description = "Generate correlated uniform random numbers using Gaussian (Normal) copula. Returns values in [0,1] suitable for transforming to any marginal distribution.", Category = "Actuarial.Copulas")]
    public static object[,] ACT_COPULA_GAUSSIAN(
        [ExcelArgument(Description = "Correlation matrix (n x n). Must be symmetric and positive definite.")] double[,] correlationMatrix,
        [ExcelArgument(Description = "Number of samples to generate")] int numSamples,
        [ExcelArgument(Description = "Random seed (optional, use 0 for random)")] int seed = 0)
    {
        int n = correlationMatrix.GetLength(0);
        if (correlationMatrix.GetLength(1) != n)
            return new object[,] { { "Error: Correlation matrix must be square" } };

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

            for (int i = 0; i < numSamples; i++)
            {
                // Generate independent standard normal samples
                var z = Vector<double>.Build.Dense(n, _ => Normal.Sample(random, 0, 1));

                // Apply Cholesky factor to get correlated normals
                var correlatedNormals = L * z;

                // Convert to uniform via standard normal CDF
                for (int j = 0; j < n; j++)
                {
                    result[i, j] = Normal.CDF(0, 1, correlatedNormals[j]);
                }
            }

            return result;
        }
        catch (Exception ex)
        {
            if (ex.Message.Contains("positive definite") || ex.Message.Contains("Cholesky"))
                return new object[,] { { "Error: Correlation matrix must be positive definite" } };
            return new object[,] { { $"Error: {ex.Message}" } };
        }
    }

    [ExcelFunction(Description = "Generate a single row of correlated uniform random numbers using Gaussian copula", Category = "Actuarial.Copulas")]
    public static object[] ACT_COPULA_GAUSSIAN_SINGLE(
        [ExcelArgument(Description = "Correlation matrix (n x n). Must be symmetric and positive definite.")] double[,] correlationMatrix,
        [ExcelArgument(Description = "Random seed (optional, use 0 for random)")] int seed = 0)
    {
        var result = ACT_COPULA_GAUSSIAN(correlationMatrix, 1, seed);

        if (result[0, 0] is string)
            return new object[] { result[0, 0] };

        int n = result.GetLength(1);
        var singleRow = new object[n];
        for (int i = 0; i < n; i++)
            singleRow[i] = result[0, i];

        return singleRow;
    }

    #endregion

    #region Student-t Copula

    [ExcelFunction(Description = "Generate correlated uniform random numbers using Student-t copula. Has tail dependence (unlike Gaussian copula), controlled by degrees of freedom. Lower df = stronger tail dependence.", Category = "Actuarial.Copulas")]
    public static object[,] ACT_COPULA_STUDENT_T(
        [ExcelArgument(Description = "Correlation matrix (n x n). Must be symmetric and positive definite.")] double[,] correlationMatrix,
        [ExcelArgument(Description = "Degrees of freedom (> 0). Lower values = heavier tails and stronger tail dependence. Typical values: 3-10 for insurance applications.")] double degreesOfFreedom,
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
            if (ex.Message.Contains("positive definite") || ex.Message.Contains("Cholesky"))
                return new object[,] { { "Error: Correlation matrix must be positive definite" } };
            return new object[,] { { $"Error: {ex.Message}" } };
        }
    }

    [ExcelFunction(Description = "Generate a single row of correlated uniform random numbers using Student-t copula", Category = "Actuarial.Copulas")]
    public static object[] ACT_COPULA_STUDENT_T_SINGLE(
        [ExcelArgument(Description = "Correlation matrix (n x n). Must be symmetric and positive definite.")] double[,] correlationMatrix,
        [ExcelArgument(Description = "Degrees of freedom (> 0). Lower values = heavier tails and stronger tail dependence.")] double degreesOfFreedom,
        [ExcelArgument(Description = "Random seed (optional, use 0 for random)")] int seed = 0)
    {
        var result = ACT_COPULA_STUDENT_T(correlationMatrix, degreesOfFreedom, 1, seed);

        if (result[0, 0] is string)
            return new object[] { result[0, 0] };

        int n = result.GetLength(1);
        var singleRow = new object[n];
        for (int i = 0; i < n; i++)
            singleRow[i] = result[0, i];

        return singleRow;
    }

    #endregion
}
