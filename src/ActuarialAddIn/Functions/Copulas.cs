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

            // Chi-squared for t-distribution - use same random source for reproducibility
            var chiSquared = new ChiSquared(degreesOfFreedom, random);

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

    #region Clayton Copula

    [ExcelFunction(Description = "Generate correlated uniform random numbers using Clayton copula. Has LOWER tail dependence - good for modeling joint extreme losses. Bivariate only.", Category = "Actuarial.Copulas")]
    public static object[,] ACT_COPULA_CLAYTON(
        [ExcelArgument(Description = "Theta parameter (> 0). Higher theta = stronger dependence. Kendall's tau = θ/(θ+2).")] double theta,
        [ExcelArgument(Description = "Number of samples to generate")] int numSamples,
        [ExcelArgument(Description = "Random seed (optional, use 0 for random)")] int seed = 0)
    {
        if (theta <= 0)
            return new object[,] { { "Error: Theta must be positive for Clayton copula" } };

        if (numSamples <= 0)
            return new object[,] { { "Error: Number of samples must be positive" } };

        try
        {
            var random = seed == 0 ? new Random() : new Random(seed);
            var result = new object[numSamples, 2];

            for (int i = 0; i < numSamples; i++)
            {
                double u = random.NextDouble();
                double t = random.NextDouble();

                // Conditional inverse method for Clayton copula
                // v = ((t^(-θ/(1+θ)) - 1) * u^(-θ) + 1)^(-1/θ)
                double v = Math.Pow(
                    Math.Pow(t, -theta / (1 + theta)) * Math.Pow(u, -theta) - Math.Pow(u, -theta) + 1,
                    -1 / theta);

                result[i, 0] = u;
                result[i, 1] = Math.Max(0, Math.Min(1, v));  // Clamp to [0,1]
            }

            return result;
        }
        catch (Exception ex)
        {
            return new object[,] { { $"Error: {ex.Message}" } };
        }
    }

    [ExcelFunction(Description = "Generate a single pair of correlated uniform random numbers using Clayton copula", Category = "Actuarial.Copulas")]
    public static object[] ACT_COPULA_CLAYTON_SINGLE(
        [ExcelArgument(Description = "Theta parameter (> 0)")] double theta,
        [ExcelArgument(Description = "Random seed (optional, use 0 for random)")] int seed = 0)
    {
        var result = ACT_COPULA_CLAYTON(theta, 1, seed);

        if (result[0, 0] is string)
            return new object[] { result[0, 0] };

        return new object[] { result[0, 0], result[0, 1] };
    }

    [ExcelFunction(Description = "Clayton copula CDF: C(u,v) = (u^(-θ) + v^(-θ) - 1)^(-1/θ)", Category = "Actuarial.Copulas")]
    public static double ACT_COPULA_CLAYTON_CDF(
        [ExcelArgument(Description = "First uniform value (0 to 1)")] double u,
        [ExcelArgument(Description = "Second uniform value (0 to 1)")] double v,
        [ExcelArgument(Description = "Theta parameter (> 0)")] double theta)
    {
        if (theta <= 0) return double.NaN;
        if (u < 0 || u > 1 || v < 0 || v > 1) return double.NaN;
        if (u == 0 || v == 0) return 0;
        if (u == 1) return v;
        if (v == 1) return u;

        double val = Math.Pow(u, -theta) + Math.Pow(v, -theta) - 1;
        if (val <= 0) return 0;
        return Math.Pow(val, -1 / theta);
    }

    #endregion

    #region Frank Copula

    [ExcelFunction(Description = "Generate correlated uniform random numbers using Frank copula. Symmetric dependence with NO tail dependence. Bivariate only.", Category = "Actuarial.Copulas")]
    public static object[,] ACT_COPULA_FRANK(
        [ExcelArgument(Description = "Theta parameter (non-zero). Positive = positive dependence, negative = negative dependence. Kendall's tau = 1 - 4*(1-D₁(θ))/θ where D₁ is Debye function.")] double theta,
        [ExcelArgument(Description = "Number of samples to generate")] int numSamples,
        [ExcelArgument(Description = "Random seed (optional, use 0 for random)")] int seed = 0)
    {
        if (Math.Abs(theta) < 1e-10)
            return new object[,] { { "Error: Theta cannot be zero for Frank copula" } };

        if (numSamples <= 0)
            return new object[,] { { "Error: Number of samples must be positive" } };

        try
        {
            var random = seed == 0 ? new Random() : new Random(seed);
            var result = new object[numSamples, 2];

            double expTheta = Math.Exp(-theta);

            for (int i = 0; i < numSamples; i++)
            {
                double u = random.NextDouble();
                double t = random.NextDouble();

                // Conditional inverse method for Frank copula
                // v = -ln(1 + t*(exp(-θ) - 1) / (t + (1-t)*exp(-θ*u))) / θ
                double expThetaU = Math.Exp(-theta * u);
                double v = -Math.Log(1 + t * (expTheta - 1) / (t + (1 - t) * expThetaU)) / theta;

                result[i, 0] = u;
                result[i, 1] = Math.Max(0, Math.Min(1, v));  // Clamp to [0,1]
            }

            return result;
        }
        catch (Exception ex)
        {
            return new object[,] { { $"Error: {ex.Message}" } };
        }
    }

    [ExcelFunction(Description = "Generate a single pair of correlated uniform random numbers using Frank copula", Category = "Actuarial.Copulas")]
    public static object[] ACT_COPULA_FRANK_SINGLE(
        [ExcelArgument(Description = "Theta parameter (non-zero)")] double theta,
        [ExcelArgument(Description = "Random seed (optional, use 0 for random)")] int seed = 0)
    {
        var result = ACT_COPULA_FRANK(theta, 1, seed);

        if (result[0, 0] is string)
            return new object[] { result[0, 0] };

        return new object[] { result[0, 0], result[0, 1] };
    }

    [ExcelFunction(Description = "Frank copula CDF: C(u,v) = -ln(1 + (exp(-θu)-1)(exp(-θv)-1)/(exp(-θ)-1)) / θ", Category = "Actuarial.Copulas")]
    public static double ACT_COPULA_FRANK_CDF(
        [ExcelArgument(Description = "First uniform value (0 to 1)")] double u,
        [ExcelArgument(Description = "Second uniform value (0 to 1)")] double v,
        [ExcelArgument(Description = "Theta parameter (non-zero)")] double theta)
    {
        if (Math.Abs(theta) < 1e-10) return u * v;  // Independence as theta -> 0
        if (u < 0 || u > 1 || v < 0 || v > 1) return double.NaN;
        if (u == 0 || v == 0) return 0;
        if (u == 1) return v;
        if (v == 1) return u;

        double expThetaU = Math.Exp(-theta * u);
        double expThetaV = Math.Exp(-theta * v);
        double expTheta = Math.Exp(-theta);

        double numer = (expThetaU - 1) * (expThetaV - 1);
        double denom = expTheta - 1;

        return -Math.Log(1 + numer / denom) / theta;
    }

    #endregion

    #region Gumbel Copula

    [ExcelFunction(Description = "Generate correlated uniform random numbers using Gumbel copula. Has UPPER tail dependence - good for modeling joint extreme gains or survival. Bivariate only.", Category = "Actuarial.Copulas")]
    public static object[,] ACT_COPULA_GUMBEL(
        [ExcelArgument(Description = "Theta parameter (>= 1). Theta=1 is independence. Higher theta = stronger dependence. Kendall's tau = 1 - 1/θ.")] double theta,
        [ExcelArgument(Description = "Number of samples to generate")] int numSamples,
        [ExcelArgument(Description = "Random seed (optional, use 0 for random)")] int seed = 0)
    {
        if (theta < 1)
            return new object[,] { { "Error: Theta must be >= 1 for Gumbel copula" } };

        if (numSamples <= 0)
            return new object[,] { { "Error: Number of samples must be positive" } };

        try
        {
            var random = seed == 0 ? new Random() : new Random(seed);
            var result = new object[numSamples, 2];

            // For Gumbel copula, use Marshall-Olkin algorithm with stable distribution
            for (int i = 0; i < numSamples; i++)
            {
                if (Math.Abs(theta - 1) < 1e-10)
                {
                    // Independence case
                    result[i, 0] = random.NextDouble();
                    result[i, 1] = random.NextDouble();
                }
                else
                {
                    // Generate stable random variable with alpha = 1/theta
                    double alpha = 1.0 / theta;
                    double s = SampleStable(random, alpha);

                    // Generate two independent exponentials
                    double e1 = -Math.Log(random.NextDouble());
                    double e2 = -Math.Log(random.NextDouble());

                    // Transform to Gumbel copula
                    double u = Math.Exp(-Math.Pow(e1 / s, alpha));
                    double v = Math.Exp(-Math.Pow(e2 / s, alpha));

                    result[i, 0] = Math.Max(0, Math.Min(1, u));
                    result[i, 1] = Math.Max(0, Math.Min(1, v));
                }
            }

            return result;
        }
        catch (Exception ex)
        {
            return new object[,] { { $"Error: {ex.Message}" } };
        }
    }

    // Sample from positive stable distribution with parameter alpha
    private static double SampleStable(Random random, double alpha)
    {
        // Chambers-Mallows-Stuck algorithm for positive stable
        double u = (random.NextDouble() - 0.5) * Math.PI;
        double w = -Math.Log(random.NextDouble());

        if (Math.Abs(alpha - 1) < 1e-10)
            return Math.Tan(u);

        double zeta = -Math.Tan(Math.PI * alpha / 2);
        double xi = Math.Atan(-zeta) / alpha;

        double term1 = Math.Pow(1 + zeta * zeta, 1 / (2 * alpha));
        double term2 = Math.Sin(alpha * (u + xi)) / Math.Pow(Math.Cos(u), 1 / alpha);
        double term3 = Math.Pow(Math.Cos(u - alpha * (u + xi)) / w, (1 - alpha) / alpha);

        return term1 * term2 * term3;
    }

    [ExcelFunction(Description = "Generate a single pair of correlated uniform random numbers using Gumbel copula", Category = "Actuarial.Copulas")]
    public static object[] ACT_COPULA_GUMBEL_SINGLE(
        [ExcelArgument(Description = "Theta parameter (>= 1)")] double theta,
        [ExcelArgument(Description = "Random seed (optional, use 0 for random)")] int seed = 0)
    {
        var result = ACT_COPULA_GUMBEL(theta, 1, seed);

        if (result[0, 0] is string)
            return new object[] { result[0, 0] };

        return new object[] { result[0, 0], result[0, 1] };
    }

    [ExcelFunction(Description = "Gumbel copula CDF: C(u,v) = exp(-((-ln u)^θ + (-ln v)^θ)^(1/θ))", Category = "Actuarial.Copulas")]
    public static double ACT_COPULA_GUMBEL_CDF(
        [ExcelArgument(Description = "First uniform value (0 to 1)")] double u,
        [ExcelArgument(Description = "Second uniform value (0 to 1)")] double v,
        [ExcelArgument(Description = "Theta parameter (>= 1)")] double theta)
    {
        if (theta < 1) return double.NaN;
        if (u < 0 || u > 1 || v < 0 || v > 1) return double.NaN;
        if (u == 0 || v == 0) return 0;
        if (u == 1) return v;
        if (v == 1) return u;

        double logU = -Math.Log(u);
        double logV = -Math.Log(v);
        double inner = Math.Pow(logU, theta) + Math.Pow(logV, theta);
        return Math.Exp(-Math.Pow(inner, 1 / theta));
    }

    #endregion

    #region Copula Utilities

    [ExcelFunction(Description = "Convert Kendall's tau to copula theta parameter. Returns theta for the specified copula type.", Category = "Actuarial.Copulas")]
    public static object ACT_COPULA_TAU_TO_THETA(
        [ExcelArgument(Description = "Kendall's tau correlation (-1 to 1)")] double tau,
        [ExcelArgument(Description = "Copula type: 'CLAYTON', 'FRANK', or 'GUMBEL'")] string copulaType)
    {
        if (tau < -1 || tau > 1)
            return "Error: Tau must be between -1 and 1";

        switch (copulaType.ToUpper())
        {
            case "CLAYTON":
                if (tau <= 0)
                    return "Error: Clayton copula requires positive tau (positive dependence)";
                // tau = theta / (theta + 2), so theta = 2*tau / (1 - tau)
                return 2 * tau / (1 - tau);

            case "FRANK":
                // tau = 1 - 4*(1 - D1(theta))/theta where D1 is first Debye function
                // Numerical inversion needed - use Newton-Raphson
                return FrankTauToTheta(tau);

            case "GUMBEL":
                if (tau < 0)
                    return "Error: Gumbel copula requires non-negative tau";
                // tau = 1 - 1/theta, so theta = 1/(1 - tau)
                if (tau >= 1) return double.PositiveInfinity;
                return 1 / (1 - tau);

            default:
                return $"Error: Unknown copula type '{copulaType}'. Use CLAYTON, FRANK, or GUMBEL.";
        }
    }

    private static double FrankTauToTheta(double tau)
    {
        if (Math.Abs(tau) < 1e-10) return 0;

        // Newton-Raphson to find theta such that tau = 1 - 4*(1 - D1(theta))/theta
        double theta = tau > 0 ? 5 : -5;  // Initial guess

        for (int i = 0; i < 100; i++)
        {
            double d1 = DebyeFunction1(theta);
            double tauCalc = 1 - 4 * (1 - d1) / theta;

            if (Math.Abs(tauCalc - tau) < 1e-10)
                break;

            // Numerical derivative
            double h = 0.0001 * Math.Abs(theta);
            if (h < 1e-8) h = 1e-8;
            double d1h = DebyeFunction1(theta + h);
            double tauCalcH = 1 - 4 * (1 - d1h) / (theta + h);
            double deriv = (tauCalcH - tauCalc) / h;

            if (Math.Abs(deriv) < 1e-15) break;

            theta = theta - (tauCalc - tau) / deriv;
        }

        return theta;
    }

    private static double DebyeFunction1(double x)
    {
        // First Debye function: D1(x) = (1/x) * integral from 0 to x of t/(exp(t)-1) dt
        if (Math.Abs(x) < 1e-10) return 1;

        // Numerical integration using Simpson's rule
        int n = 1000;
        double h = x / n;
        double sum = 0;

        for (int i = 1; i < n; i++)
        {
            double t = i * h;
            double f = t / (Math.Exp(t) - 1);
            sum += (i % 2 == 0 ? 2 : 4) * f;
        }

        // Handle t=0 limit (f -> 1) and t=x
        double fx = x / (Math.Exp(x) - 1);
        sum = (1 + sum + fx) * h / 3;

        return sum / x;
    }

    [ExcelFunction(Description = "Calculate lower tail dependence coefficient for a copula.", Category = "Actuarial.Copulas")]
    public static object ACT_COPULA_TAIL_LOWER(
        [ExcelArgument(Description = "Copula type: 'GAUSSIAN', 'STUDENT_T', 'CLAYTON', 'FRANK', or 'GUMBEL'")] string copulaType,
        [ExcelArgument(Description = "Theta/correlation parameter")] double theta,
        [ExcelArgument(Description = "Degrees of freedom (for Student-t only)")] double df = 0)
    {
        switch (copulaType.ToUpper())
        {
            case "GAUSSIAN":
                return 0.0;  // No tail dependence
            case "STUDENT_T":
                if (df <= 0) return "Error: Degrees of freedom required for Student-t";
                if (theta <= -1 || theta >= 1) return "Error: Correlation must be in (-1,1)";
                // lambda_L = 2 * T_{df+1}(-sqrt((df+1)(1-rho)/(1+rho)))
                double arg = -Math.Sqrt((df + 1) * (1 - theta) / (1 + theta));
                return 2 * StudentT.CDF(0, 1, df + 1, arg);
            case "CLAYTON":
                if (theta <= 0) return "Error: Theta must be positive for Clayton";
                return Math.Pow(2, -1 / theta);  // Lower tail dependence
            case "FRANK":
                return 0.0;  // No tail dependence
            case "GUMBEL":
                return 0.0;  // No lower tail dependence (only upper)
            default:
                return $"Error: Unknown copula type '{copulaType}'";
        }
    }

    [ExcelFunction(Description = "Calculate upper tail dependence coefficient for a copula.", Category = "Actuarial.Copulas")]
    public static object ACT_COPULA_TAIL_UPPER(
        [ExcelArgument(Description = "Copula type: 'GAUSSIAN', 'STUDENT_T', 'CLAYTON', 'FRANK', or 'GUMBEL'")] string copulaType,
        [ExcelArgument(Description = "Theta/correlation parameter")] double theta,
        [ExcelArgument(Description = "Degrees of freedom (for Student-t only)")] double df = 0)
    {
        switch (copulaType.ToUpper())
        {
            case "GAUSSIAN":
                return 0.0;  // No tail dependence
            case "STUDENT_T":
                if (df <= 0) return "Error: Degrees of freedom required for Student-t";
                if (theta <= -1 || theta >= 1) return "Error: Correlation must be in (-1,1)";
                // lambda_U = lambda_L for Student-t (symmetric)
                double arg = -Math.Sqrt((df + 1) * (1 - theta) / (1 + theta));
                return 2 * StudentT.CDF(0, 1, df + 1, arg);
            case "CLAYTON":
                return 0.0;  // No upper tail dependence (only lower)
            case "FRANK":
                return 0.0;  // No tail dependence
            case "GUMBEL":
                if (theta < 1) return "Error: Theta must be >= 1 for Gumbel";
                return 2 - Math.Pow(2, 1 / theta);  // Upper tail dependence
            default:
                return $"Error: Unknown copula type '{copulaType}'";
        }
    }

    #endregion
}
