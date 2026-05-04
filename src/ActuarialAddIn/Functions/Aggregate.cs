using System;
using System.Linq;
using ExcelDna.Integration;
using MathNet.Numerics.Distributions;

namespace ActuarialAddIn.Functions
{
    /// <summary>
    /// Aggregate claims distribution via Panjer recursion.
    /// S = X1 + X2 + ... + XN where N ~ frequency (Poisson/NegBin) and Xi ~ severity (iid)
    ///
    /// References:
    /// - Klugman, Panjer & Willmot "Loss Models" Chapter 6
    /// - actuar R package: aggregateDist()
    /// </summary>
    public static class Aggregate
    {
        #region Discretization Methods

        /// <summary>
        /// Discretize a continuous severity distribution using rounding method.
        /// f_j = F((j+0.5)h) - F((j-0.5)h) for j > 0
        /// f_0 = F(0.5h)
        /// </summary>
        [ExcelFunction(Description = "Discretize continuous severity using rounding method. Returns PMF array.", Category = "Actuarial.Experimental")]
        public static object ACT_DISCRETIZE_EXPONENTIAL(
            [ExcelArgument(Description = "Rate parameter (1/mean)")] double rate,
            [ExcelArgument(Description = "Grid spacing h")] double h,
            [ExcelArgument(Description = "Number of grid points m")] int m)
        {
            if (rate <= 0 || h <= 0 || m <= 0) return ExcelError.ExcelErrorValue;

            var pmf = new double[m + 1];
            var dist = new Exponential(rate);

            // f_0 = F(0.5h)
            pmf[0] = dist.CumulativeDistribution(0.5 * h);

            // f_j = F((j+0.5)h) - F((j-0.5)h)
            for (int j = 1; j <= m; j++)
            {
                pmf[j] = dist.CumulativeDistribution((j + 0.5) * h) - dist.CumulativeDistribution((j - 0.5) * h);
            }

            return pmf;
        }

        /// <summary>
        /// Discretize gamma severity distribution using rounding method.
        /// </summary>
        [ExcelFunction(Description = "Discretize Gamma severity using rounding method. Returns PMF array.", Category = "Actuarial.Experimental")]
        public static object ACT_DISCRETIZE_GAMMA(
            [ExcelArgument(Description = "Shape parameter alpha")] double alpha,
            [ExcelArgument(Description = "Rate parameter beta (1/scale)")] double beta,
            [ExcelArgument(Description = "Grid spacing h")] double h,
            [ExcelArgument(Description = "Number of grid points m")] int m)
        {
            if (alpha <= 0 || beta <= 0 || h <= 0 || m <= 0) return ExcelError.ExcelErrorValue;

            var pmf = new double[m + 1];
            var dist = new Gamma(alpha, 1.0 / beta);

            pmf[0] = dist.CumulativeDistribution(0.5 * h);
            for (int j = 1; j <= m; j++)
            {
                pmf[j] = dist.CumulativeDistribution((j + 0.5) * h) - dist.CumulativeDistribution((j - 0.5) * h);
            }

            return pmf;
        }

        /// <summary>
        /// Discretize lognormal severity distribution using rounding method.
        /// </summary>
        [ExcelFunction(Description = "Discretize Lognormal severity using rounding method. Returns PMF array.", Category = "Actuarial.Experimental")]
        public static object ACT_DISCRETIZE_LOGNORMAL(
            [ExcelArgument(Description = "Mu parameter (mean of log)")] double mu,
            [ExcelArgument(Description = "Sigma parameter (std dev of log)")] double sigma,
            [ExcelArgument(Description = "Grid spacing h")] double h,
            [ExcelArgument(Description = "Number of grid points m")] int m)
        {
            if (sigma <= 0 || h <= 0 || m <= 0) return ExcelError.ExcelErrorValue;

            var pmf = new double[m + 1];
            var dist = new LogNormal(mu, sigma);

            pmf[0] = dist.CumulativeDistribution(0.5 * h);
            for (int j = 1; j <= m; j++)
            {
                pmf[j] = dist.CumulativeDistribution((j + 0.5) * h) - dist.CumulativeDistribution((j - 0.5) * h);
            }

            return pmf;
        }

        #endregion

        #region Panjer Recursion

        /// <summary>
        /// Panjer recursion for Poisson(λ) frequency.
        /// g_0 = exp(-λ(1 - f_0))
        /// g_s = (λ/s) * Σ_{j=1}^{min(s,m)} j * f_j * g_{s-j}
        /// </summary>
        [ExcelFunction(Description = "Aggregate claims PMF via Panjer recursion with Poisson(λ) frequency", Category = "Actuarial.Experimental")]
        public static object ACT_PANJER_POISSON(
            [ExcelArgument(Description = "Lambda - Poisson rate parameter")] double lambda,
            [ExcelArgument(Description = "Severity PMF array (discretized)")] double[] severity_pmf,
            [ExcelArgument(Description = "Maximum aggregate value index")] int max_s)
        {
            if (lambda <= 0 || severity_pmf == null || severity_pmf.Length == 0 || max_s <= 0)
                return ExcelError.ExcelErrorValue;

            int m = severity_pmf.Length - 1;  // Maximum severity index
            var g = new double[max_s + 1];    // Aggregate PMF

            // g_0 = exp(-λ(1 - f_0))
            g[0] = Math.Exp(-lambda * (1 - severity_pmf[0]));

            // Panjer recursion: g_s = (λ/s) * Σ_{j=1}^{min(s,m)} j * f_j * g_{s-j}
            for (int s = 1; s <= max_s; s++)
            {
                double sum = 0;
                int jMax = Math.Min(s, m);
                for (int j = 1; j <= jMax; j++)
                {
                    sum += j * severity_pmf[j] * g[s - j];
                }
                g[s] = (lambda / s) * sum;
            }

            return g;
        }

        /// <summary>
        /// Panjer recursion for Negative Binomial(r, p) frequency.
        /// For NegBin, the (a, b, 0) parameters are: a = 1-p, b = (r-1)(1-p)
        /// g_0 = (p / (1 - (1-p)(1-f_0)))^r
        /// g_s = (1/s) * Σ_{j=1}^{min(s,m)} (a + b*j/s) * j * f_j * g_{s-j}
        /// </summary>
        [ExcelFunction(Description = "Aggregate claims PMF via Panjer recursion with NegBin(r,p) frequency", Category = "Actuarial.Experimental")]
        public static object ACT_PANJER_NEGBIN(
            [ExcelArgument(Description = "r - number of successes")] double r,
            [ExcelArgument(Description = "p - probability of success")] double p,
            [ExcelArgument(Description = "Severity PMF array (discretized)")] double[] severity_pmf,
            [ExcelArgument(Description = "Maximum aggregate value index")] int max_s)
        {
            if (r <= 0 || p <= 0 || p > 1 || severity_pmf == null || severity_pmf.Length == 0 || max_s <= 0)
                return ExcelError.ExcelErrorValue;

            int m = severity_pmf.Length - 1;
            var g = new double[max_s + 1];

            // Panjer (a, b) parameters for Negative Binomial
            double a = 1 - p;
            double b = (r - 1) * (1 - p);

            // g_0 = (p / (1 - (1-p)*f_0))^r
            double denominator = 1 - (1 - p) * severity_pmf[0];
            g[0] = Math.Pow(p / denominator, r);

            // Panjer recursion for (a, b, 0) class
            // g_s = (1/(1 - a*f_0)) * Σ_{j=1}^{min(s,m)} (a + b*j/s) * f_j * g_{s-j}
            double coeff = 1.0 / (1 - a * severity_pmf[0]);

            for (int s = 1; s <= max_s; s++)
            {
                double sum = 0;
                int jMax = Math.Min(s, m);
                for (int j = 1; j <= jMax; j++)
                {
                    sum += (a + b * (double)j / s) * severity_pmf[j] * g[s - j];
                }
                g[s] = coeff * sum;
            }

            return g;
        }

        /// <summary>
        /// Panjer recursion for Binomial(n, p) frequency.
        /// For Binomial, the (a, b, 0) parameters are: a = -p/(1-p), b = (n+1)p/(1-p)
        /// </summary>
        [ExcelFunction(Description = "Aggregate claims PMF via Panjer recursion with Binomial(n,p) frequency", Category = "Actuarial.Experimental")]
        public static object ACT_PANJER_BINOMIAL(
            [ExcelArgument(Description = "n - number of trials")] int n,
            [ExcelArgument(Description = "p - probability of success")] double p,
            [ExcelArgument(Description = "Severity PMF array (discretized)")] double[] severity_pmf,
            [ExcelArgument(Description = "Maximum aggregate value index")] int max_s)
        {
            if (n <= 0 || p <= 0 || p >= 1 || severity_pmf == null || severity_pmf.Length == 0 || max_s <= 0)
                return ExcelError.ExcelErrorValue;

            int m = severity_pmf.Length - 1;
            var g = new double[max_s + 1];

            // Panjer (a, b) parameters for Binomial
            double a = -p / (1 - p);
            double b = (n + 1) * p / (1 - p);

            // g_0 = ((1-p) + p*f_0)^n
            g[0] = Math.Pow((1 - p) + p * severity_pmf[0], n);

            // Panjer recursion
            double coeff = 1.0 / (1 - a * severity_pmf[0]);

            for (int s = 1; s <= max_s; s++)
            {
                double sum = 0;
                int jMax = Math.Min(s, m);
                for (int j = 1; j <= jMax; j++)
                {
                    sum += (a + b * (double)j / s) * severity_pmf[j] * g[s - j];
                }
                g[s] = coeff * sum;
            }

            return g;
        }

        #endregion

        #region Aggregate Distribution Functions

        /// <summary>
        /// Compute CDF from aggregate PMF.
        /// </summary>
        [ExcelFunction(Description = "Aggregate claims CDF at value x", Category = "Actuarial.Experimental")]
        public static double ACT_AGGREGATE_CDF(
            [ExcelArgument(Description = "Value x")] double x,
            [ExcelArgument(Description = "Aggregate PMF array")] double[] aggregate_pmf,
            [ExcelArgument(Description = "Grid spacing h")] double h)
        {
            if (aggregate_pmf == null || h <= 0) return double.NaN;
            if (x < 0) return 0;

            int index = (int)Math.Floor(x / h);
            if (index >= aggregate_pmf.Length) index = aggregate_pmf.Length - 1;

            double cdf = 0;
            for (int i = 0; i <= index; i++)
            {
                cdf += aggregate_pmf[i];
            }
            return cdf;
        }

        /// <summary>
        /// Compute VaR (quantile) from aggregate PMF.
        /// </summary>
        [ExcelFunction(Description = "Aggregate claims VaR (quantile) at probability p", Category = "Actuarial.Experimental")]
        public static double ACT_AGGREGATE_VAR(
            [ExcelArgument(Description = "Probability p (0 to 1)")] double p,
            [ExcelArgument(Description = "Aggregate PMF array")] double[] aggregate_pmf,
            [ExcelArgument(Description = "Grid spacing h")] double h)
        {
            if (aggregate_pmf == null || h <= 0 || p < 0 || p > 1) return double.NaN;

            double cdf = 0;
            for (int i = 0; i < aggregate_pmf.Length; i++)
            {
                cdf += aggregate_pmf[i];
                if (cdf >= p) return i * h;
            }
            return (aggregate_pmf.Length - 1) * h;
        }

        /// <summary>
        /// Compute TVaR (tail value at risk) from aggregate PMF.
        /// TVaR(p) = E[X | X > VaR(p)]
        /// </summary>
        [ExcelFunction(Description = "Aggregate claims TVaR at probability p", Category = "Actuarial.Experimental")]
        public static double ACT_AGGREGATE_TVAR(
            [ExcelArgument(Description = "Probability p (0 to 1)")] double p,
            [ExcelArgument(Description = "Aggregate PMF array")] double[] aggregate_pmf,
            [ExcelArgument(Description = "Grid spacing h")] double h)
        {
            if (aggregate_pmf == null || h <= 0 || p < 0 || p >= 1) return double.NaN;

            // VaR index = smallest i with cdf[i] >= p.
            double cdf = 0;
            int varIndex = aggregate_pmf.Length - 1;
            for (int i = 0; i < aggregate_pmf.Length; i++)
            {
                cdf += aggregate_pmf[i];
                if (cdf >= p)
                {
                    varIndex = i;
                    break;
                }
            }

            // CTE_p = E[X | X >= VaR_p] = (Σ_{i ≥ varIndex} (i·h)·pmf[i]) / P(X ≥ VaR_p).
            // The denominator must be the EMPIRICAL tail mass (sum of pmf from
            // varIndex onward), not the requested 1 − p — the discretisation
            // can leave the cdf at varIndex slightly above p, biasing the
            // estimate if we used 1 − p.
            double tailProb = 0;
            double tailSum = 0;
            for (int i = varIndex; i < aggregate_pmf.Length; i++)
            {
                tailProb += aggregate_pmf[i];
                tailSum += i * h * aggregate_pmf[i];
            }
            if (tailProb <= 0) return double.NaN;
            return tailSum / tailProb;
        }

        /// <summary>
        /// Compute mean of aggregate distribution.
        /// </summary>
        [ExcelFunction(Description = "Mean of aggregate claims distribution", Category = "Actuarial.Experimental")]
        public static double ACT_AGGREGATE_MEAN(
            [ExcelArgument(Description = "Aggregate PMF array")] double[] aggregate_pmf,
            [ExcelArgument(Description = "Grid spacing h")] double h)
        {
            if (aggregate_pmf == null || h <= 0) return double.NaN;

            double mean = 0;
            for (int i = 0; i < aggregate_pmf.Length; i++)
            {
                mean += i * h * aggregate_pmf[i];
            }
            return mean;
        }

        /// <summary>
        /// Compute variance of aggregate distribution.
        /// </summary>
        [ExcelFunction(Description = "Variance of aggregate claims distribution", Category = "Actuarial.Experimental")]
        public static double ACT_AGGREGATE_VAR_STAT(
            [ExcelArgument(Description = "Aggregate PMF array")] double[] aggregate_pmf,
            [ExcelArgument(Description = "Grid spacing h")] double h)
        {
            if (aggregate_pmf == null || h <= 0) return double.NaN;

            double mean = 0;
            double meanSq = 0;
            for (int i = 0; i < aggregate_pmf.Length; i++)
            {
                double x = i * h;
                mean += x * aggregate_pmf[i];
                meanSq += x * x * aggregate_pmf[i];
            }
            return meanSq - mean * mean;
        }

        /// <summary>
        /// Compute standard deviation of aggregate distribution.
        /// </summary>
        [ExcelFunction(Description = "Standard deviation of aggregate claims distribution", Category = "Actuarial.Experimental")]
        public static double ACT_AGGREGATE_STDEV(
            [ExcelArgument(Description = "Aggregate PMF array")] double[] aggregate_pmf,
            [ExcelArgument(Description = "Grid spacing h")] double h)
        {
            double variance = ACT_AGGREGATE_VAR_STAT(aggregate_pmf, h);
            if (double.IsNaN(variance)) return double.NaN;
            return Math.Sqrt(variance);
        }

        #endregion
    }
}
