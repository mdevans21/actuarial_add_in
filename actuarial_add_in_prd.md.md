# Actuarial excel addin

This addin is to be written in C# using the standard approach for excel addins in Windows. It should be compiled for deployment using the appropriate GitHub hook. 

## Features:

1. Pdf, cdf and inverse cdf distributions should be included for poisson, negative binomial, lognormal, gamma and pareto distributions. 
2. A student-t copula should be included for generating correlated random numbers based on a correlation matrix. 
3. Exposure curves should be included, including the Swiss Re curves (parameterized using the mbbefd curves) and the Lloyd's curves. Power curves and Pareto curves. 
4. Copy from chainladder python package: bootstrapping, mack bootstrap, berquist sherman. 
5. Excess of loss layer calculations. 
6. Formulae to generate losses from a return period table. 
7. A formula for linear interpolation. Include an option for extrapolation out of range: default to flat but allow constant assumed gradient from last two points. 

## Tests:

Write a test suite that works in batch, calling same c# functions from the command line, to produce a short markdown file with all results.



