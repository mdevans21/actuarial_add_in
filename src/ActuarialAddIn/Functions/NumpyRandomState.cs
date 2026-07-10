namespace ActuarialAddIn.Functions;

/// <summary>
/// NumPy legacy RandomState-compatible random stream.
///
/// StochasticReserving uses np.random.seed/choice/gamma/normal, which are
/// backed by RandomState and MT19937. Matching only the generator is not
/// enough: bounded integers and distribution samplers must consume the stream
/// in the same way as NumPy's legacy randomkit implementation.
/// </summary>
internal sealed class NumpyRandomState
{
    private const int StateSize = 624;
    private const int Period = 397;
    private readonly uint[] _state = new uint[StateSize];
    private int _position = StateSize;
    private bool _hasGaussian;
    private double _gaussian;

    public NumpyRandomState(uint seed)
    {
        _state[0] = seed;
        for (int i = 1; i < StateSize; i++)
        {
            uint previous = _state[i - 1];
            _state[i] = unchecked(1812433253U * (previous ^ (previous >> 30)) + (uint)i);
        }
    }

    public uint NextUInt32()
    {
        if (_position >= StateSize)
            Twist();

        uint value = _state[_position++];
        value ^= value >> 11;
        value ^= (value << 7) & 0x9d2c5680U;
        value ^= (value << 15) & 0xefc60000U;
        value ^= value >> 18;
        return value;
    }

    public double NextDouble()
    {
        uint a = NextUInt32() >> 5;
        uint b = NextUInt32() >> 6;
        return (a * 67108864.0 + b) / 9007199254740992.0;
    }

    public int NextIndex(int count)
    {
        if (count <= 0)
            throw new ArgumentOutOfRangeException(nameof(count));

        uint maximum = (uint)(count - 1);
        if (maximum == 0)
            return 0;

        uint mask = maximum;
        mask |= mask >> 1;
        mask |= mask >> 2;
        mask |= mask >> 4;
        mask |= mask >> 8;
        mask |= mask >> 16;

        uint value;
        do
        {
            value = NextUInt32() & mask;
        }
        while (value > maximum);

        return (int)value;
    }

    public double Normal(double mean = 0.0, double standardDeviation = 1.0)
        => mean + standardDeviation * StandardNormal();

    public double Lognormal(double mean, double sigma)
        => Math.Exp(mean + sigma * StandardNormal());

    public double Gamma(double shape, double scale = 1.0)
    {
        if (shape <= 0.0 || scale < 0.0)
            return double.NaN;
        if (scale == 0.0)
            return 0.0;

        return scale * StandardGamma(shape);
    }

    private double StandardNormal()
    {
        if (_hasGaussian)
        {
            _hasGaussian = false;
            return _gaussian;
        }

        double x1;
        double x2;
        double radiusSquared;
        do
        {
            x1 = 2.0 * NextDouble() - 1.0;
            x2 = 2.0 * NextDouble() - 1.0;
            radiusSquared = x1 * x1 + x2 * x2;
        }
        while (radiusSquared >= 1.0 || radiusSquared == 0.0);

        double factor = Math.Sqrt(-2.0 * Math.Log(radiusSquared) / radiusSquared);
        _gaussian = factor * x1;
        _hasGaussian = true;
        return factor * x2;
    }

    private double StandardExponential()
        => -Math.Log(1.0 - NextDouble());

    private double StandardGamma(double shape)
    {
        if (shape == 1.0)
            return StandardExponential();

        if (shape < 1.0)
        {
            while (true)
            {
                double u = NextDouble();
                double v = StandardExponential();
                if (u <= 1.0 - shape)
                {
                    double x = Math.Pow(u, 1.0 / shape);
                    if (x <= v)
                        return x;
                }
                else
                {
                    double y = -Math.Log((1.0 - u) / shape);
                    double x = Math.Pow(1.0 - shape + shape * y, 1.0 / shape);
                    if (x <= v + y)
                        return x;
                }
            }
        }

        double b = shape - 1.0 / 3.0;
        double c = 1.0 / Math.Sqrt(9.0 * b);
        while (true)
        {
            double x = StandardNormal();
            double v = 1.0 + c * x;
            if (v <= 0.0)
                continue;

            v = v * v * v;
            double u = NextDouble();
            if (u < 1.0 - 0.0331 * x * x * x * x)
                return b * v;
            if (Math.Log(u) < 0.5 * x * x + b * (1.0 - v + Math.Log(v)))
                return b * v;
        }
    }

    private void Twist()
    {
        for (int i = 0; i < StateSize; i++)
        {
            uint value = (_state[i] & 0x80000000U) | (_state[(i + 1) % StateSize] & 0x7fffffffU);
            _state[i] = _state[(i + Period) % StateSize] ^ (value >> 1);
            if ((value & 1U) != 0)
                _state[i] ^= 0x9908b0dfU;
        }
        _position = 0;
    }
}
