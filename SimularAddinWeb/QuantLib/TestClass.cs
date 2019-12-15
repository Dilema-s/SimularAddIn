using System;
using QLNet;

namespace SimularAddinWeb.QuantLib
{
    public class TestClass
    {
        public double Derivative(double X)
        {
            NormalDistribution nd = new NormalDistribution();
            double result = nd.derivative(X);
            return result;
        }
    }
}