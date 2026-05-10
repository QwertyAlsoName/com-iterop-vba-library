using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

namespace Iterop.VbaLibrary
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("2538A44F-A746-47F9-918E-C28BF261C372")]
    [ComVisible(true)]
    public interface IVbaHelper
    {
        string Hello();
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("ADCF43D0-9B5D-446C-9D69-3DAA2DC31015")]
    [ProgId("vbdl")]
    [ComVisible(true)]
    public class VbaHelper : IVbaHelper
    {
        public string Hello()
        {
            return "Hello";
        }
    }

    /// <summary>
    /// Reservoir engineering calculations exposed to VBA via COM interop.
    /// Pressures are in kPa, temperature in °C, GOR in m³/m³, specific gravity dimensionless.
    /// </summary>
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Guid("97AFB84F-033D-40F2-A146-C917AD978141")]
    [ComVisible(true)]
    public class DotNetCalc
    {
        /// <summary>
        /// Calculates the solution gas-to-oil ratio (SGR) at reservoir conditions.
        /// </summary>
        /// <param name="porosity">Formation porosity (fraction, 0–1).</param>
        /// <param name="bubblePressureKPa">Bubble-point pressure (kPa).</param>
        /// <param name="temperatureCelsius">Reservoir temperature (°C).</param>
        /// <param name="gorAtBubblePoint">Solution GOR at bubble point (m³/m³).</param>
        /// <param name="waterCompressibility">Water compressibility (1/kPa × 10⁻⁶).</param>
        /// <param name="oilSpecificGravity">Oil specific gravity relative to water (dimensionless). Must be greater than zero.</param>
        /// <param name="minimumPressureKPa">Minimum reservoir pressure (kPa). Must be greater than zero.</param>
        /// <returns>Solution gas-to-oil ratio (scf/bbl).</returns>
        [SuppressMessage("Performance", "CA1822:Mark members as static",
            Justification = "COM interop requires instance methods; static methods are not included in the AutoDual type library.")]
        public double Sgr(
            double porosity,
            double bubblePressureKPa,
            double temperatureCelsius,
            double gorAtBubblePoint,
            double waterCompressibility,
            double oilSpecificGravity,
            double minimumPressureKPa)
        {
            if (minimumPressureKPa <= 0.0)
                throw new ArgumentOutOfRangeException("minimumPressureKPa", "Minimum pressure must be greater than zero.");
            if (oilSpecificGravity <= 0.0)
                throw new ArgumentOutOfRangeException("oilSpecificGravity", "Oil specific gravity must be greater than zero.");
            if (waterCompressibility <= 0.0)
                throw new ArgumentOutOfRangeException("waterCompressibility", "Water compressibility must be greater than zero.");

            // Convert to field units: kPa → psi, °C → °F, m³/m³ → scf/bbl, SG → °API
            double pb   = bubblePressureKPa  * 14.22334;
            double t    = temperatureCelsius * 1.8 + 32.0;
            double rspb = gorAtBubblePoint   * 5.6145821;
            double api  = 141.5 / oilSpecificGravity - 131.5;
            double pmin = minimumPressureKPa * 14.22334;

            double a = -758.0
                       + 0.86   * porosity
                       + 5.29   * waterCompressibility
                       - 0.0444 * Math.Pow(waterCompressibility, 2)
                       - 77.2   * Math.Log(waterCompressibility)
                       - 0.0386 * rspb;

            double b = a
                       + 0.0000533 * Math.Pow(rspb, 2)
                       + 4.06      * api
                       - 0.057     * Math.Pow(api, 2);

            double pressureRatioTerm = 0.000156 * Math.Pow(pb / pmin, 2);

            double c = b
                       + pressureRatioTerm
                       + 2.02       * Math.Log(pb / pmin)
                       + 0.0123     * pb
                       - 0.00000369 * Math.Pow(pb, 2);

            return c
                   - 3.71    * t
                   + 0.00643 * Math.Pow(t, 2)
                   + 253.0   * Math.Log(t);
        }
    }
}
