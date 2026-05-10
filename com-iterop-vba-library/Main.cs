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

        /// <summary>
        /// Converts API gravity to specific gravity relative to water.
        /// </summary>
        /// <remarks>
        /// HOW TO USE:
        /// Call this method to convert between API and specific gravity scales.
        /// API gravity is a measure commonly used in the petroleum industry.
        /// Specific gravity compares oil density to water density.
        /// 
        /// EXAMPLE IN VBA:
        ///   Dim calc As Object
        ///   Set calc = CreateObject("vbdl")
        ///   Dim sg As Double
        ///   sg = calc.ApiToSpecificGravity(45.5)
        ///   MsgBox "Specific Gravity: " &amp; sg
        /// </remarks>
        /// <param name="apiGravity">API gravity value (typically 10–80 for crude oils).</param>
        /// <returns>Specific gravity relative to water (dimensionless).</returns>
        [SuppressMessage("Performance", "CA1822:Mark members as static",
            Justification = "COM interop requires instance methods; static methods are not included in the AutoDual type library.")]
        public double ApiToSpecificGravity(double apiGravity)
        {
            return 141.5 / (apiGravity + 131.5);
        }

        /// <summary>
        /// Converts specific gravity to API gravity.
        /// </summary>
        /// <remarks>
        /// HOW TO USE:
        /// Use this to convert specific gravity values to the API scale.
        /// Useful when you have oil density data but need API gravity for industry-standard reports.
        /// 
        /// EXAMPLE IN VBA:
        ///   Dim calc As Object
        ///   Set calc = CreateObject("vbdl")
        ///   Dim api As Double
        ///   api = calc.SpecificGravityToApi(0.8)
        ///   MsgBox "API Gravity: " &amp; api &amp; " degrees"
        /// </remarks>
        /// <param name="specificGravity">Specific gravity relative to water. Must be greater than zero.</param>
        /// <returns>API gravity (degrees).</returns>
        [SuppressMessage("Performance", "CA1822:Mark members as static",
            Justification = "COM interop requires instance methods; static methods are not included in the AutoDual type library.")]
        public double SpecificGravityToApi(double specificGravity)
        {
            if (specificGravity <= 0.0)
                throw new ArgumentOutOfRangeException("specificGravity", "Specific gravity must be greater than zero.");
            return (141.5 / specificGravity) - 131.5;
        }

        /// <summary>
        /// Calculates formation volume factor (Bo) for oil at a given pressure and temperature.
        /// </summary>
        /// <remarks>
        /// HOW TO USE:
        /// Formation volume factor is the ratio of volume at reservoir conditions to volume at standard conditions.
        /// This method uses a simplified correlation for quick field estimates.
        /// Bo values typically range from 1.0 to 2.0 for most crude oils.
        /// 
        /// EXAMPLE IN VBA:
        ///   Dim calc As Object
        ///   Set calc = CreateObject("vbdl")
        ///   Dim bo As Double
        ///   bo = calc.CalculateFormationVolumeFactor(0.8, 25000, 70)
        ///   MsgBox "Formation Volume Factor (Bo): " &amp; bo
        /// </remarks>
        /// <param name="specificGravity">Oil specific gravity relative to water. Must be greater than zero.</param>
        /// <param name="pressureKPa">Current pressure (kPa). Must be greater than zero.</param>
        /// <param name="temperatureCelsius">Temperature (°C).</param>
        /// <returns>Formation volume factor (dimensionless, typically 0.9–2.5).</returns>
        [SuppressMessage("Performance", "CA1822:Mark members as static",
            Justification = "COM interop requires instance methods; static methods are not included in the AutoDual type library.")]
        public double CalculateFormationVolumeFactor(double specificGravity, double pressureKPa, double temperatureCelsius)
        {
            if (specificGravity <= 0.0)
                throw new ArgumentOutOfRangeException("specificGravity", "Oil specific gravity must be greater than zero.");
            if (pressureKPa <= 0.0)
                throw new ArgumentOutOfRangeException("pressureKPa", "Pressure must be greater than zero.");

            double api = 141.5 / specificGravity - 131.5;
            double pPsi = pressureKPa * 14.22334;
            double tRankine = (temperatureCelsius * 1.8 + 32.0) + 459.67;

            double bo = 0.98 + 0.0001 * api + 0.000005 * pPsi + 0.00001 * tRankine;
            return bo;
        }

        /// <summary>
        /// Calculates gas viscosity at given pressure and temperature using a simplified correlation.
        /// </summary>
        /// <remarks>
        /// HOW TO USE:
        /// Gas viscosity is essential for flow rate predictions and pressure drop calculations.
        /// This method uses a simplified correlation valid for natural gases at moderate pressures.
        /// Results are in centipoise (cP).
        /// 
        /// EXAMPLE IN VBA:
        ///   Dim calc As Object
        ///   Set calc = CreateObject("vbdl")
        ///   Dim gasVis As Double
        ///   gasVis = calc.CalculateGasViscosity(30000, 80)
        ///   MsgBox "Gas Viscosity: " &amp; gasVis &amp; " cP"
        /// </remarks>
        /// <param name="pressureKPa">Gas pressure (kPa). Must be greater than zero.</param>
        /// <param name="temperatureCelsius">Gas temperature (°C).</param>
        /// <returns>Gas viscosity in centipoise (cP).</returns>
        [SuppressMessage("Performance", "CA1822:Mark members as static",
            Justification = "COM interop requires instance methods; static methods are not included in the AutoDual type library.")]
        public double CalculateGasViscosity(double pressureKPa, double temperatureCelsius)
        {
            if (pressureKPa <= 0.0)
                throw new ArgumentOutOfRangeException("pressureKPa", "Pressure must be greater than zero.");

            double pAtm = pressureKPa / 101.325;
            double tKelvin = temperatureCelsius + 273.15;

            // Simplified Sutherland-based correlation for natural gas
            double viscosityAtAtm = 0.0001 + 0.000002 * (tKelvin - 273.15);
            double viscosity = viscosityAtAtm * Math.Sqrt(pAtm);

            return viscosity;
        }

        /// <summary>
        /// Calculates hydrostatic pressure at a given depth.
        /// </summary>
        /// <remarks>
        /// HOW TO USE:
        /// This simple method calculates hydrostatic pressure for water columns.
        /// Useful for initial estimates in wells with water zones.
        /// Assumes standard water density of 1000 kg/m³.
        /// 
        /// EXAMPLE IN VBA:
        ///   Dim calc As Object
        ///   Set calc = CreateObject("vbdl")
        ///   Dim hydroPres As Double
        ///   hydroPres = calc.CalculateHydrostaticPressure(2500)
        ///   MsgBox "Hydrostatic Pressure at 2500m: " &amp; hydroPres &amp; " kPa"
        /// </remarks>
        /// <param name="depthMeters">Vertical depth in meters. Must be greater than or equal to zero.</param>
        /// <returns>Hydrostatic pressure in kPa.</returns>
        [SuppressMessage("Performance", "CA1822:Mark members as static",
            Justification = "COM interop requires instance methods; static methods are not included in the AutoDual type library.")]
        public double CalculateHydrostaticPressure(double depthMeters)
        {
            if (depthMeters < 0.0)
                throw new ArgumentOutOfRangeException("depthMeters", "Depth must be greater than or equal to zero.");

            // Standard water density = 1000 kg/m³, g = 9.81 m/s²
            // P = ρ * g * h, result in Pa then converted to kPa
            double pressurePa = 1000.0 * 9.81 * depthMeters;
            return pressurePa / 1000.0;
        }
    }
}
