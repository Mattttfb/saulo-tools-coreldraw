using System;
using System.Globalization;
using Corel.Interop.VGCore;

namespace SauloTools.Core
{
    public static class UnitHelper
    {
        public const double CmPerInch = 2.54;
        public const double PtPerInch = 72.0;

        public static double CmToInches(double cm) => cm / CmPerInch;

        public static double MmToInches(double mm) => mm / 25.4;

        public static double PtToCm(double pt) => (pt / PtPerInch) * CmPerInch;

        public static double MmToDocumentUnit(double valMm, cdrUnit unit)
        {
            switch (unit)
            {
                case cdrUnit.cdrMillimeter: return valMm;
                case cdrUnit.cdrCentimeter: return valMm / 10.0;
                case cdrUnit.cdrMeter: return valMm / 1000.0;
                case cdrUnit.cdrInch: return valMm / 25.4;
                case cdrUnit.cdrFoot: return valMm / 304.8;
                default: return valMm;
            }
        }

        public static bool TryParseDouble(string text, out double value)
        {
            value = 0;
            if (string.IsNullOrWhiteSpace(text)) return false;
            string normalized = text.Trim().Replace(',', '.');
            return double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
        }

        public static bool TryParseInt(string text, out int value)
        {
            value = 0;
            if (string.IsNullOrWhiteSpace(text)) return false;
            return int.TryParse(text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out value);
        }
    }
}
