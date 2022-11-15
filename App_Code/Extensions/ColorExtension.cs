using System;
using System.Drawing;
using System.Globalization;


namespace SIS.Extensions
{
    public static class ColorExtension
    {
        public static HSVColor ToHSV(this Color color)
        {
            return HSVColor.FromColor(color);
        }

        public static string ToHexString(this Color color, bool includeLeadingHash = false)
        {
            string result = color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2");
            if (includeLeadingHash)
                result = "#" + result;
            return result;
        }

        public static string ToRgbString(this Color color)
        {
            string result = string.Format("rgb({0}, {1}, {2});", color.R, color.G, color.B);
            return result;
        }

        public static string ToRgbaString(this Color color)
        {
            double alpha = (double)color.A / 255;
            NumberFormatInfo numberFormatInfo = new NumberFormatInfo();
            numberFormatInfo.NumberDecimalSeparator = ".";
            string result = string.Format("rgba({0}, {1}, {2}, {3});", color.R, color.G, color.B, alpha.ToString("0.##", numberFormatInfo));
            return result;
        }
    }
}
