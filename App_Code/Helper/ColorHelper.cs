using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Web;
using SIS.Extensions;


namespace SIS.Helper
{
    public abstract class ColorHelper
    {
        public static Color FromHex(string hexColorCode)
        {
            hexColorCode = hexColorCode.TrimStart('#');

            Color color;
            if (hexColorCode.Length == 6)
            {
                color = Color.FromArgb(255, // hardcoded opaque
                            int.Parse(hexColorCode.Substring(0, 2), NumberStyles.HexNumber),
                            int.Parse(hexColorCode.Substring(2, 2), NumberStyles.HexNumber),
                            int.Parse(hexColorCode.Substring(4, 2), NumberStyles.HexNumber));
            }
            else
            {
                color = Color.FromArgb(
                            int.Parse(hexColorCode.Substring(0, 2), NumberStyles.HexNumber),
                            int.Parse(hexColorCode.Substring(2, 2), NumberStyles.HexNumber),
                            int.Parse(hexColorCode.Substring(4, 2), NumberStyles.HexNumber),
                            int.Parse(hexColorCode.Substring(6, 2), NumberStyles.HexNumber));
            }
            return color;
        }

        public static Color FromHex(string hexColorCode, int alpha)
        {
            Color color = FromHex(hexColorCode);
            return Color.FromArgb(alpha, color);
        }

        public static List<Color> GetColorSteps(Color from, Color to, int steps)
        {
            List<Color> colors = new List<Color>();

            HSVColor hsvFrom = from.ToHSV();
            HSVColor hsvTo = to.ToHSV();
            HSVColor colorStep;

            colors.Add(from);
            if (steps > 1)
            {
                double deltaAlpha = (hsvTo.Alpha - hsvFrom.Alpha) / (steps - 1);
                double deltaHue = (hsvTo.Hue - hsvFrom.Hue) / (steps - 1);
                double deltaSaturation = (hsvTo.Saturation - hsvFrom.Saturation) / (steps - 1);
                double deltaValue = (hsvTo.Value - hsvFrom.Value) / (steps - 1);
                for (int i = 1; i < steps - 1; i++)
                {
                    colorStep = hsvFrom
                            .AddAlpha((int)(i * deltaAlpha))
                            .AddHue(i * deltaHue)
                            .AddSaturation(i * deltaSaturation)
                            .AddValue(i * deltaValue);
                    Color color = colorStep.ToColor();
                    colors.Add(color);
                }
                colors.Add(to);
            }
            return colors;
        }

        public static Color GetNextColor(Color baseColor, int alpha, int step)
        {
            return GetNextColor(Color.FromArgb(alpha, baseColor), step);
        }

        public static Color GetNextColor(Color baseColor, int step)
        {
            // Ein paar Primzahlen:
            // 17,    19,    23,    29,    31,    37,    41,    43,
            // 47,    53,    59,    61,    67,    71,    73,    79,
            // 83,    89,    97,   101,   103,   107

            HSVColor hsvColor = baseColor.ToHSV();
            hsvColor = hsvColor.AddHue(29 * step);
            Color result = hsvColor.ToColor();
            return result;
        }

        public static Color Darken(Color color, double percentage)
        {
            return color.ToHSV().Darken(percentage).ToColor();
        }

        public static Color Lighten(Color color, double percentage)
        {
            return color.ToHSV().Lighten(percentage).ToColor();
        }
    }
}