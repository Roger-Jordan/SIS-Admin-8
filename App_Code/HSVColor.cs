using System;
using System.Drawing;


namespace SIS
{
    public class HSVColor
    {
        public int Alpha;
        public double Hue;
        public double Saturation;
        public double Value;


        public HSVColor(double hue = 0, double saturation = 0, double value = 0)
        {
            Alpha = 255;
            Hue = hue;
            Saturation = saturation;
            Value = value;
        }

        public HSVColor(int alpha, double hue = 0, double saturation = 0, double value = 0)
        {
            Alpha = alpha;
            Hue = hue;
            Saturation = saturation;
            Value = value;
        }

        public HSVColor(HSVColor color)
        {
            Alpha = color.Alpha;
            Hue = color.Hue;
            Saturation = color.Saturation;
            Value = color.Value;
        }

        public HSVColor AddAlpha(int alpha)
        {
            HSVColor color = new HSVColor(this);
            color.Saturation = Math.Min(Math.Max(0, this.Alpha + alpha), 255);
            return color;
        }

        public HSVColor AddHue(double hue)
        {
            HSVColor color = new HSVColor(this);
            hue = hue % 360;
            color.Hue = (color.Hue + hue) % 360;
            if (color.Hue < 0)
                color.Hue = 360 + color.Hue;
            return color;
        }

        public HSVColor AddSaturation(double saturation)
        {
            HSVColor color = new HSVColor(this);
            color.Saturation = Math.Min(Math.Max(0.0, this.Saturation + saturation), 1.0);
            return color;
        }

        public HSVColor AddValue(double value)
        {
            HSVColor color = new HSVColor(this);
            color.Value = Math.Min(Math.Max(0.0, this.Value + value), 1.0);
            return color;
        }

        public HSVColor Darken(double percentage)
        {
            return AddValue(-1 * Math.Abs(percentage) / 100);
        }

        public HSVColor Lighten(double percentage)
        {
            return AddValue(Math.Abs(percentage) / 100);
        }

        public Color ToColor()
        {
            return ToColor(this.Hue, this.Saturation, this.Value, this.Alpha);
        }

        public static HSVColor FromColor(Color color)
        {
            int max = Math.Max(color.R, Math.Max(color.G, color.B));
            int min = Math.Min(color.R, Math.Min(color.G, color.B));

            HSVColor hsvColor = new HSVColor();
            hsvColor.Alpha = color.A;
            hsvColor.Hue = color.GetHue();
            hsvColor.Saturation = (max == 0) ? 0 : 1d - (1d * min / max);
            hsvColor.Value = max / 255d;

            return hsvColor;
        }

        public static Color ToColor(double hue, double saturation, double value, int alpha = 255)
        {
            int hi = Convert.ToInt32(Math.Floor(hue / 60)) % 6;
            double f = hue / 60 - Math.Floor(hue / 60);

            value = value * 255;
            int v = Convert.ToInt32(value);
            int p = Convert.ToInt32(value * (1 - saturation));
            int q = Convert.ToInt32(value * (1 - f * saturation));
            int t = Convert.ToInt32(value * (1 - (1 - f) * saturation));

            if (hi == 0)
                return Color.FromArgb(alpha, v, t, p);
            else if (hi == 1)
                return Color.FromArgb(alpha, q, v, p);
            else if (hi == 2)
                return Color.FromArgb(alpha, p, v, t);
            else if (hi == 3)
                return Color.FromArgb(alpha, p, q, v);
            else if (hi == 4)
                return Color.FromArgb(alpha, t, p, v);
            else
                return Color.FromArgb(alpha, v, p, q);
        }
    }
}