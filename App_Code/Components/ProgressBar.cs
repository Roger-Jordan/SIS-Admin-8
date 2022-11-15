using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace SIS.Components
{
    public enum ProgressBarSize
    {
        Normal,
        Small,
        Large,
    }

    public enum ProgressBarColor
    {
        Blue,
        Green,
        Yellow,
        Orange,
        Red,
    }

    public class ProgressBar : Panel
    {
        private int _percentage;

        public int Percentage
        {
            get
            {
                return _percentage;
            }
            set
            {
                _percentage = value;
                if (_percentage < 0)
                    _percentage = 0;
                if (_percentage > 100)
                    _percentage = 100;
            }
        }

        public ProgressBarColor Color
        {
            get;
            set;
        }

        public ProgressBarSize Size
        {
            get;
            set;
        }

        public bool Outlined
        {
            get;
            set;
        }


        public ProgressBar()
        {
            Percentage = 0;
            Color = ProgressBarColor.Blue;
            Size = ProgressBarSize.Normal;
            Outlined = false;
        }

        public ProgressBar(int percentage, ProgressBarColor color = ProgressBarColor.Blue, ProgressBarSize size = ProgressBarSize.Normal)
        {
            Percentage = percentage;
            Color = color;
            Size = size;
            Outlined = false;
        }

        protected override void Render(HtmlTextWriter writer)
        {
            /*
            <div class="c-progress-bar">
                <div class="value blue" style="width: " + tempWidth.ToString() + "%"></div>
            </div>
            */

            writer.AddAttribute(HtmlTextWriterAttribute.Class, GetProgressBarCssClasses());
            if (!string.IsNullOrEmpty(this.ClientID))
                writer.AddAttribute(HtmlTextWriterAttribute.Id, this.ClientID);
            writer.RenderBeginTag(HtmlTextWriterTag.Div);

            writer.AddAttribute(HtmlTextWriterAttribute.Class, GetValueBarCssClasses());
            writer.AddStyleAttribute("width", _percentage.ToString() + "%");
            writer.RenderBeginTag(HtmlTextWriterTag.Div);
            writer.Write(this._percentage + " %");
            writer.RenderEndTag(); // .value

            writer.RenderEndTag(); // .c.progress-bar
        }

        protected string GetProgressBarCssClasses()
        {
            string classes = "c-progress-bar";
            if (Size == ProgressBarSize.Small)
                classes += " small";
            if (Size == ProgressBarSize.Large)
                classes += " large";
            if (Outlined)
                classes += " outlined";
            return classes;
        }

        protected string GetValueBarCssClasses()
        {
            string classes = "value";
            switch (Color)
            {
                case ProgressBarColor.Blue:
                    classes += " blue";
                    break;
                case ProgressBarColor.Green:
                    classes += " green";
                    break;
                case ProgressBarColor.Yellow:
                    classes += " yellow";
                    break;
                case ProgressBarColor.Orange:
                    classes += " orange";
                    break;
                case ProgressBarColor.Red:
                    classes += " red";
                    break;
            }
            return classes;
        }

    }
}
