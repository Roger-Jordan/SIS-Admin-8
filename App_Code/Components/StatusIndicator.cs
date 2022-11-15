using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;


namespace SIS.Components
{
    public enum StatusIndicatorSize
    {
        Small,
        Default,
        Large
    }

    public enum StatusIndicatorColor
    {
        None,
        Blue,
        Green,
        Yellow,
        Orange,
        Red,
    }

    public class StatusIndicator : Label
    {
        private StatusIndicatorSize _size = StatusIndicatorSize.Default;
        private StatusIndicatorColor _color = StatusIndicatorColor.None;
        private bool _locked = false;

        public StatusIndicatorSize Size {
            get { return _size; }
            set { _size = value; }
        }

        public StatusIndicatorColor Color {
            get { return _color; }
            set { _color = value; }
        }

        public bool Locked {
            get { return _locked; }
            set { _locked = value; }
        }

        protected override void Render(HtmlTextWriter writer)
        {
            // Wenn keine Fabe defniert ist, dann wird auch kein Status-Indicator ausgegeben
            if (_color == StatusIndicatorColor.None)
                return;

            // <span class="status-dot locked green"></span>
            writer.AddAttribute(HtmlTextWriterAttribute.Class, GetCssClasses());
            if (!string.IsNullOrEmpty(ToolTip))
            {
                writer.AddAttribute("title", this.Page.Server.HtmlEncode(ToolTip));

            }
            writer.RenderBeginTag(HtmlTextWriterTag.Span);
            writer.RenderEndTag();
        }

        protected string GetCssClasses()
        {
            StringBuilder classes = new StringBuilder();
            classes.Append("status-dot");

            switch (_size)
            {
                case StatusIndicatorSize.Small:
                    classes.Append(" small");
                    break;

                case StatusIndicatorSize.Large:
                    classes.Append(" large");
                    break;
            }

            switch (_color)
            {
                case StatusIndicatorColor.Blue:
                    classes.Append(" blue");
                    break;

                case StatusIndicatorColor.Green:
                    classes.Append(" green");
                    break;

                case StatusIndicatorColor.Yellow:
                    classes.Append(" yellow");
                    break;

                case StatusIndicatorColor.Orange:
                    classes.Append(" orange");
                    break;

                case StatusIndicatorColor.Red:
                    classes.Append(" red");
                    break;
            }

            if (_locked)
            {
                classes.Append(" locked");
            }
            return classes.ToString();
        }
    }
}