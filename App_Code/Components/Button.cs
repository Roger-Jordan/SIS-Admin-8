using System.IO;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text.RegularExpressions;
using SIS.Extensions;


namespace SIS.Components
{
    public enum ThrobberTheme
    {
        None,
        White,
        Blue,
        Dark
    }

    public enum ButtonSize
    {
        ExtraSmall,
        Small,
        Normal,
        Large
    }


    public class Button : System.Web.UI.WebControls.Button
    {
        private bool _disableAfterClick = false;
        private bool _useSubmitBehavior = true;
        protected ThrobberTheme _showThrobberAfterClick;
        protected bool _resetAfterPostback = false;
        private string _openDialogId = null;
        private bool _showOnlyIcon = false;

        public Button()
        {
            Size = ButtonSize.Normal;
            Icon = null;
            ShowThrobberAfterClick = ThrobberTheme.None;
            DisableAutoCssClasses = false;

            CssClass = ""; // Sorgt dafür, dass die CssClass initialisiert wird
                           // MUSS am Ende stehen, nachdem alle anderen Standard-Werte gesetzt wurden
        }

        public bool DisableAutoCssClasses
        {
            get;
            set;
        }

        public override string CssClass
        {
            get
            {
                return base.CssClass;
            }
            set
            {
                base.CssClass = value;
                UpdateCssClass();
            }
        }

        public ButtonSize Size
        {
            get;
            set;
        }

        public string Icon
        {
            get;
            set;
        }

        public bool DisableAfterClick
        {
            get
            {
                return this._disableAfterClick;
            }
            set
            {
                this._disableAfterClick = value;
                UpdateCssClass();
                if (this._disableAfterClick)
                {
                    base.UseSubmitBehavior = false;
                }
                else
                {
                    base.UseSubmitBehavior = this._useSubmitBehavior;
                }
            }
        }

        public override bool UseSubmitBehavior
        {
            get
            {
                return base.UseSubmitBehavior;
            }
            set
            {
                if (this._disableAfterClick && value)
                {
                    this.DisableAfterClick = false;
                }
                base.UseSubmitBehavior = value;
                this._useSubmitBehavior = value;
            }
        }

        public ThrobberTheme ShowThrobberAfterClick
        {
            get
            {
                return _showThrobberAfterClick;
            }
            set
            {
                _showThrobberAfterClick = value;
                UpdateCssClass();
            }
        }

        public bool ResetAfterPostback
        {
            get
            {
                return _resetAfterPostback;
            }
            set
            {
                _resetAfterPostback = value;
                UpdateCssClass();
            }
        }

        public bool ShowOnlyIcon
        {
            get
            {
                return _showOnlyIcon;
            }
            set
            {
                _showOnlyIcon = value;
                UpdateCssClass();
            }
        }



        [System.ComponentModel.TypeConverter(typeof(AssociatedControlConverter))]
        public string ConfirmationDialogId
        {
            get;
            set;
        }

        [System.ComponentModel.TypeConverter(typeof(AssociatedControlConverter))]
        public string OpenDialogId
        {
            get
            {
                return _openDialogId;
            }
            set
            {
                _openDialogId = value;
                if (!string.IsNullOrEmpty(_openDialogId))
                {
                    UseSubmitBehavior = false;
                }
            }
        }

        protected virtual void UpdateCssClass()
        {
            if (!DisableAutoCssClasses)
            {
                string autoCssClasses = GetAutoCssClasses();
                if (!base.CssClass.Contains(autoCssClasses))
                {
                    base.CssClass = autoCssClasses + " " + base.CssClass;
                }
            }

            if (DisableAfterClick)
                base.CssClass = base.CssClass.AppendIfNotEmpty("disable-on-click");

            if (ShowThrobberAfterClick != ThrobberTheme.None)
            {
                base.CssClass = base.CssClass.AppendIfNotEmpty("show-throbber-on-click");
                this.Attributes["data-throbber-theme"] = this.ShowThrobberAfterClick.ToString().ToLower();
            }

            if (ResetAfterPostback)
                base.CssClass = base.CssClass.AppendIfNotEmpty("reset-after-postback");

            if (ShowOnlyIcon)
                base.CssClass = base.CssClass.AppendIfNotEmpty("only-icon");
        }

        protected virtual string GetAutoCssClasses()
        {
            string cssClasses = "";

            string themeCssClasses = GetThemeCssClasses();
            cssClasses = cssClasses.AppendIfNotEmpty(themeCssClasses);

            return cssClasses;
        }

        protected virtual string GetThemeCssClasses()
        {
            return "btn btn-default";
        }

        protected virtual string GetSizeCssClass()
        {
            switch (Size)
            {
                case ButtonSize.ExtraSmall: return "btn-xs";
                case ButtonSize.Small: return "btn-sm";
                case ButtonSize.Large: return "btn-lg";
                default: return "";
            }
        }

        // Sorgt dafür, dass der Button immer als <button></button> gerendert wird
        protected override HtmlTextWriterTag TagKey
        {
            get { return HtmlTextWriterTag.Button; }
        }

        // Modifiziert das onclick-Event, wenn DisableAfterClick oder ShowThrobberAfterClick
        // auf true stehen. Dann wird jeweiliger JavaScript-Code eingebaut.
        public override void RenderBeginTag(System.Web.UI.HtmlTextWriter writer)
        {
            if (ShowOnlyIcon && string.IsNullOrEmpty(this.ToolTip))
                this.ToolTip = this.Text;

            // Temporären StringWriter anlegen, sodass base.RenderBeginTag-Ausgabe abgefangen und
            // geparst werden kann
            StringWriter stringWriter = new StringWriter();
            HtmlTextWriter tempHtmlWriter = new HtmlTextWriter(stringWriter);
            base.RenderBeginTag(tempHtmlWriter);
            string beginTagHtml = stringWriter.ToString();

            if (this._disableAfterClick || this.ShowThrobberAfterClick != ThrobberTheme.None)
            {
                //string javascript = GenerateAdditionalClientClickJavaScript();
                //beginTagHtml = AddClientClickJavascriptToHtml(beginTagHtml, javascript);
            }

            if (!string.IsNullOrEmpty(ConfirmationDialogId))
            {
                // bootstrap-confirmation
                Control targetControl = this.NamingContainer.FindControl(ConfirmationDialogId);
                if (targetControl != null)
                {
                    beginTagHtml = AddConfirmationDialogOptionsToHtml(beginTagHtml, targetControl.ClientID);
                }
            }
            else if (!string.IsNullOrEmpty(OpenDialogId))
            {
                // bootstrap-confirmation
                Control targetControl = this.NamingContainer.FindControl(OpenDialogId);
                if (targetControl != null)
                {
                    beginTagHtml = AddOpenDialogOptionsToHtml(beginTagHtml, targetControl.ClientID);
                }
            }

            string sizeCssClass = GetSizeCssClass();
            if (!string.IsNullOrEmpty(sizeCssClass))
            {
                beginTagHtml = AddCssClassToHtml(beginTagHtml, sizeCssClass);
            }

            writer.Write(beginTagHtml);
        }

        private string AddClientClickJavascriptToHtml(string html, string javascript)
        {
            string pattern = "onclick=\"([^\"]*)\"";
            if (Regex.IsMatch(html, pattern))
            {
                // Erweitert den JavaScript-Code im bestehenden onclick-Attribut
                html = Regex.Replace(html, pattern, m => string.Format("onclick=\"{0} {1}\"", m.Groups[1].Value, javascript));
            }
            else
            {
                html = html.Insert(html.Length - 1, string.Format(" onclick=\"{0}\"", javascript));
            }

            return html;
        }

        private string AddConfirmationDialogOptionsToHtml(string html, string clientID)
        {
            html = html.Insert(html.Length - 1, string.Format(" data-bootstrap-dialog=\"{0}\"", clientID));
            html = AddCssClassToHtml(html, "bootstrap-confirmation");
            return html;
        }

        private string AddOpenDialogOptionsToHtml(string html, string clientID)
        {
            html = html.Insert(html.Length - 1, " data-toggle=\"modal\"");
            html = html.Insert(html.Length - 1, string.Format(" data-target=\"#{0}\"", clientID));

            string pattern = "onclick=\"([^\"]*)\"";
            if (Regex.IsMatch(html, pattern))
            {
                // Erweitert den JavaScript-Code im bestehenden onclick-Attribut
                html = Regex.Replace(html, pattern, m => "onclick=\"return false;\"");
            }

            return html;
        }

        private string AddCssClassToHtml(string html, string cssClass)
        {
            string pattern = "class=\"([^\"]*)\"";
            if (Regex.IsMatch(html, pattern))
            {
                html = Regex.Replace(html, pattern, m => string.Format("class=\"{0} {1}\"", m.Groups[1].Value, cssClass));
            }
            else
            {
                html = html.Insert(html.Length - 1, string.Format(" class=\"{0}\"", cssClass));
            }
            return html;
        }

        private string GenerateAdditionalClientClickJavaScript()
        {
            string javascript = "";
            if (this._disableAfterClick)
            {
                this.CssClass += "disable-on-click";
                // javascript += ";this.disabled = true;";
            }
            if (this.ShowThrobberAfterClick != ThrobberTheme.None)
            {
                javascript += string.Format(";$(this).addClass('throbber-{0}');", this.ShowThrobberAfterClick.ToString().ToLower());
            }

            return javascript;
        }

        // Wird leider benötigt, da sonst eine Fehlermeldung mit PopEndTag/PushEndTag erscheint.
        public override void RenderEndTag(System.Web.UI.HtmlTextWriter writer)
        {
            writer.Write("</button>");
        }

        // Gibt nicht nur den Text des Buttons aus, sondern auch noch das Icon dazu.
        protected override void RenderContents(HtmlTextWriter writer)
        {
            string icon = "";
            if (this.Icon != null && this.Icon != "")
            {
                icon += "<i class=\"si-left si-" + this.Icon + "\"></i>";
            }
            if (ShowOnlyIcon)
                writer.Write(icon);
            else
                writer.Write(icon + Text);
        }
    }



    public class WhiteButton : Button
    {
        protected override string GetThemeCssClasses()
        {
            return "btn";
        }
    }

    public class PrimaryButton : Button
    {
        protected override string GetThemeCssClasses()
        {
            return "btn btn-primary";
        }
    }


    public class SuccessButton : Button
    {
        protected override string GetThemeCssClasses()
        {
            return "btn btn-success";
        }
    }


    public class InfoButton : Button
    {
        protected override string GetThemeCssClasses()
        {
            return "btn btn-info";
        }
    }


    public class WarningButton : Button
    {
        protected override string GetThemeCssClasses()
        {
            return "btn btn-warning";
        }
    }


    public class DangerButton : Button
    {
        protected override string GetThemeCssClasses()
        {
            return "btn btn-danger";
        }
    }
}