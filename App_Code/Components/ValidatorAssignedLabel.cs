using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace SIS.Components
{
    public class ValidatorAssignedLabel : System.Web.UI.WebControls.Label
    {
        public const string DefaultValidationErrorTextPrefix = "<i class=\"si-left si-warning\"></i>";
        public const string DefaultErrorCssClass = "has-error";

        private List<BaseValidator> validators = new List<BaseValidator>();
        private string validationErrorTextPrefix = DefaultValidationErrorTextPrefix;
        private string validationErrorTextSuffix = null;
        private string errorCssClass = DefaultErrorCssClass;

        // Zum Setzen via XAML
        [System.ComponentModel.TypeConverter(typeof(AssociatedControlConverter))]
        public string ValidatorID
        {
            get;
            set;
        }

        // Zum programatischen Setzen
        public List<BaseValidator> Validators
        {
            get { return validators; }
        }

        public string ValidationErrorTextPrefix
        {
            get
            {
                return validationErrorTextPrefix;
            }
            set
            {
                validationErrorTextPrefix = value;
            }
        }

        public string ValidationErrorTextSuffix
        {
            get
            {
                return validationErrorTextSuffix;
            }
            set
            {
                validationErrorTextSuffix = value;
            }
        }

        public string ErrorCssClass
        {
            get
            {
                return errorCssClass;
            }
            set
            {
                errorCssClass = value;
            }
        }


        protected override void OnPreRender(EventArgs e)
        {
            base.OnPreRender(e);
            UpdateByValidators();
        }

        public void SetValidationErrorAttributes(string textPrefix = DefaultValidationErrorTextPrefix, string textSuffix = null, string errorCssClass = DefaultErrorCssClass)
        {
            if (!string.IsNullOrEmpty(textPrefix))
                validationErrorTextPrefix = textPrefix;
            if (!string.IsNullOrEmpty(textSuffix))
                validationErrorTextSuffix = textSuffix;
            if (!string.IsNullOrEmpty(errorCssClass))
                validationErrorTextPrefix = errorCssClass;
        }

        public void UpdateByValidators()
        {
            RemoveErronousClass();

            if (!string.IsNullOrEmpty(ValidatorID))
            {
                Control validator = this.NamingContainer.FindControl(ValidatorID);
                if (validator != null && validator is BaseValidator)
                {
                    if (!((BaseValidator)validator).IsValid)
                    {
                        MarkAsErronous();
                    }
                }
            }
            else if (this.validators != null && this.validators.Count > 0)
            {
                foreach (BaseValidator validator in this.validators)
                {
                    if (!validator.IsValid)
                    {
                        MarkAsErronous();
                        break;
                    }
                }
            }
        }

        public void MarkAsErronous()
        {
            // Prefix und Suffix nur dran hängen, wenn diese noch nicht vorhanden sind
            string prefix = (validationErrorTextPrefix == null ? "" : validationErrorTextPrefix);
            string suffix = (validationErrorTextSuffix == null ? "" : validationErrorTextSuffix);
            if (!string.IsNullOrEmpty(prefix) && Text.IndexOf(prefix) != 0)
            {
                Text = prefix + Text;
            }
            if (!string.IsNullOrEmpty(suffix) && Text.IndexOf(suffix) != Text.Length - suffix.Length)
            {
                Text = Text + suffix;
            }

            if (!string.IsNullOrEmpty(errorCssClass))
            {
                if (string.IsNullOrEmpty(CssClass))
                {
                    CssClass = errorCssClass;
                }
                else
                {
                    if (CssClass.IndexOf(errorCssClass) == -1)
                    {
                        CssClass += errorCssClass;
                    }
                }
            }
        }

        public void RemoveErronousClass()
        {
            string prefix = (validationErrorTextPrefix == null ? "" : validationErrorTextPrefix);
            string suffix = (validationErrorTextSuffix == null ? "" : validationErrorTextSuffix);

            // Prefix und Suffix entfernen, wenn diese hinzugefügt wurden
            if (!string.IsNullOrEmpty(prefix) && this.Text.IndexOf(prefix) == 0)
                this.Text = this.Text.Remove(0, prefix.Length);
            if (!string.IsNullOrEmpty(suffix))
            {
                int suffixIndex = this.Text.IndexOf(suffix);
                int suffixStartIndex = this.Text.Length - suffix.Length;
                if (suffixIndex == suffixStartIndex)
                    this.Text = this.Text.Remove(suffixStartIndex);
            }
            
            // CSS-Klassen bereinigen
            if (!string.IsNullOrEmpty(this.CssClass))
            {
                if (!string.IsNullOrEmpty(errorCssClass))
                {
                    List<string> classes = this.CssClass.Split(' ').ToList<string>();
                    for (int i = 0; i < classes.Count; i++)
                    {
                        if (classes[i] == errorCssClass)
                        {
                            classes.RemoveAt(i);
                            break;
                        }
                    }
                    this.CssClass = string.Join(" ", classes);
                }
            }
        }
    }
}