using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;


namespace SIS.Components
{
    public class PureFileUpload : FileUpload
    {
        private string uploadUrl = null;// = "ContactsImport.aspx";
        private string icon = null;
        private string description = null;
        private string submitButton = null;

        private int maxFileSize = -1; // Bytes, -1 -> Prüfung deaktiviert
        private int maxFilenameLength = -1; // -1 -> Prüfung deaktiviert
        private string allowedFileExtensionsPrefix = null;
        private string allowedFileExtensions = null; // Komma-separierter String
        private bool showAllowedFileExtensionsPrefix = true;

        private bool autoStartUploadAfterFileSelection = false;
        private bool hideDropZoneOnUpload = false;
        private bool showThrobberOnSubmitButton = true;
        private string throbberCssClass = "throbber-white";

        private string errorMessageInvalidFileSize = null;
        private string errorMessageInvalidFileExtension = null;
        private string errorMessageNoDirectoryAllowed = null;
        private string errorMessageInvalidFileNameLength = null;
        private string errorMessageErrorOnResponse = null;

        private string onInitJavaScript = null;
        private string onChangedJavaScript = null;
        private string onBeforeUploadJavaScript = null;
        private string onUploadSuccessJavaScript = null;
        private string onUploadFailedJavaScript = null;
        private string onUploadAbortedJavaScript = null;


        public string UploadUrl
        {
            get { return uploadUrl; }
            set { uploadUrl = value; }
        }

        public string Icon
        {
            get { return icon; }
            set { icon = value; }
        }

        public string Description
        {
            get { return description; }
            set { description = value; }
        }

        [IDReferenceProperty]
        public string SubmitButton
        {
            get { return submitButton; }
            set { submitButton = value; }
        }

        public int MaxFileSize
        {
            get { return maxFileSize; }
            set { maxFileSize = value; }
        }

        public int MaxFilenameLength
        {
            get { return maxFilenameLength; }
            set { maxFilenameLength = value; }
        }

        public string AllowedFileExtensionsPrefix
        {
            get { return allowedFileExtensionsPrefix; }
            set { allowedFileExtensionsPrefix = value; }
        }

        public string AllowedFileExtensions
        {
            get { return allowedFileExtensions; }
            set { allowedFileExtensions = value; }
        }

        public bool ShowAllowedFileExtensionsPrefix
        {
            get { return showAllowedFileExtensionsPrefix; }
            set { showAllowedFileExtensionsPrefix = value; }
        }

        public bool AutoStartUploadAfterFileSelection
        {
            get { return autoStartUploadAfterFileSelection; }
            set { autoStartUploadAfterFileSelection = value; }
        }

        public bool HideDropZoneOnUpload
        {
            get { return hideDropZoneOnUpload; }
            set { hideDropZoneOnUpload = value; }
        }

        public bool ShowThrobberOnSubmitButton
        {
            get { return showThrobberOnSubmitButton; }
            set { showThrobberOnSubmitButton = value; }
        }

        public string ThrobberCssClass
        {
            get { return throbberCssClass; }
            set { throbberCssClass = value; }
        }

        public string ErrorMessageInvalidFileSize
        {
            get { return errorMessageInvalidFileSize; }
            set { errorMessageInvalidFileSize = value; }
        }

        public string ErrorMessageInvalidFileExtension
        {
            get { return errorMessageInvalidFileExtension; }
            set { errorMessageInvalidFileExtension = value; }
        }

        public string ErrorMessageNoDirectoryAllowed
        {
            get { return errorMessageNoDirectoryAllowed; }
            set { errorMessageNoDirectoryAllowed = value; }
        }

        public string ErrorMessageInvalidFileNameLength
        {
            get { return errorMessageInvalidFileNameLength; }
            set { errorMessageInvalidFileNameLength = value; }
        }

        public string ErrorMessageErrorOnResponse
        {
            get { return errorMessageErrorOnResponse; }
            set { errorMessageErrorOnResponse = value; }
        }



        public string OnInitJavaScript
        {
            get { return onInitJavaScript; }
            set { onInitJavaScript = value; }
        }

        public string OnChangedJavaScript
        {
            get { return onChangedJavaScript; }
            set { onChangedJavaScript = value; }
        }

        public string OnBeforeUploadJavaScript
        {
            get { return onBeforeUploadJavaScript; }
            set { onBeforeUploadJavaScript = value; }
        }

        public string OnUploadSuccessJavaScript
        {
            get { return onUploadSuccessJavaScript; }
            set { onUploadSuccessJavaScript = value; }
        }

        public string OnUploadFailedJavaScript
        {
            get { return onUploadFailedJavaScript; }
            set { onUploadFailedJavaScript = value; }
        }

        public string OnUploadAbortedJavaScript
        {
            get { return onUploadAbortedJavaScript; }
            set { onUploadAbortedJavaScript = value; }
        }
        


        public bool IsUploadAvailable()
        {
            return (this.Page != null)
                && (this.Page.Request != null)
                && (this.Page.Request.Files.Count > 0)
                && (this.Page.Request.Files[UniqueID] != null)
                && (!string.IsNullOrEmpty(this.Page.Request.Files[UniqueID].FileName));
        }

        public HttpPostedFile GetUploadedFile()
        {
            if (this.Page == null || this.Page.Request == null)
            {
                return null;
            }
            else
            {
                HttpPostedFile uploadedFile = this.Page.Request.Files[UniqueID];
                return uploadedFile;
            }

        }

        private string GenerateClientScript()
        {
            StringBuilder script = new StringBuilder();

            script.AppendFormat("$('#{0}').pureFileUpload({{", this.ClientID);

            script.AppendFormat("url: '{0}',", uploadUrl);
            if (!string.IsNullOrEmpty(icon))
                script.AppendFormat("icon: '{0}',", icon);
            if (!string.IsNullOrEmpty(description))
                script.AppendFormat("description: '{0}',", description);
            
            if (!string.IsNullOrEmpty(submitButton))
            {
                Control targetControl = this.NamingContainer.FindControl(submitButton);
                script.AppendFormat("submitButtonId: '#{0}',", targetControl.ClientID);
            }

            if (maxFileSize > -1)
                script.AppendFormat("maxFileSize: {0},", maxFileSize);
            if (maxFilenameLength > -1)
                script.AppendFormat("maxFilenameLength: {0},", maxFilenameLength);
            if (!string.IsNullOrEmpty(allowedFileExtensionsPrefix))
                script.AppendFormat("allowedFileExtensionsPrefix: '{0}',", allowedFileExtensionsPrefix);
            script.AppendFormat("allowedFileExtensions: [{0}],", GetImplodedAllowedFileExtensions());
            script.AppendFormat("showAllowedFileExtensionsPrefix: {0},", showAllowedFileExtensionsPrefix.ToString().ToLower());

            script.AppendFormat("autoStartUploadAfterFileSelection: {0},", autoStartUploadAfterFileSelection.ToString().ToLower());
            script.AppendFormat("hideDropZoneOnUpload: {0},", hideDropZoneOnUpload.ToString().ToLower());

            List<string> errorMessagesParams = new List<string>();
            if (!string.IsNullOrEmpty(errorMessageInvalidFileSize))
                errorMessagesParams.Add(string.Format("invalidFileSize: '{0}'", errorMessageInvalidFileSize));
            if (!string.IsNullOrEmpty(errorMessageInvalidFileExtension))
                errorMessagesParams.Add(string.Format("invalidFileExtension: '{0}'", errorMessageInvalidFileExtension));
            if (!string.IsNullOrEmpty(errorMessageNoDirectoryAllowed))
                errorMessagesParams.Add(string.Format("noDirectoryAllowed: '{0}'", errorMessageNoDirectoryAllowed));
            if (!string.IsNullOrEmpty(errorMessageInvalidFileNameLength))
                errorMessagesParams.Add(string.Format("invalidFileNameLength: '{0}'", errorMessageInvalidFileNameLength));
            if (!string.IsNullOrEmpty(errorMessageErrorOnResponse))
                errorMessagesParams.Add(string.Format("errorOnResponse: '{0}'", errorMessageErrorOnResponse));
            if (errorMessagesParams.Count > 0)
            {
                script.Append("errorMessage: {");
                script.Append(string.Join(", ", errorMessagesParams));
                script.Append("},");
            }

            string additionOnBeforeUploadJavaScript = null;
            string additionOnUploadFailedJavaScript = null;
            string additionOnUploadAbortedJavaScript = null;
            if (showThrobberOnSubmitButton && !string.IsNullOrEmpty(throbberCssClass) && !string.IsNullOrEmpty(submitButton))
            {
                Control targetControl = this.NamingContainer.FindControl(submitButton);
                additionOnBeforeUploadJavaScript = string.Format("$('#{0}').addClass('{1}');", targetControl.ClientID, throbberCssClass);
                additionOnUploadFailedJavaScript = string.Format("$('#{0}').removeClass('{1}');", targetControl.ClientID, throbberCssClass);
                additionOnUploadAbortedJavaScript = string.Format("$('#{0}').removeClass('{1}');", targetControl.ClientID, throbberCssClass);
            }
            if (!string.IsNullOrEmpty(onBeforeUploadJavaScript) || !string.IsNullOrEmpty(additionOnBeforeUploadJavaScript))
                script.AppendFormat("onBeforeUpload: {0},", CreateOrMergeJavaScriptCallback(onBeforeUploadJavaScript, additionOnBeforeUploadJavaScript));

            if (!string.IsNullOrEmpty(onUploadFailedJavaScript) || !string.IsNullOrEmpty(additionOnUploadFailedJavaScript))
                script.AppendFormat("onUploadFailed: {0},", CreateOrMergeJavaScriptCallback(onUploadFailedJavaScript, additionOnUploadFailedJavaScript));

            if (!string.IsNullOrEmpty(onUploadAbortedJavaScript) || !string.IsNullOrEmpty(additionOnUploadAbortedJavaScript))
                script.AppendFormat("onUploadAborted: {0},", CreateOrMergeJavaScriptCallback(onUploadAbortedJavaScript, additionOnUploadAbortedJavaScript));
            
            if (!string.IsNullOrEmpty(onInitJavaScript))
                script.AppendFormat("onInit: {0},", onInitJavaScript);
            if (!string.IsNullOrEmpty(onChangedJavaScript))
                script.AppendFormat("onChanged: {0},", onChangedJavaScript);
            if (!string.IsNullOrEmpty(onUploadSuccessJavaScript))
                script.AppendFormat("onUploadSuccess: {0},", onUploadSuccessJavaScript);

            script.Append("});");

            return script.ToString();
        }

        protected override void Render(HtmlTextWriter writer)
        {
            Assets.RegisterJavaScriptFile(this.Page, Assets.JavaScriptFile.PureFileUpload);

            string clientScript = GenerateClientScript();
            Assets.RegisterJavaScript(this.Page, this.ClientID, clientScript);

            writer.AddAttribute("id", this.ClientID);
            writer.AddAttribute("name", this.UniqueID);
            writer.AddAttribute("type", "file");
            writer.RenderBeginTag(HtmlTextWriterTag.Input);
            writer.RenderEndTag();
        }

        private string GetImplodedAllowedFileExtensions()
        {
            List<string> extensions = GetAllowedFileExtensions();
            if (extensions.Count == 0)
            {
                return "";
            }
            else
            {
                return "'" + string.Join("', '", extensions) + "'";
            }
        }

        private List<string> GetAllowedFileExtensions()
        {
            List<string> allowedExtensions = new List<string>();
            // allowedFileExtensions: [], // Array von Extensions, ohne(!) Punkt; leeres Array, wenn alle Erweiterungen erlaubt sind
            if (!string.IsNullOrEmpty(allowedFileExtensions))
            {
                foreach (string extension in allowedFileExtensions.Split(','))
                {
                    allowedExtensions.Add(extension.Trim().ToLower());
                }
            }
            return allowedExtensions;
        }

        private string CreateOrMergeJavaScriptCallback(string userDefinedJavaScript, string additionalJavaScript)
        {
            string result;
            if (string.IsNullOrEmpty(additionalJavaScript))
            {
                result = userDefinedJavaScript;
            }
            else if (string.IsNullOrEmpty(userDefinedJavaScript))
            {
                result = string.Format("function(args) {{ {0} }}", additionalJavaScript);
            }
            else
            {
                if (userDefinedJavaScript.IndexOf("function") == 0)
                {
                    // Wenn ganz am Anfang ein "function" steht, dann handelt es sich um eine Inline-Funktion und wir können
                    // direkt nach dem ersten { das additionalJavaScript einfügen, sodass es als erster Befehl ausgeführt wird.
                    // Im Anschluss folgt dann automatisch die Ausführung des benutzerdefinierten Codes
                    result = userDefinedJavaScript.Insert(userDefinedJavaScript.IndexOf("{") + 1, additionalJavaScript);
                }
                else
                {
                    // Wenn kein "function" am Anfang steht, dann handelt es sich wohl um eine globale Methode, die aufgerufen
                    // werden soll. Dann wir das ganz in eine Funktion gekapselt und entsprechend augerufen
                    result = string.Format("function(args) {{ {0} {1}(args); }}", additionalJavaScript, userDefinedJavaScript);
                }
            }

            return result;
        }

    }
}