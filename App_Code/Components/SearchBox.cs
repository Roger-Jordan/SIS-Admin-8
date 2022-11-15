using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Security.Permissions;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;


namespace SIS.Components
{
    public enum SearchBoxSize
    {
        Normal,
        Small
    }


    public class SearchBox : CompositeControl
    {
        private System.Web.UI.WebControls.Panel inputGroup;
        private IconLabel iconLabel;
        private HyperLink clearButton;
        private TextBox textBox;
        private Button submitButton;

        private string tableSelector = null;
        private bool clientSideSearch = true;
        private string icon = null;
        private string placeholder = null;
        private SearchBoxSize size = SearchBoxSize.Normal;
        private bool disableClearButton = false;
        private bool showThrobberOnSearch = true;

        private static readonly object EventSubmitKey = new object();


        public string TableSelector
        {
            get { return tableSelector; }
            set { tableSelector = value; }
        }

        public bool ClientSideSearch
        {
            get { return clientSideSearch; }
            set { clientSideSearch = value; }
        }

        public string Icon
        {
            get { return icon; }
            set { icon = value; }
        }

        public string Placeholder
        {
            get
            {
                if (textBox != null)
                {
                    return textBox.Attributes["placeholder"];
                }
                return placeholder;
            }
            set {
                placeholder = value;
                if (textBox != null)
                {
                    textBox.Attributes["placeholder"] = placeholder;
                }
            }
        }

        public SearchBoxSize Size
        {
            get { return size; }
            set { size = value; }
        }

        public bool DisableClearButton
        {
            get { return disableClearButton; }
            set { disableClearButton = value; }
        }

        public bool ShowThrobberOnSearch
        {
            get { return showThrobberOnSearch; }
            set { showThrobberOnSearch = value; }
        }

        public int MaxLength
        {
            get { return textBox.MaxLength; }
            set { textBox.MaxLength = value; }
        }

        public string SearchText
        {
            get { return textBox.Text; }
            set { textBox.Text = value; }
        }

        public event EventHandler Submit
        {
            add { Events.AddHandler(EventSubmitKey, value); }
            remove { Events.RemoveHandler(EventSubmitKey, value); }
        }


        protected override void RecreateChildControls()
        {
            EnsureChildControls();
        }

        protected override void CreateChildControls()
        {
            Controls.Clear();

            GenerateInputGroup();
            GenerateClearButton();
            GenerateTextBox();
            GenerateIconLabel(); // Muss nach der TextBox erzeugt werden, da das Icon auf die ID der TextBox zugreift
            GenerateSubmitButton();

            if (iconLabel != null)
                inputGroup.Controls.Add(iconLabel);
            if (clearButton != null)
                inputGroup.Controls.Add(clearButton);
            inputGroup.Controls.Add(textBox);
            if (submitButton != null)
                inputGroup.Controls.Add(submitButton);

            this.Controls.Add(inputGroup);
        }

        private void GenerateInputGroup()
        {
            //<div runat = "server" ID="PanStructureTableSearchBox" class="search-box search-box-sm client-side-search input-group">

            inputGroup = new System.Web.UI.WebControls.Panel();
            inputGroup.ID = this.ID + "_SearchBoxInputGroup";
            inputGroup.CssClass = "search-box input-group";
            if (size == SearchBoxSize.Small)
                inputGroup.CssClass += " search-box-sm";
            if (clientSideSearch)
                inputGroup.CssClass += " client-side-search";
            if (showThrobberOnSearch)
                inputGroup.CssClass += " show-throbber-on-search";
        }

        private void GenerateIconLabel()
        {
            if (string.IsNullOrEmpty(icon))
                return;

            // <sis:IconLabel runat="server" AssociatedControlID="EdtStructureClientSideSearch" CssClass="input-group-addon"></sis:IconLabel>
            iconLabel = new IconLabel();
            iconLabel.ID = this.ID + "_SearchBoxIconLabel";
            iconLabel.Icon = icon;
            iconLabel.CssClass = "search-box-icon-label input-group-addon si-left si-" + icon;
            iconLabel.AssociatedControlID = textBox.ID;
        }

        private void GenerateClearButton()
        {
            if (disableClearButton)
                return;

            // <a href="javascript:void(0);" class="clear-search-box"></a>
            clearButton = new HyperLink();
            clearButton.ID = this.ID + "_SeachBoxClearButton";
            clearButton.NavigateUrl = "javascript:void(0);";
            clearButton.CssClass = "clear-search-box";
        }

        private void GenerateTextBox()
        {
            textBox = new TextBox();
            textBox.ID = this.ID + "_SearchBoxTextBox";
            textBox.CssClass = "form-control";
            textBox.Attributes.Add("data-target", tableSelector);
            if (!string.IsNullOrEmpty(placeholder))
            {
                textBox.Attributes.Add("placeholder", placeholder);
            }
        }

        private void GenerateSubmitButton()
        {
            if (clientSideSearch)
                return;

            submitButton = new Button();
            submitButton.ID = "SearchBoxHiddenSubmitButton";// + internalSearchBoxID.ToString();
            submitButton.Click += new EventHandler(_submitButton_click);
            submitButton.CssClass = "search-box-submit-button";
            submitButton.CausesValidation = false;
        }

        private void _submitButton_click(object source, EventArgs e)
        {
            OnSubmit(EventArgs.Empty);
        }

        private void OnSubmit(EventArgs e)
        {
            EventHandler SubmitHandler = (EventHandler)Events[EventSubmitKey];
            if (SubmitHandler != null)
            {
                SubmitHandler(this, e);
            }
        }

        protected override void Render(HtmlTextWriter writer)
        {
            Assets.RegisterJavaScriptFile(this.Page, Assets.JavaScriptFile.SearchBox);

            AddAttributesToRender(writer);

            inputGroup.RenderControl(writer);
        }

        // Diese Methode ist dafür da, wenn man ein Label zum Suchfeld einbauen will, kann man darüber die ID des Textfeldes
        // ermitteln und via AssociatedControlID das Label erknüpfen
        public string GetTextBoxClientID()
        {
            return textBox.ClientID;
        }

    }
}