using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using SIS;


namespace SIS.Components
{
    public class Panel : System.Web.UI.WebControls.Panel
    {
        public bool Active
        {
            get;
            set;
        }

        public override void RenderBeginTag(HtmlTextWriter writer)
        {
            if (string.IsNullOrEmpty(this.CssClass))
                this.CssClass = "";
            else
                this.CssClass += " ";

            this.CssClass += "tab-pane fade";
            this.Attributes["role"] = "tabpanel";

            if (Active)
                this.CssClass += " active in";

            base.RenderBeginTag(writer);
        }
    }


    [ParseChildren(true, "Items")]
    public class Panels : WebControl
    {
        private ArrayList items = new ArrayList();

        public ArrayList Items
        {
            get
            {
                return items;
            }
        }


        public Panels() : base(HtmlTextWriterTag.Div) // <div>-Element
        {
        }

        protected override void OnInit(EventArgs e)
        {
            EnsureChildControls();
            base.OnInit(e);
        }

        protected override void CreateChildControls()
        {
            this.Controls.Clear();

            foreach (Panel panel in items)
            {
                this.Controls.Add(panel);
            }

            base.CreateChildControls();
        }

        public override void RenderBeginTag(HtmlTextWriter writer)
        {
            if (string.IsNullOrEmpty(this.CssClass))
                this.CssClass = "";
            else
                this.CssClass += " ";
            this.CssClass += "tab-content";

            base.RenderBeginTag(writer);
        }
    }


    [ParseChildren(false, "Text")]
    public class Tab : WebControl
    {
        private bool _active;

        [DefaultValue("")]
        [Description("Titel des Panels")]
        public string Text
        {
            get;
            set;
        }

        [System.ComponentModel.TypeConverter(typeof(AssociatedControlConverter))]
        [Description("ID des Ziel-Panels")]
        public string TargetPanelID
        {
            get;
            set;
        }

        public Panel TargetPanel
        {
            get
            {
                if (this.NamingContainer == null)
                    return null;

                Control control = this.NamingContainer.FindControl(TargetPanelID);
                if (!(control is Panel))
                    control = null;
                return (Panel)control;                
            }
        }

        public bool Active
        {
            get
            {
                return this._active;
            }
            set
            {
                this._active = value;
                if (TargetPanel != null)
                    TargetPanel.Active = this._active;
            }
        }


        public Tab() : base()
        {
            this.Text = "";
            this._active = false;
        }

        protected override void OnInit(EventArgs e)
        {
            EnsureChildControls();
            base.OnInit(e);
        }

        protected override void AddParsedSubObject(object obj)
        {
            if (obj != null && obj is Control)
            {
                this.Controls.Add((Control)obj);
                //if (obj is LiteralControl)
                //{
                //    var literal = obj as LiteralControl;
                //    Text = literal.Text;
                //}                
                //else if (obj is Label)
                //{
                //    var label = obj as Label;
                //    Text = label.Text;
                //}
            }
        }

        protected override void Render(HtmlTextWriter writer)
        {
            //<li id="Tab3" runat="server" role="presentation">
            //    <a href="#<%=Pan3.ClientID%>" aria-controls="<%=Pan3.ClientID%>" role="tab" data-toggle="tab">
            //        <asp:Label ID="LblEditTabColleagues" runat="server" Text="Kollegen" />
            //    </a>
            //</li>

            writer.AddAttribute(HtmlTextWriterAttribute.Id, this.ClientID);
            writer.AddAttribute("role", "presentation");
            if (this.Active)
                writer.AddAttribute("class", "active");
            writer.RenderBeginTag(HtmlTextWriterTag.Li);

            writer.AddAttribute("href", "#" + TargetPanel.ClientID);
            writer.AddAttribute("data-toggle", "tab");
            writer.AddAttribute("aria-controls", TargetPanel.ClientID);
            writer.AddAttribute("role", "tab");
            writer.RenderBeginTag(HtmlTextWriterTag.A);
            foreach (Control control in this.Controls)
                control.RenderControl(writer);
            writer.RenderEndTag(); // </a>

            writer.RenderEndTag(); // </li>
        }
    }


    [ParseChildren(true, "Items")]
    public class Tabs : WebControl
    {
        private ArrayList items = new ArrayList();

        public ArrayList Items
        {
            get
            {
                return items;
            }
        }


        public Tabs() : base(HtmlTextWriterTag.Ul) // <ul>-Element
        {
        }

        protected override void OnInit(EventArgs e)
        {
            EnsureChildControls();
            base.OnInit(e);
        }

        protected override void CreateChildControls()
        {
            this.Controls.Clear();

            foreach (Tab tab in items)
            {
                this.Controls.Add(tab);
            }

            base.CreateChildControls();
        }

        public override void RenderBeginTag(System.Web.UI.HtmlTextWriter writer)
        {
            RegisterJavaScript();

            if (string.IsNullOrEmpty(this.CssClass))
                this.CssClass = "";
            else
                this.CssClass += " ";
            this.CssClass += "nav nav-tabs";

            this.Attributes["role"] = "tablist";

            base.RenderBeginTag(writer);
        }

        protected void RegisterJavaScript()
        {
            string script = "$('#" + this.ClientID + " a').click(function(e) { e.preventDefault(); $(this).tab('show'); });";
            Assets.RegisterJavaScript(this.Page, this.ClientID + "_BSTab", script);
        }

        public void ActivateTab(Tab tab)
        {
            foreach (Tab tempTab in items)
            {
                tempTab.Active = (tempTab == tab);
            }
        }

        public bool HasActiveTab()
        {
            foreach (Tab tempTab in items)
            {
                if (tempTab.Active)
                    return true;
            }
            return false;
        }
    }
}