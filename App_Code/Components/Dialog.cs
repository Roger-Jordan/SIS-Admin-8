using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace SIS.Components
{
    public enum DialogSize
    {
        small,  // modal-sm
        normal,
        large   // modal-lg
    }


    public static class DialogSizeMethods
    {
        public static string GetString(this DialogSize size)
        {
            switch (size)
            {
                case DialogSize.small:
                    return "modal-sm";

                case DialogSize.normal:
                    return "";

                case DialogSize.large:
                    return "modal-lg";

                default:
                    return "";
            }
        }
    }


    [ParseChildren(true, "Text")]
    public class DialogHeaderTitle : WebControl
    {
        [DefaultValue("")]
        [Description("Überschrift eines Dialogs")]
        public string Text
        {
            get;
            set;
        }
        [DefaultValue("")]
        [Description("Unterüberschrift eines Dialogs")]
        public string Subtitle
        {
            get;
            set;
        }

        public DialogHeaderTitle() : base()
        {
            this.Text = "";
        }

        protected override void AddParsedSubObject(object obj)
        {
            if (obj is LiteralControl)
            {
                var literal = obj as LiteralControl;
                if (literal == null) return;
                Text = literal.Text;
            }
        }

        protected override void Render(HtmlTextWriter writer)
        {
            // Der Text wird zum aria-label
            // <sis:DialogTitle runat="server" ID="BtnCloseXYZ">Titel</sis:DialogTitle>

            // <h4 id="dialog-delete-contact" class="modal-title"><asp:Literal runat="server" ID="LblContactsDialogDeleteContactTitle">Ordner löschen</asp:Literal></h4>

            // Wichtig ist, dass hier die ID ausgegeben wird, da der Dialog darauf verweist
            writer.AddAttribute(HtmlTextWriterAttribute.Id, this.ClientID);
            writer.AddAttribute(HtmlTextWriterAttribute.Class, "modal-title");
            writer.RenderBeginTag(HtmlTextWriterTag.H4);
            writer.Write(this.Text);
            writer.RenderEndTag(); // </h4>

            if (!string.IsNullOrEmpty(Subtitle))
            {
                writer.AddAttribute(HtmlTextWriterAttribute.Class, "modal-subtitle");
                writer.RenderBeginTag(HtmlTextWriterTag.H6);
                writer.Write(Subtitle);
                writer.RenderEndTag(); // </h6>
            }
        }
    }


    [ParseChildren(true, "Text")]
    public class DialogHeaderCloseButton : WebControl
    {
        [DefaultValue("")]
        [Description("ARIA-Label des Schließen-Buttons")]
        public string Text
        {
            get;
            set;
        }

        public DialogHeaderCloseButton() : base()
        {
            this.Text = "";
        }

        protected override void AddParsedSubObject(object obj)
        {
            if (obj is LiteralControl)
            {
                var literal = obj as LiteralControl;
                if (literal == null) return;
                Text = literal.Text;
            }
        }

        protected override void Render(HtmlTextWriter writer)
        {
            // Der Text wird zum aria-label
            // <Close runat="server" ID="BtnCloseXYZ">Schließen</Close>

            // <button
            //     id="LblContactsDialogDeleteContactCloseIconAriaLabel"
            //     type="button"
            //     class="close"
            //     data -dismiss="modal"
            //     aria-label="Schließen">
            //
            //     <span aria-hidden="true">&times;</span>
            //
            // </button>

            writer.AddAttribute(HtmlTextWriterAttribute.Id, this.ClientID);
            writer.AddAttribute(HtmlTextWriterAttribute.Type, "button");
            writer.AddAttribute(HtmlTextWriterAttribute.Class, "close");
            writer.AddAttribute("data-dismiss", "modal");
            if (!string.IsNullOrEmpty(this.Text))
            {
                writer.AddAttribute("aria-label", HttpUtility.HtmlEncode(this.Text));
            }
            writer.RenderBeginTag(HtmlTextWriterTag.Button);

            writer.AddAttribute("aria-hidden", "true");
            writer.RenderBeginTag(HtmlTextWriterTag.Span);
            writer.Write("&times;"); // entspricht einem X
            writer.RenderEndTag(); // </span>

            writer.RenderEndTag(); // </div>
        }
    }


    [ParseChildren(true)]
    public class DialogHeader : WebControl
    {
        private DialogHeaderTitle title;
        private DialogHeaderCloseButton closeButton;

        [PersistenceMode(PersistenceMode.InnerProperty)]
        public DialogHeaderTitle Title
        {
            get
            {
                return title;
            }
            set
            {
                this.Controls.Remove(title);
                title = value;
                if (title != null)
                    this.Controls.Add(title);
            }
        }

        [PersistenceMode(PersistenceMode.InnerProperty)]
        public DialogHeaderCloseButton CloseButton
        {
            get
            {
                return closeButton;
            }
            set
            {
                this.Controls.Remove(closeButton);
                closeButton = value;
                if (closeButton != null)
                    this.Controls.Add(closeButton);
            }
        }

        protected override void Render(HtmlTextWriter writer)
        {
            // <sis:DialogHeader>
            //     <CloseButton runat="server" ID="BtnCloseXYZ">Schließen</Close>
            //     <Title></Title>
            // </sis:DialogHeader>

            // <div class="modal-header">
            //     [Button]
            //     [Title]
            // </div>

            writer.AddAttribute(HtmlTextWriterAttribute.Class, "modal-header");
            writer.RenderBeginTag(HtmlTextWriterTag.Div);

            if (CloseButton != null)
            {
                CloseButton.RenderControl(writer);
            }
            if (Title != null)
            {
                Title.RenderControl(writer);
            }

            writer.RenderEndTag(); // div.modal-header
        }
    }


    [ParseChildren()]
    public class DialogBody : WebControl
    {
        protected override string TagName
        {
            get
            {
                return "div";
            }
        }

        public override void RenderBeginTag(HtmlTextWriter writer)
        {
            writer.Write("<" + this.TagName + " class=\"modal-body\">");
        }

        public override void RenderEndTag(HtmlTextWriter writer)
        {
            writer.Write("</" + this.TagName + ">");
        }
    }


    public class DialogFooterCloseButton : SIS.Components.Button
    {
        public override void RenderBeginTag(System.Web.UI.HtmlTextWriter writer)
        {
            // Fügt auf leider sehr umständliche Weise das Attribut data-dismiss="modal" ein
            StringWriter stringWriter = new StringWriter();
            HtmlTextWriter tempHtmlWriter = new HtmlTextWriter(stringWriter);
            base.RenderBeginTag(tempHtmlWriter);
            string beginTagHtml = stringWriter.ToString();
            beginTagHtml = beginTagHtml.Insert(beginTagHtml.Length - 1, " data-dismiss=\"modal\"");
            writer.Write(beginTagHtml);
        }
    }


    [ParseChildren()]
    public class DialogFooter : Panel
    {
        protected override string TagName
        {
            get
            {
                return "div";
            }
        }

        public override void RenderBeginTag(HtmlTextWriter writer)
        {
            writer.Write("<" + this.TagName + " class=\"modal-footer\">");
        }

        public override void RenderEndTag(HtmlTextWriter writer)
        {
            writer.Write("</" + this.TagName + ">");
        }
    }


    [ParseChildren(true)]
    public class Dialog : WebControl
    {
        private DialogHeader header;
        private DialogBody body;
        private DialogFooter footer;

        public DialogSize Size { get; set; }

        [PersistenceMode(PersistenceMode.InnerProperty)]
        public DialogHeader Header
        {
            get
            {
                return header;
            }
            set
            {
                this.Controls.Remove(header);
                header = value;
                if (header != null)
                    this.Controls.Add(header);
            }
        }

        [PersistenceMode(PersistenceMode.InnerProperty)]
        public DialogBody Body
        {
            get
            {
                return body;
            }
            set
            {
                this.Controls.Remove(body);
                body = value;
                if (body != null)
                    this.Controls.Add(body);
            }
        }

        [PersistenceMode(PersistenceMode.InnerProperty)]
        public DialogFooter Footer
        {
            get
            {
                return footer;
            }
            set
            {
                this.Controls.Remove(footer);
                footer = value;
                if (footer != null)
                    this.Controls.Add(footer);
            }
        }

        public Dialog() : base()
        {
            Size = DialogSize.normal;
        }

        protected override void Render(HtmlTextWriter writer)
        {
            //<div id="contacts-delete-confirmation-dialog" class="modal fade" tabindex="-1" role="dialog" aria-labelledby="">
            //    <div class="modal-dialog" role="document">
            //        <div class="modal-content">
            //            <div class="modal-header"></div>
            //            <div class="modal-body"></div>
            //            <div class="modal-footer"></div>
            //        </div>
            //    </div>
            //</div>

            writer.AddAttribute(HtmlTextWriterAttribute.Id, this.ClientID);
            writer.AddAttribute(HtmlTextWriterAttribute.Class, "modal fade" + (!string.IsNullOrEmpty(this.CssClass) ? " " + this.CssClass : ""));
            writer.AddAttribute(HtmlTextWriterAttribute.Tabindex, "-1");
            writer.AddAttribute("role", "dialog");
            // ARIA-Referenz auf den Titel des Dialogs
            if (Header != null && Header.Title != null)
            {
                writer.AddAttribute("aria-labelledby", Header.Title.ClientID);
            }          
            writer.RenderBeginTag(HtmlTextWriterTag.Div);

            string dialogClass = "modal-dialog";
            dialogClass += " " + Size.GetString();
            writer.AddAttribute(HtmlTextWriterAttribute.Class, dialogClass);
            writer.AddAttribute("role", "document");
            writer.RenderBeginTag(HtmlTextWriterTag.Div);

            writer.AddAttribute(HtmlTextWriterAttribute.Class, "modal-content");
            writer.RenderBeginTag(HtmlTextWriterTag.Div);

            
            if (Header != null)
            {
                Header.RenderControl(writer);
            }
            if (Body != null)
            {
                Body.RenderControl(writer);
            }
            if (Footer != null)
            {
                Footer.RenderControl(writer);
            }


            writer.RenderEndTag(); // div.modal-content
            writer.RenderEndTag(); // div.modal-dialog
            writer.RenderEndTag(); // div.modal
        }
    }
}