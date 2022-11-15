using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace SIS.Components
{
    public class IconedTextBox : TextBox
    {
        public string Icon { get; set; }


        protected override void Render(HtmlTextWriter writer)
        {
            // Wenn das Icon leer ist, dann wird eine Standard-Textbox ausgegeben
            if (string.IsNullOrEmpty(Icon))
            {
                base.Render(writer);
            }
            else
            {
                writer.AddAttribute(HtmlTextWriterAttribute.Class, "input-group");
                writer.RenderBeginTag(HtmlTextWriterTag.Div);

                writer.AddAttribute(HtmlTextWriterAttribute.Class, "input-group-addon si-left no-padding si-" + Icon);
                writer.AddAttribute(HtmlTextWriterAttribute.For, this.ClientID);
                writer.RenderBeginTag(HtmlTextWriterTag.Label);
                writer.RenderEndTag();

                base.Render(writer);

                writer.RenderEndTag();
            }
        }
    }
}