using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace SIS.Components
{
    public class DetailsSidebar : Panel
    {
        private bool showClose = true;
        private bool showTitle = true;

        // Folgende Sektionen sind spezielle Sektionen, die automatisch befüllt werden. Dabei gilt es ein
        // paar Konventionen einzuhalten:
        // - meta-information:
        //       Diese Sektion enthält eine Key-Value-Darstellung von allen Spaltenüberchriften-Zellen-Darstellungen,
        //       bei denen die Spaltenüberschriften, also die <th>-Elemente mit der Klasse "sidebar-meta-information"
        //       ausgestattet sind.
        // - comment:
        //       In dieser Sektion wird eine Textarea dargestellt, die den Inhalt einer Zelle darstellt. Die entsprechende
        //       Spaltenüberschrift <th> muss die Klasse "sidebar-comment" versehen werden. Sind mehrere Spalten mit dieser
        //       Klasse vorhanden, so wird nur der Inhalt der Zelle zur der ersten Splate dargestellt. Die Überschrift der
        //       Textarea entspricht der zur Zelle zugehörigen Spaltenüberschrift.
        // - buttons:
        //       Buttons in einer Tabellen-Zeile, die in der Buttons-Sektion angezeigt werden sollen müssen ein <a>
        //       sein und eine der zwei Klassen "sidebar-button" oder "sidebar-primary-button" beinhalten
        private string sections = null;

        private string onBeforeOpenJavaScript = null;
        private string onOpenedJavaScript = null;
        private string onCloseJavaScript = null;


        public bool ShowTitle
        {
            get { return showTitle; }
            set { showTitle = value; }
        }

        public bool ShowClose
        {
            get { return showClose; }
            set { showClose = value; }
        }

        public string Sections
        {
            get { return sections; }
            set { sections = value; }
        }

        public string OnBeforeOpenJavaScript
        {
            get { return onBeforeOpenJavaScript; }
            set { onBeforeOpenJavaScript = value; }
        }

        public string OnOpenedJavaScript
        {
            get { return onOpenedJavaScript; }
            set { onOpenedJavaScript = value; }
        }

        public string OnCloseJavaScript
        {
            get { return onCloseJavaScript; }
            set { onCloseJavaScript = value; }
        }



        protected override void OnPreRender(EventArgs e)
        {
            base.OnPreRender(e);

            RegisterJavaScriptCallback("beforeOpen", onBeforeOpenJavaScript);
            RegisterJavaScriptCallback("opened", onOpenedJavaScript);
            RegisterJavaScriptCallback("close", onCloseJavaScript);
        }

        private void RegisterJavaScriptCallback(string eventName, string javaScriptCallback)
        {
            if (!string.IsNullOrEmpty(javaScriptCallback))
            {
                string script = string.Format("$('#{0}').on('details-sidebar:{1}', {2});", this.ClientID, eventName, javaScriptCallback);
                Assets.RegisterJavaScript(this.Page, this.ClientID + "_" + eventName, script);
            }
        }

        protected override void Render(HtmlTextWriter writer)
        {
            Assets.RegisterJavaScriptFile(this.Page, Assets.JavaScriptFile.DetailsSidebar);

            //<div class="details-sidebar">
            //    <div class="wrapper">
            //        <a href="javascript:void(0)" class="sidebar-close" title=""></a>
            //        <div class="sidebar-title">
            //            <h4></h4>
            //        </div>
            //        <section class="sidebar-section-meta-information"></section>
            //        <section class="sidebar-section-buttons"></section>
            //    </div>
            //</div>

            writer.AddAttribute(HtmlTextWriterAttribute.Class, "details-sidebar");
            writer.AddAttribute(HtmlTextWriterAttribute.Id, this.ClientID);
            writer.RenderBeginTag(HtmlTextWriterTag.Div);

            writer.AddAttribute(HtmlTextWriterAttribute.Class, "wrapper");
            writer.RenderBeginTag(HtmlTextWriterTag.Div);

            if (showClose)
            {
                writer.AddAttribute(HtmlTextWriterAttribute.Href, "javascript:void(0);");
                writer.AddAttribute(HtmlTextWriterAttribute.Class, "sidebar-close");
                writer.RenderBeginTag(HtmlTextWriterTag.A);
                writer.RenderEndTag();
            }

            if (showTitle)
            {
                writer.AddAttribute(HtmlTextWriterAttribute.Class, "sidebar-title");
                writer.RenderBeginTag(HtmlTextWriterTag.Div);

                writer.RenderBeginTag(HtmlTextWriterTag.H4);
                writer.RenderEndTag();

                writer.RenderEndTag();
            }

            if (!string.IsNullOrEmpty(sections))
            {
                foreach (string section in sections.Split(','))
                {
                    writer.AddAttribute(HtmlTextWriterAttribute.Class, "sidebar-" + section.Trim());
                    writer.RenderBeginTag("section");
                    writer.RenderEndTag(); ;
                }
            }            

            writer.RenderEndTag(); // .wrapper

            writer.RenderEndTag(); // .details-sidebar
        }
    }
}