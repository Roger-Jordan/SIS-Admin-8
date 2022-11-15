using System.IO;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text.RegularExpressions;

namespace SIS.Components
{
    // Beispiel für Nutzung:
    // <div class="input-group">
    //     <sis:IconLabel runat="server" AssociatedControlID="EdtLostUsernameEmail" Icon="email"></sis:IconLabel>
    //     <asp:TextBox ID="EdtLostUsernameEmail" runat="server" CssClass="form-control"></asp:TextBox>
    // </div>  
    public class IconLabel : Label
    {
        private string _icon = null;

        public string Icon
        {
            get { return this._icon; }
            set {
                this._icon = value;
                UpdateCssClasses();
            }
        }

        private void UpdateCssClasses()
        {
            this.CssClass = "input-group-addon si-left no-padding";
            if (!string.IsNullOrEmpty(this._icon))
            {
                this.CssClass += " si-" + this._icon;
            }
        }

        // Sorgt dafür, dass auch Buttons mit einem Icon sauber gerendert werden
        //protected override void AddParsedSubObject(object obj)
        //{
        //    var literal = obj as LiteralControl;
        //    if (literal == null) return;
        //    Text = literal.Text;
        //}
    }
}