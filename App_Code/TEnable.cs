using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

/// <summary>
/// Klasse zum Setzen der CSS-Klasse und Austauschen des Images bei Buttons für Aktiv-/Inaktiv-Status
/// </summary>
public class TEnable
{
    /// <summary>
    /// CSS-Klasse eines Buttons setzen
    /// </summary>
    /// <param name="aButton">Button dessen CSS-Klasse gesetzt werden soll</param>
    /// <param name="aValue">zu setzender Status</param>
    public static void setEnabled(System.Web.UI.WebControls.Button aButton, bool aValue)
    {
        aButton.Enabled = aValue;
        if (aValue)
            aButton.CssClass = "submit-button";
        else
            aButton.CssClass = "submit-button-disabled";
    }
    /// <summary>
    /// Image und CSS-Klasse eines Buttons setzen
    /// </summary>
    /// <param name="aButton">Button dessen CSS-Klasse gesetzt werden soll</param>
    /// <param name="aValue">zu setzender Status</param>
    /// <param name="aImageUrl">Dateiname (und Extension) des zu setzenden Images</param>
    public static void setEnabled(System.Web.UI.WebControls.ImageButton aButton, bool aValue, string aImageUrl)
    {
        aButton.Enabled = aValue;
        if (aValue)
            aButton.ImageUrl = "~/Images/Buttons/Enabled/" + aImageUrl;
        else
            aButton.ImageUrl = "~/Images/Buttons/Disabled/" +aImageUrl;
    }
}
