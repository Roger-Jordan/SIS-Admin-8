using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;

/// <summary>
/// Klasse zur Ermittlung von Sprachspezifischen Texten aus der DB
/// </summary>
public class Labeling
{
    public class TFieldLabel
    {
        public string title;
        public string tooltipp;
        public string errorRegEx;
        public string errorMandatory;
        public string errorRange;
        public string errorMaxLength;
        public string errorInvalidValue;
        public string exportHintHeader;
        public string exportHint1;
        public string exportHint2;
        public string exportHint3;
    }
    
    /// <summary>
    /// Ermittlung eines Textes in einer Sprache
    /// </summary>
    /// <param name="name">ID des gesuchten Textes</param>
    /// <param name="language">Sprache</param>
    /// <param name="project">ID des Projektes</param>
    /// <returns></returns>
    public static string getLabel(string name, string language, string project)
    {
        string sTemp = "";
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("field", "string", name);
        parameterList.addParameter("language", "string", language);
        SqlDB dataReader = new SqlDB("SELECT text FROM dictionary WHERE field=@field AND language=@language",parameterList, "");
        if (dataReader.read())
        {
            sTemp = dataReader.getString(0);
        }
        dataReader.close();
        return sTemp;
    }
    /// <summary>
    /// Anpassen eines Controls an eine Sprache
    /// </summary>
    /// <param name="vater">Control das die anzupassenden Elemente enthält</param>
    /// <param name="language">Sprache</param>
    /// <param name="project">ID des Projektes</param>
    private static void LabelComponent(ref Control vater, string language, string project)
    {
        string s;
        if (vater.GetType().Name == "Button")
        {
            Button b = (Button)vater;
            s = getLabel(b.ID, language, project);
            if (s != "")
            {
                b.Text = s;
            }
        }

        if (vater.GetType().Name == "Label")
        {
            Label l = (Label)vater;
            s = getLabel(l.ID, language, project);
            if (s != "")
            {
                l.Text = s;
            }
        }

        if (vater.GetType().Name == "LinkButton")
        {
            LinkButton l = (LinkButton)vater;
            s = getLabel(l.ID, language, project);
            if (s != "")
            {
                l.Text = s;
            }
        }

        if (vater.GetType().Name == "Image")
        {
            Image i = (Image)vater;
            s = getLabel(i.ID, language, project);
            i.ToolTip = s;
            i.AlternateText = s;
        }

        if (vater.GetType().Name == "RadioButtonList")
        {
            RadioButtonList rbl = (RadioButtonList)vater;
            for (int n = 1; n <= rbl.Items.Count; n++)
            {
                s = getLabel(rbl.ID + n, language, project);
                if (s != "")
                {
                    rbl.Items[n - 1].Text = s;
                }
            }
        }

        ControlCollection cColl = vater.Controls;
        foreach (Control cont in cColl)
        {
            Control c = cont;
            LabelComponent(ref c, language, project);
        }
    }
    /// <summary>
    /// Anpassen aller Elemente innerhalb Seite an eine Sprache
    /// </summary>
    /// <param name="p">anzupassende Seite</param>
    /// <param name="language">Sprache</param>
    /// <param name="project">ID des Projektes</param>
    public static void LabelComponents(ref Page p, string language, string project)
    {
        ControlCollection cColl = p.Controls;
        foreach (Control cont in (ControlCollection)p.Controls)
        {
            Control c = cont;
            LabelComponent(ref c, language, project);
        }
    }
    /// <summary>
    /// Anpassen aller Elemente innerhalb Einer Masterpage an eine Sprache
    /// </summary>
    /// <param name="p">anzupassende Masterpage</param>
    /// <param name="language">Sprache</param>
    /// <param name="project">ID des Projektes</param>
    public static void LabelComponents(ref MasterPage p, string language, string project)
    {
        ControlCollection cColl = p.Controls;
        foreach (Control cont in (ControlCollection)p.Controls)
        {
            Control c = cont;
            LabelComponent(ref c, language, project);
        }
    }
    public static TFieldLabel getStructureFieldLabel(string aTable, string name, string language, string project)
    {
        TFieldLabel result = new TFieldLabel();

        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("fieldID", "string", name);
        parameterList.addParameter("language", "string", language);
        SqlDB dataReader = new SqlDB("SELECT title, tooltipp, errorRegEx, errorMandatory, errorRange, errorMaxLength, errorInvalidValue, exportHintHeader, exportHint1, exportHint2, exportHint3 FROM " + aTable + "dictionary WHERE fieldID=@fieldID AND language=@language", parameterList, project);
        if (dataReader.read())
        {
            result.title = dataReader.getString(0);
            result.tooltipp = dataReader.getString(1);
            result.errorRegEx = dataReader.getString(2);
            result.errorMandatory = dataReader.getString(3);
            result.errorRange = dataReader.getString(4);
            result.errorMaxLength = dataReader.getString(5);
            result.errorInvalidValue = dataReader.getString(6);
            result.exportHintHeader = dataReader.getString(7);
            result.exportHint1 = dataReader.getString(8);
            result.exportHint2 = dataReader.getString(9);
            result.exportHint3 = dataReader.getString(10);
        }
        dataReader.close();
        return result;
    }
    public static TFieldLabel getEmployeeFieldLabel(string name, string language, string project)
    {
        TFieldLabel result = new TFieldLabel();

        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("fieldID", "string", name);
        parameterList.addParameter("language", "string", language);
        SqlDB dataReader = new SqlDB("SELECT title, tooltipp, errorRegEx, errorMandatory, errorRange, errorMaxLength, errorInvalidValue, exportHintHeader, exportHint1, exportHint2, exportHint3  FROM orgmanager_employeefieldsdictionary WHERE fieldID=@fieldID AND language=@language", parameterList, project);
        if (dataReader.read())
        {
            result.title = dataReader.getString(0);
            result.tooltipp = dataReader.getString(1);
            result.errorRegEx = dataReader.getString(2);
            result.errorMandatory = dataReader.getString(3);
            result.errorRange = dataReader.getString(4);
            result.errorMaxLength = dataReader.getString(5);
            result.errorInvalidValue = dataReader.getString(6);
            result.exportHintHeader = dataReader.getString(7);
            result.exportHint1 = dataReader.getString(8);
            result.exportHint2 = dataReader.getString(9);
            result.exportHint3 = dataReader.getString(10);
        }
        dataReader.close();
        return result;
    }
}
