using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Collections;

/// <summary>
/// Zusammenfassungsbeschreibung für TMultiUser
/// </summary>
public class TMultiReport
{
    public string Filename;				
    public string DisplayName;	    	// Anzeigename des aktuellen Benutzers
    public int OrgID;
    public int FileType;
    public int FileLogo; 

    public TMultiReport()
    {
    }
    protected string checkUpdate(string aValue, string aField, string aUpdate, ref TParameterList parameterList)
    {
        parameterList.addParameter(aField, "string", aValue);

        if (aValue == "*")
            return aUpdate;
        else
            if (aUpdate == "")
                return aField + "=@" + aField;
            else
                return aUpdate + ", " + aField + "=@" + aField;
    }
    protected string checkUpdate(int aValue, string aField, string aUpdate, string aInvalid, ref TParameterList parameterList)
    {
        parameterList.addParameter(aField, "int", aValue.ToString());

        if (aValue.ToString() == aInvalid)
            return aUpdate;
        else
            if (aUpdate == "")
                return aField + "=@" + aField;
            else
                return aUpdate + ", " + aField + "=@" + aField;
    }
    public void update(string aProjectID, ArrayList aTrustList)
    {
        SqlDB dataReader;

        // Update-String konstruieren
        string tempUpdate = "";
        TParameterList parameterList = new TParameterList();
        tempUpdate = checkUpdate(DisplayName, "displayName", tempUpdate, ref parameterList);
        tempUpdate = checkUpdate(OrgID, "orgID", tempUpdate, "-1", ref parameterList);
        tempUpdate = checkUpdate(FileType, "fileType", tempUpdate, "-1", ref parameterList);
        tempUpdate = checkUpdate(FileLogo, "fileLogo", tempUpdate, "-1", ref parameterList);
        parameterList.addParameter("ID", "int", "");

        // Schleife über alle UserIDs
        foreach (string tempTrust in aTrustList)
        {
            parameterList.changeParameterValue("ID", tempTrust);
            // Datenschreiben
            if (tempUpdate != "")
            {
                dataReader = new SqlDB(aProjectID);
                dataReader.execSQLwithParameter("UPDATE download_reports SET " + tempUpdate + " WHERE ID=@ID", parameterList);
            }
        }
    }
}

