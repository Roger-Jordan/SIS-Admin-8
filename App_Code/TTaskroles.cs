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
/// Zusammenfassungsbeschreibung für TStructure
/// </summary>
public class TTaskroles
{
    public int roleID;			            // eindeutige ID des Containers
    public string displayName;	    	// Anzeigetext des Containers

    public TTaskroles()
    {
    }
    public static ListItemCollection getRoleList()
    {
        ListItemCollection result = new ListItemCollection();

        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        SqlDB dataReader;
        dataReader = new SqlDB("SELECT roleID, displayName1 FROM teamspace_roles", SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
            ListItem tempItem = new ListItem(dataReader.getString(1), dataReader.getInt32(0).ToString());
            result.Add(tempItem);
        }
        dataReader.close();

        return result;
    }
}