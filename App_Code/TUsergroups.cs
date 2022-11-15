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
/// Zusammenfassungsbeschreibung für TUser
/// </summary>
public class TUsergroups
{

    public ArrayList groupIDs;

    public TUsergroups()
    {
        groupIDs = new ArrayList();
    }
    public TUsergroups(string aProjectID)
    {
        groupIDs = new ArrayList();
        SqlDB dataReader;
        dataReader = new SqlDB("select groupID from loginusergroups", aProjectID);
        while (dataReader.read())
        {
            groupIDs.Add(dataReader.getString(0));
        }
        dataReader.close();
    }
    public static void delete(string aGroupID, string aProjectID)
    {
        SqlDB dataReader;
        // aus Datenbank löschen
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQL("DELETE FROM loginusergroups WHERE groupID='" + aGroupID + "'");
    }
    public static ListItemCollection getUsergroupList()
    {
        ListItemCollection result = new ListItemCollection();

        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        SqlDB dataReader;
        dataReader = new SqlDB("SELECT groupID FROM loginusergroups", SessionObj.Project.ProjectID);
        while (dataReader.read())
        {
            ListItem tempItem = new ListItem(dataReader.getString(0), dataReader.getString(0));
            result.Add(tempItem);
        }
        dataReader.close();

        return result;
    }
}

