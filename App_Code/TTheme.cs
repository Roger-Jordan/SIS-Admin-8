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
/// Zusammenfassungsbeschreibung für Theme
/// </summary>
public class TTheme
{
	public int ID;
	public int orgID;
	public string title;
    public string text;
    public string comment;
	public ArrayList actions;
	public ArrayList addCheck1;
	public ArrayList addCheck2;
	public int addRadio1;
	public int addRadio2;
    public string addText1;
    public string addText2;
    public DateTime createdate;
    public DateTime updatedate;
    public ArrayList reminder;
    public ArrayList derivedFrom;
    public ArrayList derivedMeasures;
    public ArrayList forwardedThemes;
    public int ShowToColleagues;
    public int ShowToSubUnits;

    /// <summary>
    /// Leere Maßnahme mit Default-Werten erzeugen
    /// </summary>
    /// <param name="project">ID des Projektes</param>
	public TTheme(string project)
	{
		ID = 0;
		orgID = 0;
		title = "";
        text = "";
        comment = "";
		actions = new ArrayList();
		addCheck1 = new ArrayList();
		addCheck2 = new ArrayList();
		addRadio1 = 0;
		addRadio2 = 0;
        addText1 = "";
        addText2 = "";
		createdate = DateTime.Now.Date;
		updatedate = DateTime.Now.Date;
        //reminder = new ArrayList();
        //derivedFrom = new ArrayList();
        //derivedMeasures = new ArrayList();
        //forwardedThemes = new ArrayList();
        ShowToColleagues = 0;
        ShowToSubUnits = 0;
    }
	/// <summary>
	/// Maßnahme aus DB lesen
	/// </summary>
	/// <param name="aid">ID der Maßnahme</param>
	/// <param name="aProjectID">ID des Projektes</param>
    public TTheme(string aSelectedStructure, int aid, string aProjectID)
	{
		ID = aid;
		SqlDB dataReader;
        dataReader = new SqlDB("SELECT orgID, title, text, comment, showtocolleagues, showtosubunits, createdate, updatedate, addValue3, addValue4, addText1, addText2 FROM followup_themes" + aSelectedStructure + " WHERE ID='" + aid + "'", aProjectID);
		if (dataReader.read()) 
		{
			orgID = dataReader.getInt32(0);
			title = dataReader.getString(1);
            text = dataReader.getString(2);
            comment = dataReader.getString(3);
            ShowToColleagues = dataReader.getInt32(4);
            ShowToSubUnits = dataReader.getInt32(5);
            createdate = dataReader.getDateTime(6);
            updatedate = dataReader.getDateTime(7);
			addRadio1 = dataReader.getInt32(8);
            addRadio2 = dataReader.getInt32(9);
            addText1 = dataReader.getString(10);
            addText2 = dataReader.getString(11);
		}
		dataReader.close();

		actions = new ArrayList();
		dataReader = new SqlDB("SELECT action FROM followup_mapthemeactions" + aSelectedStructure + " WHERE theme='" + ID + "' ORDER BY action", aProjectID);
		while (dataReader.read()) 
		{
			actions.Add(dataReader.getInt32(0));
		}
		dataReader.close();

		addCheck1 = new ArrayList();
        dataReader = new SqlDB("SELECT addvalue1 FROM followup_maptheme" + aSelectedStructure + "addvalue1 WHERE theme='" + ID + "' ORDER BY addvalue1", aProjectID);
		while (dataReader.read()) 
		{
			addCheck1.Add(dataReader.getInt32(0));
		}
		dataReader.close();

		addCheck2 = new ArrayList();
        dataReader = new SqlDB("SELECT addvalue2 FROM followup_maptheme" + aSelectedStructure + " addvalue2 WHERE theme='" + ID + "' ORDER BY addvalue2", aProjectID);
		while (dataReader.read()) 
		{
			addCheck2.Add(dataReader.getInt32(0));
		}
		dataReader.close();

        //derivedFrom = new ArrayList();
        //dataReader = new SqlDB("SELECT themebase FROM followup_mapthemethemebase" + aSelectedStructure + " WHERE theme='" + ID + "' ORDER BY themebase", aProjectID);
        //while (dataReader.read())
        //{
        //    derivedFrom.Add(dataReader.getInt32(0));
        //}
        //dataReader.close();

        //reminder = new ArrayList();
        //dataReader = new SqlDB("SELECT ID FROM followup_themereminder" + aSelectedStructure + " WHERE themeID='" + ID + "'", aProjectID);
        //while (dataReader.read())
        //{
        //    reminder.Add(new TThemeReminder(aSelectedStructure, dataReader.getInt32(0),aProjectID));
        //}
        //dataReader.close();

        //derivedMeasures = new ArrayList();
        //dataReader = new SqlDB("SELECT measure FROM followup_mapmeasurethemebase" + aSelectedStructure + " WHERE themebase='" + ID + "'", aProjectID);
        //while (dataReader.read())
        //{
        //    derivedMeasures.Add(dataReader.getInt32(0));
        //}
        //dataReader.close();

        //forwardedThemes = new ArrayList();
        //dataReader = new SqlDB("SELECT theme FROM followup_mapthemethemebase" + aSelectedStructure + " WHERE themebase='" + ID + "'", aProjectID);
        //while (dataReader.read())
        //{
        //    forwardedThemes.Add(dataReader.getInt32(0));
        //}
        //dataReader.close();

	}
    public int getTopOrg(string aSelectedStructure, string aProjectID)
    {
        int result = 0;
        SqlDB dataReader = new SqlDB("SELECT topOrgID FROM structure" + aSelectedStructure + " WHERE orgID='" + orgID + "'", aProjectID);
        while (dataReader.read())
        {
            result = dataReader.getInt32(0);
        }
        dataReader.close();
        return result;
    }
}
