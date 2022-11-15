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
/// Zusammenfassungsbeschreibung für Measure
/// </summary>
public class TMeasure
{
	public decimal ID;
	public int orgID;
	public string title;
    public string text;
    public string solution;
	public int state;
	public DateTime startdate;
    public DateTime enddate;
	public string name;
	public int source;
	public string mail;
	public ArrayList actions;
	public ArrayList people;
	public ArrayList addCheck1;
	public ArrayList addCheck2;
	public int addRadio1;
	public int addRadio2;
    public string addText1;
    public string addText2;
    public int time;
	public int money;
    public int publish;
    public DateTime createdate;
    public DateTime updatedate;
    public DateTime progressdate;
    public bool ShowToSubunits;
    public bool ShowToColleagues;
    public bool IsOrder;

    /// <summary>
    /// Leere Maßnahme mit Default-Werten erzeugen
    /// </summary>
    /// <param name="project">ID des Projektes</param>
	public TMeasure(string project)
	{
		ID = 0;
		orgID = 0;
		title = "";
        text = "";
        solution = "";
		state = 0;
		startdate = DateTime.Now.Date;
		enddate = DateTime.Now.Date;
		name = "";
		source = 0;
		mail = "";
		actions = new ArrayList();
		people = new ArrayList();
		addCheck1 = new ArrayList();
		addCheck2 = new ArrayList();
		addRadio1 = 0;
		addRadio2 = 0;
        addText1 = "";
        addText2 = "";
		createdate = DateTime.Now.Date;
		updatedate = DateTime.Now.Date;
		progressdate = DateTime.Now.Date;
	}
	/// <summary>
	/// Maßnahme aus DB lesen
	/// </summary>
	/// <param name="aid">ID der Maßnahme</param>
	/// <param name="aProjectID">ID des Projektes</param>
    public TMeasure(string aSelectedStructure, int aid, string aProjectID)
	{
		ID = aid;
		SqlDB dataReader;
        dataReader = new SqlDB("SELECT orgID, title, text, solution, state, startdate, enddate, name, source, mail, createdate, updatedate, progressdate, addValue3, addValue4, addText1, addText2, ShowToSubunits, ShowToColleagues, isOrder FROM followup_measures" + aSelectedStructure + " WHERE ID='" + aid + "'", aProjectID);
		if (dataReader.read()) 
		{
			orgID = dataReader.getInt32(0);
			title = dataReader.getString(1);
            text = dataReader.getString(2);
            solution = dataReader.getString(3);
			state = dataReader.getInt32(4);
			startdate = dataReader.getDateTime(5);
			enddate = dataReader.getDateTime(6);
			name = dataReader.getString(7);
			source = dataReader.getInt32(8);
			mail = dataReader.getString(9);
            createdate = dataReader.getDateTime(10);
            updatedate = dataReader.getDateTime(11);
            progressdate = dataReader.getDateTime(12);
			addRadio1 = dataReader.getInt32(13);
            addRadio2 = dataReader.getInt32(14);
            addText1 = dataReader.getString(15);
            addText2 = dataReader.getString(16);
            ShowToSubunits = dataReader.getBool(17);
            ShowToColleagues = dataReader.getBool(18);
            IsOrder = dataReader.getBool(19);
		}
		dataReader.close();

		actions = new ArrayList();
		dataReader = new SqlDB("SELECT action FROM followup_measureaction" + aSelectedStructure + " WHERE measure='" + ID + "' ORDER BY action", aProjectID);
		while (dataReader.read()) 
		{
			actions.Add(dataReader.getInt32(0));
		}
		dataReader.close();

		people = new ArrayList();
        dataReader = new SqlDB("SELECT people FROM followup_measureparticipants" + aSelectedStructure + " WHERE measure='" + ID + "' ORDER BY people", aProjectID);
		while (dataReader.read()) 
		{
			people.Add(dataReader.getInt32(0));
		}
		dataReader.close();

		addCheck1 = new ArrayList();
        dataReader = new SqlDB("SELECT addvalue1 FROM followup_measure" + aSelectedStructure + "addvalue1 WHERE measure='" + ID + "' ORDER BY addvalue1", aProjectID);
		while (dataReader.read()) 
		{
			addCheck1.Add(dataReader.getInt32(0));
		}
		dataReader.close();

		addCheck2 = new ArrayList();
        dataReader = new SqlDB("SELECT addvalue2 FROM followup_measure" + aSelectedStructure + "addvalue2 WHERE measure='" + ID + "' ORDER BY addvalue2", aProjectID);
		while (dataReader.read()) 
		{
			addCheck2.Add(dataReader.getInt32(0));
		}
		dataReader.close();

	}
}
