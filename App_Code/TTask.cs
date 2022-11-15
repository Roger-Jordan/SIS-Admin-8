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
public class TTask
{
    public bool exists;
    public int ID;
    public int topID;
    public int position;
	public string title1;
    public string text1;
    public string title2;
    public string text2;
	public DateTime startdate;
    public DateTime enddate;
    public bool milestone;
    public int role;
    public int state;
    public bool isSummary;
    public ArrayList childs;

    /// <summary>
    /// Leere Ausgabe mit Default-Werten erzeugen
    /// </summary>
    /// <param name="project">ID des Projektes</param>
	public TTask(string project)
	{
		ID = 0;
		topID = 0;
        position = 0;
		title1 = "";
        text1 = "";
        title2 = "";
        text2 = "";
		startdate = DateTime.Now.Date;
		enddate = DateTime.Now.Date;
        milestone = false;
        role = 0;
        state = 0;
        isSummary = false;
        childs = new ArrayList();
    }
	/// <summary>
	/// Aufgabe aus DB lesen
	/// </summary>
	/// <param name="aTaskID">ID der Aufgabe</param>
	/// <param name="aProjectID">ID des Projektes</param>
    public TTask(int aTaskID, string aProjectID)
	{
		ID = aTaskID;
        exists = false;
		SqlDB dataReader;
        dataReader = new SqlDB("SELECT topID, position, title1, text1, title2, text2, startdate, enddate, milestone, role, status FROM teamspace_tasks WHERE taskID='" + aTaskID + "'", aProjectID);
		if (dataReader.read()) 
		{
            exists = true;
			topID = dataReader.getInt32(0);
            position = dataReader.getInt32(1);
			title1 = dataReader.getString(2);
            text1 = dataReader.getString(3);
            title2 = dataReader.getString(4);
            text2 = dataReader.getString(5);
			startdate = dataReader.getDateTime(6).Date;
            enddate = dataReader.getDateTime(7).Date.AddHours(23);
            milestone = dataReader.getBool(8);
            role = dataReader.getInt32(9);
            state = dataReader.getInt32(10);
            if (milestone)
            {
                // Meilenstein liegt immer auf Tagesmitte und Startdatum muss Enddatum sein
                enddate = enddate.AddHours(-11);
                startdate = enddate;
            }
            SqlDB dataReader1;
            dataReader1 = new SqlDB("SELECT taskID FROM teamspace_tasks WHERE topID='" + aTaskID + "'", aProjectID);
            if (dataReader1.read())
                isSummary = true;
            else
                isSummary = false;
            dataReader1.close();
            childs = new ArrayList();
        }
		dataReader.close();
	}
    /// <summary>
    /// Ausgabe in DB schreiben
    /// </summary>
    /// <param name="aProjectID">ID des Projektes</param>
	public void Update(string aProjectID)
	{
        int tempMilestone = 0;
        if (milestone)
            tempMilestone = 1;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("topID", "int", topID.ToString());
        parameterList.addParameter("position", "int", position.ToString());
        parameterList.addParameter("title1", "string", title1);
        parameterList.addParameter("text1", "string", text1);
        parameterList.addParameter("title2", "string", title2);
        parameterList.addParameter("text2", "string", text2);
        parameterList.addParameter("startdate", "datetime", startdate.Date.ToString());
        parameterList.addParameter("enddate", "datetime", enddate.Date.ToString());
        parameterList.addParameter("milestone", "int", tempMilestone.ToString());
        parameterList.addParameter("role", "int", role.ToString());
        parameterList.addParameter("status", "int", state.ToString());
        parameterList.addParameter("taskID", "int", ID.ToString());
        
		SqlDB dataReader;
    	// update
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("UPDATE teamspace_tasks SET  topID=@topID, position=@position, title1=@title1, text1=@text1, title2=@title2, text2=@text2, startdate=@startdate, enddate=@enddate, milestone=@milestone, role=@role, status=@status  WHERE taskID=@taskID", parameterList);
	}
    /// <summary>
    /// Löschen einer Maßnahme aus DB
    /// </summary>
    /// <param name="aProjectID"></param>
	public static void delTask(int aTaskID, string aProjectID)
	{
        SqlDB dataReader;

        // untergeordnete Tasks löschen
        dataReader = new SqlDB("SELECT taskID FROM teamspace_tasks WHERE topID='" + aTaskID + "'", aProjectID);
        while (dataReader.read())
        {
            TTask.delTask(dataReader.getInt32(0), aProjectID);
        }
        dataReader.close();

        // verknüpfte Dateien löschen, sofern diese nicht noch mit anderen Aufgaben verknüpft sind
        //###

        // Task aus Datenbank löschen
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("ID", "int", aTaskID.ToString());
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("DELETE FROM teamspace_tasks WHERE taskID=@ID", parameterList);
    }
    public static int getTopID(int aTaskID, string aProjectID)
    {
        int result = 0;
        SqlDB dataReader;
        dataReader = new SqlDB("SELECT topID FROM teamspace_tasks WHERE taskID='" + aTaskID + "'", aProjectID);
        if (dataReader.read())
        {
            result = dataReader.getInt32(0);
        }
        dataReader.close();
        return result;
    }
    public static int getNewPosition(int aTaskID, string aProjectID)
    {
        SqlDB dataReader;

        int actPosition = 0;
        int actTopID = 0;

        // Position der aktuellen Aufgabe ermitteln
        dataReader = new SqlDB("SELECT topID, position FROM teamspace_tasks WHERE taskID='" + aTaskID + "'", aProjectID);
        if (dataReader.read())
        {
            actTopID = dataReader.getInt32(0);
            actPosition = dataReader.getInt32(1);
        }
        dataReader.close();

        // alle nachfogenden Aufgaben um eine Position nach unten verschieben
        dataReader = new SqlDB("SELECT taskID FROM teamspace_tasks WHERE topID='" + actTopID.ToString() + "' AND position > '" + actPosition.ToString() + "' ORDER by position", aProjectID);
        while (dataReader.read())
        {
            int tempTaskID = dataReader.getInt32(0);

            // diesen Task eine Position nach unten
            TTask downTask = new TTask(tempTaskID, aProjectID);
            downTask.position++;
            downTask.Update(aProjectID);
        }
        dataReader.close();

        actPosition++;
        return actPosition;
    }
    public static int getEndPosition(int aTaskID, string aProjectID)
    {
        SqlDB dataReader;

        int actPosition = 0;

        // Position der aktuellen Aufgabe ermitteln
        dataReader = new SqlDB("SELECT max(position) FROM teamspace_tasks WHERE topID='" + aTaskID + "'", aProjectID);
        if (dataReader.read())
        {
            actPosition = dataReader.getInt32(0);
        }
        dataReader.close();

        actPosition++;
        return actPosition;
    }
    public static void upTask(int aTaskID, string aProjectID)
    {
        TTask actTask = new TTask(aTaskID, aProjectID);

        int upTaskID = -1;
        // Schleife über alle Tasks der geleicher Ebene
        SqlDB dataReader;
        dataReader = new SqlDB("SELECT taskID FROM teamspace_tasks WHERE topID='" + actTask.topID + "' ORDER by position", aProjectID);
        while (dataReader.read())
        {
            int tempTaskID = dataReader.getInt32(0);

            if (tempTaskID == aTaskID)
            {
                // diesen Task eine Position nach oben, falls möglich
                if (upTaskID != -1)
                {
                    TTask upTask = new TTask(upTaskID, aProjectID);
                    int upTaskPostition = upTask.position;
                    upTask.position = actTask.position;
                    upTask.Update(aProjectID);
                    actTask.position = upTaskPostition;
                    actTask.Update(aProjectID);
                }
            }
            upTaskID = tempTaskID;
        }
        dataReader.close();
    }
    public static void downTask(int aTaskID, string aProjectID)
    {
        TTask actTask = new TTask(aTaskID, aProjectID);

        bool foundActTask = false;
        // Schleife über alle Tasks der geleicher Ebene
        SqlDB dataReader;
        dataReader = new SqlDB("SELECT taskID FROM teamspace_tasks WHERE topID='" + actTask.topID + "' ORDER by position", aProjectID);
        while (dataReader.read())
        {
            int tempTaskID = dataReader.getInt32(0);

            if (foundActTask)
            {
                // diesen Task eine Position nach oben
                TTask upTask = new TTask(tempTaskID, aProjectID);
                int upTaskPostition = upTask.position;
                upTask.position = actTask.position;
                upTask.Update(aProjectID);
                actTask.position = upTaskPostition;
                actTask.Update(aProjectID);

                foundActTask = false;
            }
            if (tempTaskID == aTaskID)
                foundActTask = true;
        }
        dataReader.close();

    }
}
