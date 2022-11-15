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
using System.IO;

/// <summary>
/// Klasse zur Bearbeitung der Eigenschaften aller Tools
/// </summary>
public class TTools
{
    public class TToolEntry
    {
        public int orderID;
        public string toolID;
        public string toolName;
        public string toolLink;
        public string description;
        public bool active;
        public DateTime startdate;
        public DateTime enddate;
    }

    public ArrayList tools;

    /// <summary>
    /// Einlesen aller Eigenschaften aller Tools
    /// </summary>
    /// <param name="aLanguage">Sprache der Toolbeschreibungen</param>
    /// <param name="aProjectID">Eindeutige ID des Projektes</param>
    public TTools(string aLanguage, string aProjectID)
    {
        tools = new ArrayList();

        SqlDB dataReader;
        dataReader = new SqlDB("SELECT orderID, toolID, toolName, toolLink, active, startdate, enddate FROM tools ORDER BY orderID", aProjectID);
        while (dataReader.read())
        {
            TToolEntry aTool = new TToolEntry();
            aTool.orderID = dataReader.getInt32(0);
            aTool.toolID = dataReader.getString(1);
            aTool.toolName = dataReader.getString(2);
            aTool.description = Labeling.getLabel("TxtTools" + dataReader.getString(1) + "Description", aLanguage, aProjectID);
            aTool.toolLink = dataReader.getString(3);
            aTool.active = dataReader.getBool(4);
            aTool.startdate = dataReader.getDateTime(5);
            aTool.enddate = dataReader.getDateTime(6);
            tools.Add(aTool);
        }
        dataReader.close();
    }
    /// <summary>
    /// Speichern der Einstellungen
    /// </summary>
    /// <param name="aReportPath">Pfad zum Download-Verzeichnis</param>
    /// <param name="aProjectID">ID des Projektes</param>
    public void save(string aProjectID)
    {
        deleteALL(aProjectID);
        // Werte in DB schreiben
        foreach (TTools.TToolEntry tempEntry in tools)
        {
            string tempBool = "0";
            if (tempEntry.active)
                tempBool = "1";
            TParameterList parameterList = new TParameterList();
            parameterList.addParameter("orderID", "int", tempEntry.orderID.ToString());
            parameterList.addParameter("toolID", "string", tempEntry.toolID);
            parameterList.addParameter("toolName", "string", tempEntry.toolName);
            parameterList.addParameter("toolLink", "string", tempEntry.toolLink);
            parameterList.addParameter("active", "int", tempBool);
            parameterList.addParameter("startdate", "datetime", tempEntry.startdate.ToString());
            parameterList.addParameter("enddate", "datetime", tempEntry.enddate.ToString());
            SqlDB dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO tools (orderID, toolID, toolName, toolLink, active, startdate, enddate) VALUES (@orderID, @toolID, @toolName, @toolLink, @active, @startdate, @enddate)", parameterList);
        }
    }
    /// <summary>
    /// Neues Tool in DB einfügen
    /// </summary>
    /// <param name="aOrderID">Ordnungsnummer des Tools</param>
    /// <param name="aToolID">ID des Tools</param>
    /// <param name="aToolName">Anzeigename des Tools</param>
    /// <param name="aToolLink">Link zum Tool</param>
    /// <param name="aToolPath">Pfad zum Tool</param>
    /// <param name="aActive">Aktivstatus des Tools</param>
    /// <param name="aReportPath">Pfad zu Berichten</param>
    /// <param name="aProjectID">ID des Projektes</param>
    public static void add(int aOrderID, string aToolID, string aToolName, string aToolLink, bool aActive, string aStartdate, string aEnddate, string aProjectID)
    {
        // ToolID auf Eindeutigkeit prüfen
        SqlDB dataReader = new SqlDB("SELECT toolID from tools WHERE toolID='" + aToolID + "'", aProjectID);
        if (!dataReader.read())
        {
            // Werte in DB schreiben
            string tempBool = "0";
            if (aActive)
                tempBool = "1";
            DateTime startdate = Convert.ToDateTime(aStartdate);
            DateTime enddate = Convert.ToDateTime(aEnddate);
            TParameterList parameterList = new TParameterList();
            parameterList.addParameter("orderID", "int", aOrderID.ToString());
            parameterList.addParameter("toolID", "string", aToolID);
            parameterList.addParameter("toolName", "string", aToolName);
            parameterList.addParameter("toolLink", "string", aToolLink);
            parameterList.addParameter("active", "int", tempBool);
            parameterList.addParameter("startdate", "datetime", startdate.ToString());
            parameterList.addParameter("enddate", "datetime", enddate.ToString());
            SqlDB dataReader1 = new SqlDB(aProjectID);
            dataReader1.execSQLwithParameter("INSERT INTO tools (orderID, toolID, toolName, toolLink, active, startdate, enddate) VALUES (@orderID, @toolID, @toolName, @toolLink, @active, @startdate, @enddate)", parameterList);
        }
        dataReader.close();
    }
    /// <summary>
    /// Löschen eines Tools aus DB
    /// </summary>
    /// <param name="aToolID">ID des Tools</param>
    /// <param name="aProjectID">ID des Projektes</param>
    public void delete(string aToolID, string aProjectID)
    {
        // aus DB löschen
        SqlDB dataReader = new SqlDB(aProjectID);
        dataReader.execSQL("DELETE FROM tools WHERE toolID='" + aToolID + "'");
        // Userrechte auf Tool löschen
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQL("DELETE FROM mapusertool WHERE toolID='" + aToolID + "'");
    }
    /// <summary>
    /// Alle Tools aus DB löschen
    /// </summary>
    /// <param name="aProjectID">ID des Projektes</param>
    public void deleteALL(string aProjectID)
    {
        // aus DB löschen
        SqlDB dataReader = new SqlDB(aProjectID);
        dataReader.execSQL("TRUNCATE TABLE tools");
    }
    /// <summary>
    /// Ermittlung der URL eiens Tools
    /// </summary>
    /// <param name="aToolID">ID des Tools</param>
    /// <returns>URL des Tools</returns>
    public string getLink(string aToolID)
    {
        string link = "";
        foreach (TToolEntry tempEntry in tools)
        {
            if (tempEntry.toolID == aToolID)
                link = tempEntry.toolLink;
        }
        return link;
    }
    /// <summary>
    /// Ermittlung des Aktiv-Status eines Tools
    /// </summary>
    /// <param name="aToolID">ID des Tools</param>
    /// <returns>true wenn Tool aktiv, sonst false</returns>
    public bool getActive(string aToolID)
    {
        bool active = false;
        foreach (TToolEntry tempEntry in tools)
        {
            DateTime actDate = DateTime.Now;
            if ( (tempEntry.toolID == aToolID) && (tempEntry.startdate < actDate) && (tempEntry.enddate > actDate))
                active = tempEntry.active;
        }
        return active;
    }
}
