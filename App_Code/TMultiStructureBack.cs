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
/// Zusammenfassungsbeschreibung für TStructure1
///     Klasse zur Bearbeitung aller Eigenschaften einer Organisationseinheit,
///     gleichzeitig für mehrere Organisationseinheiten,
///     unabhängig von der Zugehörigkeit zu einem Tool
///     string mit "*" entspricht verschiedenen Inhalten innerhalb der Organisationseinheiten
///     int mit "-1" oder "-2" entspricht verschiedenen Inhalten innerhalb der Organisationseinheiten
/// </summary>
public class TMultiStructureBack
{
    public class TCheck
    {
        public int value;
        public int active;
    }
    public string OrgID;			// eindeutige ID der Einheit
    public string TopOrgID;			// eindeutige ID der Einheit
    public string orgDisplayName;			// Text der Einheit
    public string orgDisplayNameShort;			// Text der Einheit
    public string Filter;

    /// <summary>
    /// Erzuegen eines Objektes für die Eigenschaften mehrerer Organisationseinheiten
    /// </summary>
    public TMultiStructureBack()
    {
    }
    /// <summary>
    /// Ermittlung ob eine Eigenschaft einheitlich für alle Organisationseinheiten auf Wert gesetzt werden soll;
    /// Es wird ein passender Parameter für das DB-Update erzeugt und das Update-Command entsprechen ergänzt
    /// </summary>
    /// <param name="aValue">In die DB einzutragender Wert; bei * als Wert wird keine Änderung vorgenommen</param>
    /// <param name="aField">Datenbankfeld für den Wert</param>
    /// <param name="aUpdate">Update-Command für das DB-Update; Rückgabewert</param>
    /// <param name="parameterList">Paramterliste für das DB-Update</param>
    /// <returns></returns>
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
    /// <summary>
    /// Ermittlung ob eine Eigenschaft einheitlich für alle Organisationseinheiten auf Wert gesetzt werden soll;
    /// Es wird ein passender Parameter für das DB-Update erzeugt und das Update-Command entsprechen ergänzt
    /// </summary>
    /// <param name="aValue">In die DB einzutragender Wert; bei aInvalid als Wert wird keine Änderung vorgenommen</param>
    /// <param name="aField">Datenbankfeld für den Wert</param>
    /// <param name="aUpdate">Update-Command für das DB-Update; Rückgabewert</param>
    /// <param name="aInvalid">Wert der die "Nicht-Änderung" kennzeichnet</param>
    /// <param name="parameterList">Paramterliste für das DB-Update</param>
    /// <returns></returns>
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
    /// <summary>
    /// Update der Datenbank mit allen zu änderden Eigenschaften mehrerer Organisationseinheiten
    /// </summary>
    /// <param name="aProjectID">Projekt ID des Projektes</param>
    /// <param name="aOrgIDList">Liste der OrgID aller zu ändernder Organisationseinheiten</param>
    public void update(string aProjectID, ArrayList aOrgIDList)
    {
        SqlDB dataReader;

        // Update-String konstruieren
        string tempUpdate = "";
        TParameterList parameterList = new TParameterList();
        tempUpdate = checkUpdate(orgDisplayName, "displayName", tempUpdate, ref parameterList);
        tempUpdate = checkUpdate(orgDisplayNameShort, "displayNameShort", tempUpdate, ref parameterList);
        tempUpdate = checkUpdate(Filter, "filter", tempUpdate, ref parameterList);
        parameterList.addParameter("orgID", "string", "");

        // Schleife über alle UserIDs
        foreach (string tempOrgID in aOrgIDList)
        {
            parameterList.changeParameterValue("orgID", tempOrgID);
            // Datenschreiben
            if (tempUpdate != "")
            {
                dataReader = new SqlDB(aProjectID);
                dataReader.execSQLwithParameter("UPDATE structureBack SET " + tempUpdate + " WHERE orgID=@orgID", parameterList);
            }
        }
    }
    public static string getOrgPath(int aOrgID, string aProjectID)
    {
        string result = "";

        SqlDB dataReader;
        while (aOrgID != 0)
        {
            dataReader = new SqlDB("select topOrgID, displayName from structureBack where orgID='" + aOrgID + "'", aProjectID);
            if (dataReader.read())
            {
                aOrgID = dataReader.getInt32(0);
                if (result == "")
                {
                    result = dataReader.getString(1);
                }
                else
                {
                    result = dataReader.getString(1) + "->" + result;
                }
            }
            dataReader.close();
        }
        return result;
    }
}