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
///     Klasse zur Bearbeitung aller Eigenschaften einer Organisationseinheit,
///     gleichzeitig für mehrere Organisationseinheiten,
///     unabhängig von der Zugehörigkeit zu einem Tool
///     string mit "*" entspricht verschiedenen Inhalten innerhalb der Organisationseinheiten
///     int mit "-1" oder "-2" entspricht verschiedenen Inhalten innerhalb der Organisationseinheiten
/// </summary>
public class TMultiStructure
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
    // OrgManager
    public string ONumberSoll;
    public string ONumberTotalSoll;
    public int OStatus;
    public string OBackComparison;
    public int OLocked;
    // FollowUp
    public string FMailAdress;	    // Email-Adresse der Einheit
    public int FWorkshop;			// O=soll keine Maßnahmen durchführen, 1=soll Maßnahmen durchführen
    public int FLifePDF;			// PDF kann direkt angezeigt werden
    public int FMaxpdfLevel;			// maximale Tiefe der Maßnahmen-PDF-Liste
    public int FCommunicationfinished;			// Kennzeichen ob Kommunikation der Ergebnisse durchgeführt wurde
    public string FCommunicationfinsheddate;			//  Datum der Kommunikation der Ergebnisse
    public string FParticipants;        // Anzahl Teilnehmer bei Ergebniskommunikation
    // Download
    // Response
    public string RNumberSoll;
    public string RNumberIst;

    /// <summary>
    /// Erzuegen eines Objektes für die Eigenschaften mehrerer Organisationseinheiten
    /// </summary>
    public TMultiStructure()
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
    protected string checkUpdate(string aValue, string aField, string aUpdate, ref TParameterList parameterList, bool aDateTime)
    {
        parameterList.addParameter(aField, "datetime", aValue);

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
        // OrgManager
        string tempUpdateOrgManager = "";
        TParameterList parameterListOrgManager = new TParameterList();
        tempUpdateOrgManager = checkUpdate(ONumberSoll, "number", tempUpdateOrgManager, ref parameterListOrgManager);
        tempUpdateOrgManager = checkUpdate(ONumberTotalSoll, "numberTotal", tempUpdateOrgManager, ref parameterListOrgManager);
        tempUpdateOrgManager = checkUpdate(OStatus, "status", tempUpdateOrgManager, "-1", ref parameterListOrgManager);
        tempUpdateOrgManager = checkUpdate(OBackComparison, "backComparison", tempUpdateOrgManager, ref parameterListOrgManager);
        tempUpdateOrgManager = checkUpdate(OLocked, "locked", tempUpdateOrgManager, "-1", ref parameterListOrgManager);
        // FollowUp
        string tempUpdateFollowUp = "";
        TParameterList parameterListFollowUp = new TParameterList();
        tempUpdateFollowUp = checkUpdate(FMailAdress, "mail", tempUpdateFollowUp, ref parameterListFollowUp);
        tempUpdateFollowUp = checkUpdate(FWorkshop, "workshop", tempUpdateFollowUp, "-1", ref parameterListFollowUp);
        tempUpdateFollowUp = checkUpdate(FLifePDF, "livepdf", tempUpdateFollowUp, "-1", ref parameterListFollowUp);
        tempUpdateFollowUp = checkUpdate(FMaxpdfLevel, "maxpdflevel", tempUpdateFollowUp, "-1", ref parameterListFollowUp);
        tempUpdateFollowUp = checkUpdate(FCommunicationfinished, "communicationfinished", tempUpdateFollowUp, "-1", ref parameterListFollowUp);
        tempUpdateFollowUp = checkUpdate(FCommunicationfinsheddate, "communicationfinsheddate", tempUpdateFollowUp, ref parameterListFollowUp, true);
        tempUpdateFollowUp = checkUpdate(FParticipants, "participants", tempUpdateFollowUp, ref parameterListFollowUp);
        // Response
        string tempUpdateResponse = "";
        TParameterList parameterListResponse = new TParameterList();
        tempUpdateResponse = checkUpdate(RNumberSoll, "target", tempUpdateResponse, ref parameterListResponse);
        tempUpdateResponse = checkUpdate(RNumberIst, "number", tempUpdateResponse, ref parameterListResponse);
        // Download
        // bisher keine Ergänzungen

        parameterList.addParameter("orgID", "string", "");
        parameterListOrgManager.addParameter("orgID", "string", "");
        parameterListFollowUp.addParameter("orgID", "string", "");
        parameterListResponse.addParameter("orgID", "string", "");
        //parameterListDownload.addParameter("orgID", "string", "");

        // Schleife über alle UserIDs
        foreach (string tempOrgID in aOrgIDList)
        {
            parameterList.changeParameterValue("orgID", tempOrgID);
            parameterListOrgManager.changeParameterValue("orgID", tempOrgID);
            parameterListFollowUp.changeParameterValue("orgID", tempOrgID);
            parameterListResponse.changeParameterValue("orgID", tempOrgID);
            //parameterListDownload.changeParameterValue("orgID", tempOrgID);
            // Datenschreiben
            if (tempUpdate != "")
            {
                dataReader = new SqlDB(aProjectID);
                dataReader.execSQLwithParameter("UPDATE structure SET " + tempUpdate + " WHERE orgID=@orgID", parameterList);
            }
            // OrgManager
            if (tempUpdateOrgManager != "")
            {
                dataReader = new SqlDB(aProjectID);
                dataReader.execSQLwithParameter("UPDATE orgmanager_structure SET " + tempUpdateOrgManager + " WHERE orgID=@orgID", parameterListOrgManager);
            }
            // FollowUp
            if (tempUpdateFollowUp != "")
            {
                dataReader = new SqlDB(aProjectID);
                dataReader.execSQLwithParameter("UPDATE followup_structure SET " + tempUpdateFollowUp + " WHERE orgID=@orgID", parameterListFollowUp);
            }
            // Response
            if (tempUpdateResponse != "")
            {
                dataReader = new SqlDB(aProjectID);
                dataReader.execSQLwithParameter("UPDATE response_structure SET " + tempUpdateResponse + " WHERE orgID=@orgID", parameterListResponse);
            }
            // Download
            //if (tempUpdateDownload != "")
            //{
            //    dataReader = new SqlDB(aProjectID);
            //    dataReader.execSQLwithParameter("UPDATE download_structure SET " + tempUpdateDownload + " WHERE orgID=@orgID", parameterListDownload);
            //}
        }
    }
    public static string getOrgPath(int aOrgID, string aProjectID, bool recursive, bool displaynameShort)
    {
        string result = "";

        SqlDB dataReader;

        while (aOrgID != 0)
        {
            if (displaynameShort)
                dataReader = new SqlDB("select topOrgID, displayNameShort from structure where orgID='" + aOrgID + "'", aProjectID);
            else
                dataReader = new SqlDB("select topOrgID, displayName from structure where orgID='" + aOrgID + "'", aProjectID);
            if (dataReader.read())
            {
                aOrgID = dataReader.getInt32(0);
                if (result == "")
                {
                    result = dataReader.getString(1);
                    if (!recursive)
                    {
                        // Pfad abbrechen
                        aOrgID = 0;
                    }
                }
                else
                {
                    result = dataReader.getString(1) + "->" + result;
                }
            }
            else
            {
                // Fehler ausgetreten, Abbruch des Pfades
                aOrgID = 0;
            }
            dataReader.close();
        }
        return result;
    }
}