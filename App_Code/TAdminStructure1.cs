using System;
using System.Collections;

/// <summary>
/// Zusammenfassungsbeschreibung für TStructure
///     Klasse zur Bearbeitung aller Eigenschaften einer Organisationseinheit,
///     unabhängig von der Zugehörigkeit zu einem Tool
/// </summary>
public class TAdminStructure1
{
    public class TKeyValue
    {
        public string key;
        public string value;
    }

    public bool expandedNav;		// Anzeige der Untereinheiten in der Navigation
    public int OrgID;			// eindeutige ID der Einheit
    public int TopOrgID;			// eindeutige ID der übergeordneten Einheit
    public string orgDisplayName;			// Text der Einheit
    public string orgDisplayNameShort;			// Text der Einheit
    public TAdminStructure1 parent;		// Verweis auf übergeordneten Einheit
    public ArrayList childs;		// Liste von Verweisen auf untergeordnete Einheiten
    public int level;				// Level der aktuellen Einheit ist 0
    public bool hasChilds;
    public string projectID;
    public int maxLevel;
    // OrgManager
    public int numberEmployees;          // Anzahl Mitarbeiter laut Angabe
    public int numberEmployeesTotal;     // Anzahl Mitarbeiter laut Angabe aggregiert
    public int numberEmployeesSum;			// Anzahl der Mitarbeiter dieser Einheit inkl. Untereinheiten (berechnet laut Angabe)
    public int numberEmployeesInput;     // Anzahl der eingetragenen Mitarbeiter dieser Einheit
    public int numberEmployeesInputSum;	    // Anzahl der eingetragen Mitarbeiter dieser Einheit inkl. Untereinheiten (berechnet)
    public int status;
    public int backComparison;
    public bool locked;
    public int countFK;             // Anzahld er Fk der Einheit
    public ArrayList values;        // alle Werte der variablen Felder
    // FollowUp
    public string mailAdress;	    // Email-Adresse der Einheit
    public bool workshop;			// O=soll keine Maßnahmen durchführen, 1=soll Maßnahmen durchführen
    public bool LifePDF;			// PDF kann direkt angezeigt werden
    public int maxpdfLevel;			// maximale Tiefe der Maßnahmen-PDF-Liste
    public bool communicationfinished;			// Kennzeichen ob Ergebnisse kommuniziert sind
    public DateTime communicationfinsheddate;			// Datum der Kommunikation der Ergebnisse
    public int numberOfParticipants;			// Anzahl Workshopteilnehmer
    // Download
    // Response
    public int numberIst;
    public int target;

    /// <summary>
    ///  Einlesen einer vollständigen Organisationseinheit mit Eigenschaften aller Tools
    /// </summary>
    /// <param name="aOrgID">Eindeutige ID der zu erzeugenden Organisationseinheit; -1 entspricht der obersten Organisationseinheit des Projektes</param>
    /// <param name="aChild">Objekt einer Organisationseinheit die als Child der zu erzeugenden eingetagen wird</param>
    /// <param name="aProject">Eindeutige ID des Projektes</param>
    /// <param name="aLevel">Level auf dem sich die zu erzeugende Organisationseinheit befindet</param>
    public TAdminStructure1(int aOrgID, TAdminStructure1 aParent, int aLevel, string aProject)
    {
        if (aOrgID == -1)
        {
            // Root-Element nehmen;
            SqlDB dataReader1 = new SqlDB("SELECT OrgID from structure1 WHERE topOrgID='0'", aProject);
            if (dataReader1.read())
                aOrgID = dataReader1.getInt32(0);
            dataReader1.close();
        }
        // einzelnes Strukturelement einlesen
        childs = new ArrayList();
        parent = aParent;
        projectID = aProject;
        OrgID = aOrgID;
        orgDisplayName = "";
        expandedNav = false;
        level = aLevel;
        maxLevel = -1;
        mailAdress = "";
        values = new ArrayList();

        // Namen der Unit aus der Datenbank lesen
        SqlDB dataReader;
        dataReader = new SqlDB("select orgID, topOrgID, displayName, displayNameShort from structure1 where orgID='" + aOrgID + "'", aProject);
        if (dataReader.read())
        {
            TopOrgID = dataReader.getInt32(1);
            orgDisplayName = dataReader.getString(2);
            orgDisplayNameShort = dataReader.getString(3);
        }
        dataReader.close();
        // OrgManager
        dataReader = new SqlDB("select number, numberTotal, numberSum, numberSumInput, status, backComparison, locked from orgmanager_structure1 where orgID='" + aOrgID + "'", aProject);
        if (dataReader.read())
        {
            numberEmployees = dataReader.getInt32(0);
            numberEmployeesTotal = dataReader.getInt32(1);
            numberEmployeesSum = dataReader.getInt32(2);
            numberEmployeesInputSum = dataReader.getInt32(3);
            status = dataReader.getInt32(4);
            backComparison = dataReader.getInt32(5);
            locked = dataReader.getBool(6);
        }
        dataReader.close();
        values.Clear();
        dataReader = new SqlDB("SELECT fieldID, value FROM orgmanager_structure1fieldsvalues WHERE orgID='" + aOrgID + "'", aProject);
        while (dataReader.read())
        {
            TKeyValue tempKeyValue = new TKeyValue();
            tempKeyValue.key = dataReader.getString(0);
            tempKeyValue.value = dataReader.getString(1);
            values.Add(tempKeyValue);
        }
        dataReader.close();
        
        // Anzahl eingetragen Mitarbeiter der Einheit ermitteln
        dataReader = new SqlDB("SELECT COUNT(employeeID) FROM orgmanager_employee WHERE orgIDAdd1='" + aOrgID + "'", aProject);
        if (dataReader.read())
        {
            numberEmployeesInput = dataReader.getInt32(0);
        }
        dataReader.close();

        // FollowUp
        dataReader = new SqlDB("select workshop, mail, livepdf, maxpdfLevel, communicationfinished, communicationfinsheddate, participants from followup_structure1 where orgID='" + aOrgID + "'", aProject);
        if (dataReader.read())
        {
            workshop = dataReader.getBool(0);
            mailAdress = dataReader.getString(1);
            LifePDF = dataReader.getBool(2);
            maxpdfLevel = dataReader.getInt32(3);
            communicationfinished = dataReader.getBool(4);
            communicationfinsheddate = dataReader.getDateTime(5);
            numberOfParticipants = dataReader.getInt32(6);
        }
        dataReader.close();
        // Download
        // Response
        dataReader = new SqlDB("select number, target from response_structure1 where orgID='" + aOrgID + "'", aProject);
        if (dataReader.read())
        {
            numberIst = dataReader.getInt32(0);
            target = dataReader.getInt32(1);
        }
        dataReader.close();

        hasChilds = false;
        dataReader = new SqlDB("SELECT orgID from structure1 where topOrgID='" + aOrgID + "'", aProject);
        if (dataReader.read())
        {
            hasChilds = true;
        }
        dataReader.close();
    }
    public TAdminStructure1(int aOrgID, string aProject)
    {
        // einzelnes Strukturelement einlesen
        OrgID = aOrgID;

        // Namen der Unit aus der Datenbank lesen
        SqlDB dataReader;
        dataReader = new SqlDB("select orgID, topOrgID, displayName, displayNameShort from structure1 where orgID='" + aOrgID + "'", aProject);
        if (dataReader.read())
        {
            TopOrgID = dataReader.getInt32(1);
            orgDisplayName = dataReader.getString(2);
            orgDisplayNameShort = dataReader.getString(3);
        }
        dataReader.close();
    }
    /// <summary>
    ///  Aktzualisieren der Daten einer Einheit aus den Daten der Datenbank inkl. aller geladenen Untereinheiten
    /// </summary>
    /// <param name="aChild">Objekt der obersten Organisationseinheit die neu geladen werden soll</param>
    public void Reload()
    {
        // einzelnes Strukturelement neu einlesen

        // Namen der Unit aus der Datenbank lesen
        SqlDB dataReader = new SqlDB("select orgID, topOrgID, displayName, displayNameShort from structure1 where orgID='" + OrgID + "'", projectID);
        if (dataReader.read())
        {
            orgDisplayName = dataReader.getString(2);
            orgDisplayNameShort = dataReader.getString(3);
        }
        dataReader.close();
        // OrgManager
        dataReader = new SqlDB("select number, numberTotal, numberSum, numberSumInput, status, backComparison, locked from orgmanager_structure1 where orgID='" + OrgID + "'", projectID);
        if (dataReader.read())
        {
            numberEmployees = dataReader.getInt32(0);
            numberEmployeesTotal = dataReader.getInt32(1);
            numberEmployeesSum = dataReader.getInt32(2);
            numberEmployeesInputSum = dataReader.getInt32(3);
            status = dataReader.getInt32(4);
            backComparison = dataReader.getInt32(5);
            locked = dataReader.getBool(6);
        }
        dataReader.close();
        values.Clear();
        dataReader = new SqlDB("SELECT fieldID, value FROM orgmanager_structure1fieldsvalues WHERE orgID='" + OrgID + "'", projectID);
        while (dataReader.read())
        {
            TKeyValue tempKeyValue = new TKeyValue();
            tempKeyValue.key = dataReader.getString(0);
            tempKeyValue.value = dataReader.getString(1);
            values.Add(tempKeyValue);
        }
        dataReader.close();

        // Anzahl eingetragen Mitarbeiter der Einheit ermitteln
        dataReader = new SqlDB("SELECT COUNT(employeeID) FROM orgmanager_employee WHERE orgIDAdd1='" + OrgID + "'", projectID);
        if (dataReader.read())
        {
            numberEmployeesInput = dataReader.getInt32(0);
        }
        dataReader.close();
        // FollowUp
        dataReader = new SqlDB("select workshop, mail, livepdf, maxpdfLevel, communicationfinished, communicationfinsheddate, participants from followup_structure1 where orgID='" + OrgID + "'", projectID);
        if (dataReader.read())
        {
            workshop = dataReader.getBool(0);
            mailAdress = dataReader.getString(1);
            LifePDF = dataReader.getBool(2);
            maxpdfLevel = dataReader.getInt32(3);
            communicationfinished = dataReader.getBool(4);
            communicationfinsheddate = dataReader.getDateTime(5);
            numberOfParticipants = dataReader.getInt32(6);
        }
        dataReader.close();
        // Download
        // Response
        dataReader = new SqlDB("select number from response_structure1 where orgID='" + OrgID + "'", projectID);
        if (dataReader.read())
        {
            numberIst = dataReader.getInt32(0);
        }
        dataReader.close();

        // Untereinleiten neu laden
        if (expandedNav)
        {
            foreach (TAdminStructure1 tempStructure1 in childs)
            {
                tempStructure1.Reload();
            }
        }
    }
    /// <summary>
    /// Einlsen der nächsten Ebene untergeordneter Organisationseinheiten
    /// </summary>
    public void getChilds(string aProjectID)
    {
        awayChilds();
        SqlDB dataReader;
        dataReader = new SqlDB("select orgID from structure1 where topOrgID='" + this.OrgID + "'", projectID);
        while (dataReader.read())
        {
            TAdminStructure1 newChild = new TAdminStructure1(dataReader.getInt32(0), this, this.level + 1, aProjectID);
            this.childs.Add(newChild);
        }
        dataReader.close();
    }
    /// <summary>
    /// Freigeben aller untergeordneten Organisationseinheiten
    /// </summary>
    public void awayChilds()
    {
        // ausblenden der nächsten Ebene der untergeordneten Elemente
        // nur ausblenden wenn im zugeordneten Bereich
        if (level >= 0)
            childs.Clear();
    }
    /// <summary>
    /// Ermittlung der OrgID aller untergeordneter Organisationseinheiten über alle Level; die Start-Organisationseinheit wird nciht eingeschlossen
    /// </summary>
    /// <param name="aOrgID">OrgID der Start-Organisationseinheit</param>
    /// <param name="aOrgIDList">Liste der OrgID</param>
    /// <param name="aProjectID">Eindeutige ID des Projektes</param>
    public static void getOrgIDList(int aOrgID, ref ArrayList aOrdIDList, string aProjectID)
    {
        // schreibt die OrgIDs aller Untereinheten von aOrgID in aOrgIDList
        SqlDB dataReader;
        dataReader = new SqlDB("select orgID from structure1 where topOrgID='" + aOrgID + "'", aProjectID);
        while (dataReader.read())
        {
            int childOrgID = dataReader.getInt32(0);
            aOrdIDList.Add(childOrgID);
            // der Aufruf wird rekursiv ausgeführt
            getOrgIDList(childOrgID, ref aOrdIDList, aProjectID);
        }
        dataReader.close();
    }
    /// <summary>
    /// Rekursive Berechung der Summen der Soll- und Ist-Summen auf Basis der Soll- und Ist-Zahlen der einzelen Organisationseinheiten
    /// </summary>
    /// <param name="aOrgID">OrgID der zu berechnenden Organisationseinheit</param>
    /// <param name="aSollOrg">Rückgabeparamter dür Soll-Summe</param>
    /// <param name="aIstRes">Rückgabeparameter für Ist-Summe</param>
    /// <param name="aProjectID">Eindeutige ID des Projektes</param>
    protected static void getCounter(int aOrgID, ref int aSollOrg, ref int aIstRes, ref int aIstOrgInputEmployee, ref int aSollRes, string aProjectID)
    {
        SqlDB dataReader;
        dataReader = new SqlDB("select orgID from structure1 where topOrgID='" + aOrgID.ToString() + "'", aProjectID);
        while (dataReader.read())
        {
            int singleSollOrg = 0;
            int singleSollRes = 0;
            int singleIstRes = 0;
            int singleIstOrgInputEmployee = 0;
            // rekursiver Aufruf
            getCounter(dataReader.getInt32(0), ref singleSollOrg, ref singleIstRes, ref singleIstOrgInputEmployee, ref singleSollRes, aProjectID);
            aSollOrg = aSollOrg + singleSollOrg;
            aSollRes = aSollRes + singleSollRes;
            aIstRes = aIstRes + singleIstRes;
            aIstOrgInputEmployee = aIstOrgInputEmployee + singleIstOrgInputEmployee;
        }
        dataReader.close();
        // Werte der Einheit ermitteln und hinzuaddieren
        dataReader = new SqlDB("select number, numberSumInput from orgmanager_structure1 where orgID='" + aOrgID.ToString() + "'", aProjectID);
        if (dataReader.read())
        {
            aSollOrg = aSollOrg + dataReader.getInt32(0);
        }
        dataReader.close();
        dataReader = new SqlDB("select count(employeeID) from orgmanager_employee where orgIDAdd1='" + aOrgID.ToString() + "'", aProjectID);
        if (dataReader.read())
        {
            aIstOrgInputEmployee = aIstOrgInputEmployee + dataReader.getInt32(0);
        }
        dataReader.close();
        dataReader = new SqlDB("select number, target from response_structure1 where orgID='" + aOrgID.ToString() + "'", aProjectID);
        if (dataReader.read())
        {
            aIstRes = aIstRes + dataReader.getInt32(0);
            aSollRes = aSollRes + dataReader.getInt32(1);
        }
        dataReader.close();
        // Summen in DB zurückschreiben
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("numberSum", "int", aSollOrg.ToString());
        parameterList.addParameter("numberSumInput", "int", aIstOrgInputEmployee.ToString());
        parameterList.addParameter("orgID", "int", aOrgID.ToString());
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("UPDATE orgmanager_structure1 SET numbersum=@numbersum, numbersumInput=@numbersumInput WHERE orgID=@orgID", parameterList);
        parameterList = new TParameterList();
        parameterList.addParameter("numberSum", "int", aIstRes.ToString());
        parameterList.addParameter("targetSum", "int", aSollRes.ToString());
        parameterList.addParameter("orgID", "int", aOrgID.ToString());
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("UPDATE response_structure1 SET numberSum=@numberSum, targetSum=@targetSum WHERE orgID=@orgID", parameterList);
    }
    /// <summary>
    /// Neuberechung der Summen der Soll- und Ist-Zahlen
    /// </summary>
    /// <param name="aProjectID"></param>
    public static void recalc(string aProjectID)
    {
        // Initialisierung der Summen
        int sumSollOrg = 0;
        int sumSollRes = 0;
        int sumIstRes = 0;
        int sumIstOrgInputEmployee = 0;
        // Root-Object finden
        SqlDB dataReader;
        dataReader = new SqlDB("select orgID from structure1 where topOrgID='0'", aProjectID);
        if (dataReader.read())
        {
            int rootID = dataReader.getInt32(0);
            getCounter(rootID, ref sumSollOrg, ref sumIstRes, ref sumIstOrgInputEmployee, ref sumSollRes, aProjectID);
        }
        dataReader.close();
    }

    public string getValue(string aKey)
    {
        string result = "";
        bool found = false;
        int index = 0;
        while ((index < values.Count) && (!found))
        {
            if (((TKeyValue)values[index]).key == aKey)
            {
                result = ((TKeyValue)values[index]).value;
                found = true;
            }
            index++;
        }
        return result;
    }
    public ArrayList getValueList(string aKey)
    {
        ArrayList result = new ArrayList();
        int index = 0;
        while (index < values.Count)
        {
            if (((TKeyValue)values[index]).key == aKey)
            {
                result.Add(((TKeyValue)values[index]).value);
            }
            index++;
        }
        return result;
    }
    /// <summary>
    /// Ermittlung der OrgID aller übergeordnetrer Organisationseinheiten über alle Level; die Start-Organisationseinheit wird eingeschlossen
    /// </summary>
    /// <param name="aOrgID">OrgID der Start-Organisationseinheit</param>
    /// <param name="aOrgIDList">Liste der OrgID</param>
    /// <param name="aProjectID">Eindeutige ID des Projektes</param>
    public static bool getTopOrgIDList(int aOrgID, ref ArrayList aOrgIDList, string aProjectID)
    {
        bool result = aOrgID != 0;
        // schreibt die OrgIDs aller Überreinheiten von aOrgID in aOrgIDList
        aOrgIDList = new ArrayList();
        aOrgIDList.Add(aOrgID);
        SqlDB dataReader;
        while (aOrgID != 0)
        {
            dataReader = new SqlDB("select topOrgID from structure1 where orgID='" + aOrgID + "'", aProjectID);
            if (dataReader.read())
            {
                aOrgID = dataReader.getInt32(0);
                if (aOrgID != 0)
                    aOrgIDList.Insert(0, aOrgID);
            }
            else
            {
                // TopOrgID existiert nicht -> Pfad ist ungültig
                result = false;
                aOrgID = 0;
            }
            dataReader.close();
        }
        return result;
    }
    public static int getTopOrgID(int aOrgID, string aProjectID)
    {
        int result = 0;
        SqlDB dataReader;
        dataReader = new SqlDB("SELECT topOrgID FROM structure1 where orgID='" + aOrgID.ToString() + "'", aProjectID);
        if (dataReader.read())
        {
            result = dataReader.getInt32(0);
        }
        dataReader.close();
        return result;
    }
    public static void getOrgName(int aOrgID, ref string displayName, ref string displayNameShort, string aProjectID)
    {
        displayName = "";
        displayNameShort = "";
        SqlDB dataReader;
        dataReader = new SqlDB("select displayname, displaynameshort from structure1 where orgID='" + aOrgID.ToString() + "'", aProjectID);
        if (dataReader.read())
        {
            displayName = dataReader.getString(0);
            displayNameShort = dataReader.getString(1);
        }
        dataReader.close();
    }
    public static bool structureExits(int aOrgID, string aTablename, string aProjectID)
    {
        bool result = false;

        SqlDB dataReader;
        dataReader = new SqlDB("select orgID from " + aTablename + " WHERE orgID='" + aOrgID + "'", aProjectID);
        if (dataReader.read())
        {
            result = true;
        }
        return result;
    }
}