using System.Collections;

/// <summary>
/// Zusammenfassungsbeschreibung für TStructure
///     Klasse zur Bearbeitung aller Eigenschaften einer Organisationseinheit,
///     unabhängig von der Zugehörigkeit zu einem Tool
/// </summary>
public class TAdminStructureBack3
{
    public bool expandedNav;		// Anzeige der Untereinheiten in der Navigation
    public int OrgID;			// eindeutige ID der Einheit
    public int TopOrgID;			// eindeutige ID der übergeordneten Einheit
    public string orgDisplayName;			// Text der Einheit
    public string orgDisplayNameShort;			// Text der Einheit
    public string filter;
    public TAdminStructureBack3 parent;		// Verweis auf übergeordneten Einheit
    public ArrayList childs;		// Liste von Verweisen auf untergeordnete Einheiten
    public int level;				// Level der aktuellen Einheit ist 0
    public bool hasChilds;
    public string projectID;
    public int maxLevel;

    public TAdminStructureBack3(int aOrgID, TAdminStructureBack3 aParent, int aLevel, string aProject)
    {
        if (aOrgID == -1)
        {
            // Root-Element nehmen;
            SqlDB dataReader1 = new SqlDB("SELECT OrgID from structureBack3 WHERE topOrgID='0'", aProject);
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

        // Namen der Unit aus der Datenbank lesen
        SqlDB dataReader;
        dataReader = new SqlDB("select orgID, topOrgID, displayName, displayNameShort, filter from structureBack3 where orgID='" + aOrgID + "'", aProject);
        if (dataReader.read())
        {
            TopOrgID = dataReader.getInt32(1);
            orgDisplayName = dataReader.getString(2);
            orgDisplayNameShort = dataReader.getString(3);
            filter = dataReader.getString(4);
        }
        dataReader.close();

        hasChilds = false;
        dataReader = new SqlDB("SELECT orgID from structureBack3 where topOrgID='" + aOrgID + "'", aProject);
        if (dataReader.read())
        {
            hasChilds = true;
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
        SqlDB dataReader = new SqlDB("select orgID, topOrgID, displayName, displayNameShort, filter from structureBack3 where orgID='" + OrgID + "'", projectID);
        if (dataReader.read())
        {
            orgDisplayName = dataReader.getString(2);
            orgDisplayNameShort = dataReader.getString(3);
            filter = dataReader.getString(4);
        }
        dataReader.close();

        // Untereinleiten neu laden
        if (expandedNav)
        {
            foreach (TAdminStructureBack3 tempStructure in childs)
            {
                tempStructure.Reload();
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
        dataReader = new SqlDB("select orgID from structureBack3 where topOrgID='" + this.OrgID + "'", projectID);
        while (dataReader.read())
        {
            TAdminStructureBack3 newChild = new TAdminStructureBack3(dataReader.getInt32(0), this, this.level + 1, aProjectID);
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
    /// <param name="aOrdIDList">Liste der OrgID</param>
    /// <param name="aProjectID">Eindeutige ID des Projektes</param>
    public static void getOrgIDList(int aOrgID, ref ArrayList aOrdIDList, string aProjectID)
    {
        // schreibt die OrgIDs aller Untereinheten von aOrgID in aOrgIDList
        SqlDB dataReader;
        dataReader = new SqlDB("select orgID from structureBack3 where topOrgID='" + aOrgID + "'", aProjectID);
        while (dataReader.read())
        {
            int childOrgID = dataReader.getInt32(0);
            aOrdIDList.Add(childOrgID);
            // der Aufruf wird rekursiv ausgeführt
            getOrgIDList(childOrgID, ref aOrdIDList, aProjectID);
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
    public static int getTopOrgID(int aOrgID, string aProjectID)
    {
        int result = 0;
        SqlDB dataReader;
        dataReader = new SqlDB("SELECT topOrgID FROM structureBack3 where orgID='" + aOrgID.ToString() + "'", aProjectID);
        if (dataReader.read())
        {
            result = dataReader.getInt32(0);
        }
        dataReader.close();
        return result;
    }
}