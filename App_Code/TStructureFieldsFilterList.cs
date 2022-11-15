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
/// Zusammenfassungsbeschreibung für TAddValue1List
/// </summary>
public class TStructureFieldsFilter
{
    public class TEntry
    {
        public string fieldID; 
        public int orgID;      
        public string righttype;
    }
    public ArrayList filterList;

    // erzeugt neues Objekt
    public TStructureFieldsFilter(string aTable, string aProjectID)
    {
        filterList = new ArrayList();
        SqlDB dataReader;
        dataReader = new SqlDB("select fieldID, orgID, righttype FROM " + aTable + " ORDER BY fieldID",aProjectID);
        while (dataReader.read())
        {
            TEntry TempEntry = new TEntry();
            TempEntry.fieldID = dataReader.getString(0);
            TempEntry.orgID = dataReader.getInt32(1);
            TempEntry.righttype = dataReader.getString(2);
            filterList.Add(TempEntry);
        }
        dataReader.close();
    }
    private ArrayList getFieldFilterList(string aFieldID)
    {
        ArrayList Result = new ArrayList();
        int i;
        for (i = 0; i < filterList.Count; i++)
        {
            if (((TEntry)filterList[i]).fieldID == aFieldID)
            {
                Result.Add(((TEntry)filterList[i]).orgID);
            }
        }
        return Result;
    }
    public ArrayList getOrgIDFilterList(int aOrgID)
    {
        ArrayList Result = new ArrayList();
        int i;
        for (i = 0; i < filterList.Count; i++)
        {
            if (((TEntry)filterList[i]).orgID == aOrgID)
            {
                Result.Add(filterList[i]);
            }
        }
        return Result;
    }
    public static ArrayList getOrgIDFilterListComplete(TStructureFieldsFilter aFilterList, int aOrgID, string aProjectID)
    {
        ArrayList Result = new ArrayList();

        // Liste der OrgIDs dieser Einheit bis Root ermitteln
        ArrayList actOrgIDList = new ArrayList();
        TAdminStructure.getOrgIDList(aOrgID, ref actOrgIDList, aProjectID);

        // Schleife über alle OrgIDs von aktueller Einheit bis Root
        while (actOrgIDList.Count > 0)
        {
            // FilterListe einer Einheit ermitteln
            int tempOrgID = (int)actOrgIDList[actOrgIDList.Count - 1];
            ArrayList tempFilterList = aFilterList.getOrgIDFilterList(tempOrgID);

            // Filterliste mit Ergebnisliste zusammenführen -> wenn noch kein Eintrag für ein Field vorhanden, dann hinzufügen
            foreach (TStructureFieldsFilter.TEntry tempEntry in tempFilterList)
            {
                // Test, ob in Ergebnisliste bereits vorhanden
                bool exists = false;
                foreach (TStructureFieldsFilter.TEntry resultEntry in Result)
                {
                    if (resultEntry.fieldID == tempEntry.fieldID)
                        exists = true;
                }
                if (!exists)
                {
                    Result.Add(tempEntry);
                }
            }
            actOrgIDList.RemoveAt(actOrgIDList.Count - 1);
        }
        return Result;
    }
    public string getRight(string aFieldID, int aOrgID)
    {
        string Result = "";
        int i;
        for (i = 0; i < filterList.Count; i++)
        {
            if ((((TEntry)filterList[i]).fieldID == aFieldID) && (((TEntry)filterList[i]).orgID == aOrgID))
            {
                Result = (((TEntry)filterList[i]).righttype);
            }
        }
        return Result;
    }
}
