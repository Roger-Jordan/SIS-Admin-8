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
/// Zusammenfassungsbeschreibung für TProject
/// </summary>
public class TEmployeeFieldsList
{
    public class TEmployeeFieldsListEntry
    {
        public string fieldID;
        public int division;
        public int positionRow;
        public int positionCol;
        public string fieldType;
        public bool madantory;
        public int maxchar;
        public int width;
        public int rows;
        public string minselect;
        public string maxselect;
        public string regex;
        public bool show;
    }

    public TEmployeeFieldsListEntry[,,] employeeFieldMatrix;
    public ArrayList employeeFieldList;
    public int divisionCount;
    public int colCount;
    public int rowCount;

    public int getMax(string aColumn, string aProjectID)
    {
        int result = 0;

        SqlDB dataReader;
        dataReader = new SqlDB("select COUNT(" + aColumn + ") from orgmanager_employeefields", aProjectID);
        if (dataReader.read())
        {
            if (dataReader.getInt32(0) > 0)
            {
                SqlDB dataReader1;
                dataReader1 = new SqlDB("select MAX(" + aColumn + ") from orgmanager_employeefields", aProjectID);
                if (dataReader1.read())
                {
                    result = dataReader1.getInt32(0) + 1;
                }
                dataReader1.close();
            }
        }
        dataReader.close();

        return result;
    }
    public TEmployeeFieldsList(string aProjectID)
    {
        SqlDB dataReader;

        // maximale Gruppenanzahl ermitteln
        divisionCount = getMax("division", aProjectID);
        // maximale Spaltenanzahl ermitteln
        colCount = getMax("positionCol", aProjectID);
        // maximale Zeilenanzahl ermitteln
        rowCount = getMax("positionRow", aProjectID);

        employeeFieldMatrix = new TEmployeeFieldsListEntry[divisionCount, rowCount, colCount];
        employeeFieldList = new ArrayList();

        dataReader = new SqlDB("select fieldID, division, positionRow, positionCol, fieldType, mandatory, maxchar, width, rows, minvalue, maxvalue, regex from orgmanager_employeefields ORDER BY positionRow, positionCol", aProjectID);
        while (dataReader.read())
        {
            TEmployeeFieldsListEntry tempEntry = new TEmployeeFieldsListEntry();
            tempEntry.fieldID = dataReader.getString(0);
            tempEntry.division = dataReader.getInt32(1);
            tempEntry.positionRow = dataReader.getInt32(2);
            tempEntry.positionCol = dataReader.getInt32(3);
            tempEntry.fieldType = dataReader.getString(4);
            tempEntry.madantory = dataReader.getBool(5);
            tempEntry.maxchar = dataReader.getInt32(6);
            tempEntry.width = dataReader.getInt32(7);
            tempEntry.rows = dataReader.getInt32(8);
            tempEntry.minselect = dataReader.getString(9);
            tempEntry.maxselect = dataReader.getString(10);
            tempEntry.regex = dataReader.getString(11);
            tempEntry.show = true;

            employeeFieldMatrix[tempEntry.division, tempEntry.positionRow, tempEntry.positionCol] = tempEntry;
            employeeFieldList.Add(tempEntry);
        }
        dataReader.close();
        
    }
    public TEmployeeFieldsList(int aOrgID, TEmployeeFieldsFilter aEmployeeFieldsFilterList, ArrayList aActOrgIDList, string aProjectID)
    {
        SqlDB dataReader;

        // maximale Gruppenanzahl ermitteln
        divisionCount = getMax("division", aProjectID);
        // maximale Spaltenanzahl ermitteln
        colCount = getMax("positionCol", aProjectID);
        // maximale Zeilenanzahl ermitteln
        rowCount = getMax("positionRow", aProjectID);

        employeeFieldMatrix = new TEmployeeFieldsListEntry[divisionCount, rowCount, colCount];
        employeeFieldList = new ArrayList();

        dataReader = new SqlDB("select fieldID, division, positionRow, positionCol, fieldType, mandatory, maxchar, width, rows, minvalue, maxvalue, regex from orgmanager_employeefields ORDER BY positionRow, positionCol", aProjectID);
        while (dataReader.read())
        {
            TEmployeeFieldsListEntry tempEntry = new TEmployeeFieldsListEntry();
            tempEntry.fieldID = dataReader.getString(0);
            tempEntry.division = dataReader.getInt32(1);
            tempEntry.positionRow = dataReader.getInt32(2);
            tempEntry.positionCol = dataReader.getInt32(3);
            tempEntry.fieldType = dataReader.getString(4);
            tempEntry.madantory = dataReader.getBool(5);
            tempEntry.maxchar = dataReader.getInt32(6);
            tempEntry.width = dataReader.getInt32(7);
            tempEntry.rows = dataReader.getInt32(8);
            tempEntry.minselect = dataReader.getString(9);
            tempEntry.maxselect = dataReader.getString(10);
            tempEntry.regex = dataReader.getString(11);

            //ArrayList actOrgIDList = TStructure.getOrgIDList(aOrgID, aProjectID);
            string actRight = "allow";
            // Schleife über alle OrgIDs von aktueller bis Root, solange kein Eintrag gefunden wurde
            int actIndex = 0;
            while ((actIndex < aActOrgIDList.Count) && ((int)aActOrgIDList[actIndex] != 0))
            {
                string tempRight = aEmployeeFieldsFilterList.getRight(tempEntry.fieldID, (int)aActOrgIDList[actIndex]);
                if (tempRight != "")
                {
                    actRight = tempRight;
                    // suche beenden, da Eintrag gefunden und Einträge darüber nicht mehr relevant
                    actIndex = aActOrgIDList.Count;
                }
                else
                {
                    // zur nächsthöheren OrgID weitergehen
                    actIndex++;
                }
            }
            tempEntry.show = actRight == "allow";

            employeeFieldMatrix[tempEntry.division, tempEntry.positionRow, tempEntry.positionCol] = tempEntry;
            if (tempEntry.show)
                employeeFieldList.Add(tempEntry);
        }
        dataReader.close();

    }
    public string getFieldID(string aFieldType)
    {
        string result = "";
        for (int divIndex = 0; divIndex < divisionCount; divIndex++)
        {
            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                for (int colIndex = 0; colIndex < colCount; colIndex++)
                {
                    if ((employeeFieldMatrix[divIndex, rowIndex, colIndex] != null) && (employeeFieldMatrix[divIndex, rowIndex, colIndex].fieldType == aFieldType))
                        result = employeeFieldMatrix[divIndex, rowIndex, colIndex].fieldID;
                }
            }
        }
        return result;
    }

    public static TEmployeeFieldsListEntry getEntry(int aRow, int aCol, string aProjectID)
    {
        TEmployeeFieldsListEntry result = null;

        SqlDB dataReader;
        dataReader = new SqlDB("select fieldID, division, positionRow, positionCol, fieldType, mandatory, maxchar, width, rows, minvalue, maxvalue, regex from orgmanager_employeefields WHERE positionRow='" + aRow.ToString() + "' AND positionCol='" + aCol.ToString() + "'", aProjectID);
        if (dataReader.read())
        {
            result = new TEmployeeFieldsListEntry();
            result.fieldID = dataReader.getString(0);
            result.division = dataReader.getInt32(1);
            result.positionRow = dataReader.getInt32(2);
            result.positionCol = dataReader.getInt32(3);
            result.fieldType = dataReader.getString(4);
            result.madantory = dataReader.getBool(5);
            result.maxchar = dataReader.getInt32(6);
            result.width = dataReader.getInt32(7);
            result.rows = dataReader.getInt32(8);
            result.minselect = dataReader.getString(9);
            result.maxselect = dataReader.getString(10);
            result.regex = dataReader.getString(11);
        }
        dataReader.close();

        return result;
    }
    public static TEmployeeFieldsListEntry getEntry(string aFieldID, string aProjectID)
    {
        TEmployeeFieldsListEntry result = null;

        SqlDB dataReader;
        dataReader = new SqlDB("select fieldID, division, positionRow, positionCol, fieldType, mandatory, maxchar, width, rows, minvalue, maxvalue, regex from orgmanager_employeefields WHERE fieldID='" + aFieldID + "'", aProjectID);
        if (dataReader.read())
        {
            result = new TEmployeeFieldsListEntry();
            result.fieldID = dataReader.getString(0);
            result.division = dataReader.getInt32(1);
            result.positionRow = dataReader.getInt32(2);
            result.positionCol = dataReader.getInt32(3);
            result.fieldType = dataReader.getString(4);
            result.madantory = dataReader.getBool(5);
            result.maxchar = dataReader.getInt32(6);
            result.width = dataReader.getInt32(7);
            result.rows = dataReader.getInt32(8);
            result.minselect = dataReader.getString(9);
            result.maxselect = dataReader.getString(10);
            result.regex = dataReader.getString(11);
        }
        dataReader.close();

        return result;
    }
}
