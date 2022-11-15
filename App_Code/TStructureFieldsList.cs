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
public class TStructureFieldsList
{
    public class TStructureFieldsListEntry
    {
        public string fieldID;
        public int division;
        public int positionRow;
        public int positionCol;
        public string fieldType;
        public bool isRecipient;
        public bool madantory;
        public int maxchar;
        public int width;
        public int rows;
        public string minselect;
        public string maxselect;
        public string regex;
        public bool show;
    }

    public TStructureFieldsListEntry[,,] structureFieldMatrix;
    public ArrayList structureFieldList;
    public int divisionCount;
    public int colCount;
    public int rowCount;

    public int getMax(string aColumn, string aTable, string aProjectID)
    {
        int result = 0;

        SqlDB dataReader;
        dataReader = new SqlDB("select COUNT(" + aColumn + ") from " + aTable, aProjectID);
        if (dataReader.read())
        {
            if (dataReader.getInt32(0) > 0)
            {
                SqlDB dataReader1;
                dataReader1 = new SqlDB("select MAX(" + aColumn + ") from " + aTable, aProjectID);
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
    public TStructureFieldsList(ref bool aShowEmployeeCount, ref bool aShowEmployeeCountTotal, ref bool aShowBackComparison, ref bool aShowLanguageUser, string aTable, string aProjectID)
    {
        SqlDB dataReader;

        // maximale Gruppenanzahl ermitteln
        divisionCount = getMax("division", aTable, aProjectID);
        // maximale Spaltenanzahl ermitteln
        colCount = getMax("positionCol", aTable, aProjectID);
        // maximale Zeilenanzahl ermitteln
        rowCount = getMax("positionRow", aTable, aProjectID);

        structureFieldMatrix = new TStructureFieldsListEntry[divisionCount, rowCount, colCount];
        structureFieldList = new ArrayList();

        aShowEmployeeCount = false;
        aShowEmployeeCountTotal = false;
        aShowBackComparison = false;
        aShowLanguageUser = false;
        dataReader = new SqlDB("select fieldID, division, positionRow, positionCol, fieldType, mandatory, recipient, maxchar, width, rows, minvalue, maxvalue, regex from " + aTable + " ORDER BY positionRow, positionCol", aProjectID);
        while (dataReader.read())
        {
            TStructureFieldsListEntry tempEntry = new TStructureFieldsListEntry();
            tempEntry.fieldID = dataReader.getString(0);
            tempEntry.division = dataReader.getInt32(1);
            tempEntry.positionRow = dataReader.getInt32(2);
            tempEntry.positionCol = dataReader.getInt32(3);
            tempEntry.fieldType = dataReader.getString(4);
            tempEntry.madantory = dataReader.getBool(5);
            tempEntry.isRecipient = dataReader.getBool(6);
            tempEntry.maxchar = dataReader.getInt32(7);
            tempEntry.width = dataReader.getInt32(8);
            tempEntry.rows = dataReader.getInt32(9);
            tempEntry.minselect = dataReader.getString(10);
            tempEntry.maxselect = dataReader.getString(11);
            tempEntry.regex = dataReader.getString(12);
            tempEntry.show = true;

            if (tempEntry.fieldType == "EmployeeCount")
                aShowEmployeeCount = true;
            if (tempEntry.fieldType == "EmployeeCountTotal")
                aShowEmployeeCountTotal = true;
            if (tempEntry.fieldType == "BackComparison")
                aShowBackComparison = true;
            if (tempEntry.fieldType == "LanguageUser")
                aShowLanguageUser = true;

            structureFieldMatrix[tempEntry.division, tempEntry.positionRow, tempEntry.positionCol] = tempEntry;
            structureFieldList.Add(tempEntry);
        }
        dataReader.close();
        
    }
    public TStructureFieldsList(int aOrgID, TStructureFieldsFilter aStructureFieldsFilterList, ArrayList aActOrgIDList, string aTable, string aProjectID)
    {
        SqlDB dataReader;

        // maximale Gruppenanzahl ermitteln
        divisionCount = getMax("division", aTable, aProjectID);
        // maximale Spaltenanzahl ermitteln
        colCount = getMax("positionCol", aTable, aProjectID);
        // maximale Zeilenanzahl ermitteln
        rowCount = getMax("positionRow", aTable, aProjectID);

        structureFieldMatrix = new TStructureFieldsListEntry[divisionCount, rowCount, colCount];
        structureFieldList = new ArrayList();

        dataReader = new SqlDB("select fieldID, division, positionRow, positionCol, fieldType, recipient, mandatory, maxchar, width, rows, minvalue, maxvalue, regex from " + aTable + " ORDER BY positionRow, positionCol", aProjectID);
        while (dataReader.read())
        {
            TStructureFieldsListEntry tempEntry = new TStructureFieldsListEntry();
            tempEntry.fieldID = dataReader.getString(0);
            tempEntry.division = dataReader.getInt32(1);
            tempEntry.positionRow = dataReader.getInt32(2);
            tempEntry.positionCol = dataReader.getInt32(3);
            tempEntry.fieldType = dataReader.getString(4);
            tempEntry.isRecipient = dataReader.getBool(5);
            tempEntry.madantory = dataReader.getBool(6);
            tempEntry.maxchar = dataReader.getInt32(7);
            tempEntry.width = dataReader.getInt32(8);
            tempEntry.rows = dataReader.getInt32(9);
            tempEntry.minselect = dataReader.getString(10);
            tempEntry.maxselect = dataReader.getString(11);
            tempEntry.regex = dataReader.getString(12);

            string actRight = "allow";
            // Schleife über alle OrgIDs von aktueller bis Root, solange kein Eintrag gefunden wurde
            int actIndex = 0;
            while ((actIndex < aActOrgIDList.Count) && ((int)aActOrgIDList[actIndex] != 0))
            {
                string tempRight = aStructureFieldsFilterList.getRight(tempEntry.fieldID, (int)aActOrgIDList[actIndex]);
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

            structureFieldMatrix[tempEntry.division, tempEntry.positionRow, tempEntry.positionCol] = tempEntry;
            if (tempEntry.show)
                structureFieldList.Add(tempEntry);
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
                    if ((structureFieldMatrix[divIndex, rowIndex, colIndex] != null) && (structureFieldMatrix[divIndex, rowIndex, colIndex].fieldType == aFieldType))
                        result = structureFieldMatrix[divIndex, rowIndex, colIndex].fieldID;
                }
            }
        }
        return result;
    }

    public static TStructureFieldsListEntry getEntry(int aDivision, int aRow, int aCol, string aTable, string aProjectID)
    {
        TStructureFieldsListEntry result = null;

        SqlDB dataReader;
        dataReader = new SqlDB("select fieldID, division, positionRow, positionCol, fieldType, mandatory, recipient, maxchar, width, rows, minvalue, maxvalue, regex from " + aTable + " WHERE positionRow='" + aRow.ToString() + "' AND positionCol='" + aCol.ToString() + "'", aProjectID);
        if (dataReader.read())
        {
            result = new TStructureFieldsListEntry();
            result.fieldID = dataReader.getString(0);
            result.division = dataReader.getInt32(1);
            result.positionRow = dataReader.getInt32(2);
            result.positionCol = dataReader.getInt32(3);
            result.fieldType = dataReader.getString(4);
            result.madantory = dataReader.getBool(5);
            result.isRecipient = dataReader.getBool(6);
            result.maxchar = dataReader.getInt32(7);
            result.width = dataReader.getInt32(8);
            result.rows = dataReader.getInt32(9);
            result.minselect = dataReader.getString(10);
            result.maxselect = dataReader.getString(11);
            result.regex = dataReader.getString(12);
        }
        dataReader.close();

        return result;
    }
    public static TStructureFieldsListEntry getEntry(string aFieldID, string aTable, string aProjectID)
    {
        TStructureFieldsListEntry result = null;

        SqlDB dataReader;
        dataReader = new SqlDB("select fieldID, division, positionRow, positionCol, fieldType, mandatory, recipient, maxchar, width, rows, minvalue, maxvalue, regex from " + aTable + " WHERE fieldID='" + aFieldID + "'", aProjectID);
        if (dataReader.read())
        {
            result = new TStructureFieldsListEntry();
            result.fieldID = dataReader.getString(0);
            result.division = dataReader.getInt32(1);
            result.positionRow = dataReader.getInt32(2);
            result.positionCol = dataReader.getInt32(3);
            result.fieldType = dataReader.getString(4);
            result.madantory = dataReader.getBool(5);
            result.isRecipient = dataReader.getBool(6);
            result.maxchar = dataReader.getInt32(7);
            result.width = dataReader.getInt32(8);
            result.rows = dataReader.getInt32(9);
            result.minselect = dataReader.getString(10);
            result.maxselect = dataReader.getString(11);
            result.regex = dataReader.getString(12);
        }
        dataReader.close();

        return result;
    }
}
