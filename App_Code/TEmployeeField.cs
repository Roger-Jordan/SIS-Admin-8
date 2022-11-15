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
using System.Collections.Specialized;
using System.IO;
using System.Text;

/// <summary>
/// Klasse zur Verwaltung der Benutzers der Tools
/// </summary>
public class TEmployeeField
{
    public string FieldID;				// eindeutige Kennung des aktuellen Feldes
    public int Division;
    public int PositionRow;
    public int PositionCol;
    public string FieldType;
    public bool Mandatory;
    public int MaxChar;
    public int Width;
    public int Rows;
    public string MinValue;
    public string MaxValue;
    public string RegEx;

    /// <summary>
    /// Objekt erzeugen
    /// </summary>
    public TEmployeeField()
    {
    }
    /// <summary>
    /// Eigenschaften eines Benutzer aus der DB lesen
    /// </summary>
    /// <param name="aUserID">ID des Bentuzers</param>
    /// <param name="aProjectID">ID des Projektes</param>
    public TEmployeeField(string aFieldID, string aProjectID)
    {
        SqlDB dataReader;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("fieldID", "string", aFieldID);
        dataReader = new SqlDB("select fieldID, division, positionRow, positionCol, fieldType, mandatory, maxchar, width, rows, minValue, maxValue, regex FROM orgmanager_employeefields WHERE fieldID=@fieldID", parameterList, aProjectID);
        if (dataReader.read())
        {
            FieldID = dataReader.getString(0);
            Division = dataReader.getInt32(1);
            PositionRow = dataReader.getInt32(2);
            PositionCol = dataReader.getInt32(3);
            FieldType = dataReader.getString(4);
            Mandatory = dataReader.getBool(5);
            MaxChar = dataReader.getInt32(6);
            Width = dataReader.getInt32(7);
            Rows = dataReader.getInt32(8);
            MinValue = dataReader.getString(9);
            MaxValue = dataReader.getString(10);
            RegEx = dataReader.getString(11);
        }
        dataReader.close();
    }
    /// <summary>
    /// Speichern eines neuen Benutzers in der DB
    /// </summary>
    /// <param name="aProjectID">ID des Projektes</param>
    /// <returns></returns>
    public bool save(string aProjectID)
    {
        SqlDB dataReader;

        // fieldID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("fieldID", "string", this.FieldID);
        dataReader = new SqlDB("SELECT userID FROM orgmanager_employeefields WHERE fieldID=@fieldID", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            string tempMandatory = "0";
            if (Mandatory)
                tempMandatory = "1";

            parameterList = new TParameterList();
            parameterList.addParameter("fieldID", "string", FieldID);
            parameterList.addParameter("division", "int", Division.ToString());
            parameterList.addParameter("positionRow", "int", PositionRow.ToString());
            parameterList.addParameter("positionCol", "int", PositionCol.ToString());
            parameterList.addParameter("fieldType", "string", FieldType);
            parameterList.addParameter("mandatory", "int", tempMandatory);
            parameterList.addParameter("maxchar", "int", MaxChar.ToString());
            parameterList.addParameter("width", "int", Width.ToString());
            parameterList.addParameter("rows", "int", Rows.ToString());
            parameterList.addParameter("minValue", "string", MinValue.ToString());
            parameterList.addParameter("maxValue", "string", MaxValue.ToString());
            parameterList.addParameter("regex", "string", RegEx);
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO orgmanager_employeefields (fieldID, division, positionRow, positionCol, fieldType, mandatory, maxchar, width, rows, minValue, maxValue, regex) VALUES (@fieldID, @division, @positionRow, @positionCol, @fieldType, @mandatory, @maxchar, @width, @rows, @minValue, @maxValue, @regex)", parameterList);
            return true;
        }
        else
        {
            return false;
        }
    }
    /// <summary>
    /// Aktualisieren der Eigenschaften eines Benutzers in der DB
    /// </summary>
    /// <param name="aProjectID">ID des Projektes</param>
    public void update(string aProjectID)
    {
        SqlDB dataReader;

        string tempMandatory = "0";
        if (Mandatory)
            tempMandatory = "1";

        TParameterList parameterList = new TParameterList();
        parameterList = new TParameterList();
        parameterList.addParameter("fieldID", "string", FieldID);
        parameterList.addParameter("division", "int", Division.ToString());
        parameterList.addParameter("positionRow", "int", PositionRow.ToString());
        parameterList.addParameter("positionCol", "int", PositionCol.ToString());
        parameterList.addParameter("fieldType", "string", FieldType);
        parameterList.addParameter("mandatory", "int", tempMandatory);
        parameterList.addParameter("maxchar", "int", MaxChar.ToString());
        parameterList.addParameter("width", "int", Width.ToString());
        parameterList.addParameter("rows", "int", Rows.ToString());
        parameterList.addParameter("minValue", "string", MinValue.ToString());
        parameterList.addParameter("maxValue", "string", MaxValue.ToString());
        parameterList.addParameter("regex", "string", RegEx);
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("UPDATE orgmanager_employeefields SET division=@division, positionRow=@positionRow, positionCol=@positionCol, fieldType=@fieldType, mandatory=@mandatory, maxchar=@maxchar, width=@width, rows=@rows, minValue=@minValue, maxValue=@maxValue, regex=@regex WHERE fieldID=@fieldID", parameterList);
    }
    /// <summary>
    /// Löschen eines Benutzers aus der DB
    /// </summary>
    /// <param name="aUserID">ID des Benutzers</param>
    /// <param name="aProjectID">ID des Projektes</param>
    public static void delete(string aFieldID, string aProjectID)
    {
        SqlDB dataReader;
        // aus Datenbank löschen
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("fieldID", "string", aFieldID);
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("DELETE FROM orgmanager_employeefields WHERE fieldID=@fieldID", parameterList);
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("DELETE FROM orgmanager_employeefieldsfilter WHERE fieldID=@fieldID", parameterList);
    }
}

