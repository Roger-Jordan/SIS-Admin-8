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
public class TAccessCode
{
    string projectID;
    public int aGroup1Counter;
    public int aGroup2Counter;
    public int aGroup3Counter;
    public int aGroup4Counter;
    public string aGroup1;
    public string aGroup2;
    public string aGroup3;
    public string aGroup4;

    public TAccessCode(string aProjectID)
    {
        projectID = aProjectID;
        // Code-Definition ermitteln und anzeigen
        SqlDB dataReader = new SqlDB(aProjectID);
        aGroup1Counter = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='AccessCodeCounter1'").intValue;
        aGroup2Counter = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='AccessCodeCounter2'").intValue;
        aGroup3Counter = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='AccessCodeCounter3'").intValue;
        aGroup4Counter = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='AccessCodeCounter4'").intValue;
        aGroup1 = dataReader.getScalarString("SELECT text FROM orgmanager_config WHERE ID='AccessCodeTextGroup1'").stringValue;
        aGroup2 = dataReader.getScalarString("SELECT text FROM orgmanager_config WHERE ID='AccessCodeTextGroup2'").stringValue;
        aGroup3 = dataReader.getScalarString("SELECT text FROM orgmanager_config WHERE ID='AccessCodeTextGroup3'").stringValue;
        aGroup4 = dataReader.getScalarString("SELECT text FROM orgmanager_config WHERE ID='AccessCodeTextGroup4'").stringValue;
    }
    public string createAccessCode()
    {
        // Code erzeugen und auf Eindeutigkeit und unerlaubte führende Zeichen prüfen, ggf. verwerfen und neu erstellen
        Random random = new Random();
        int randomValue = 0;
        int randomPosition = 0;
        bool codeExists = false;
        string tempCode = "";
        do
        {
            codeExists = false;
            tempCode = "";
            // Anzahl Ziffern erzeugen
            for (int i = 0; i < aGroup1Counter; i++)
            {
                randomValue = random.Next(aGroup1.Length);
                randomPosition = random.Next(tempCode.Length);

                tempCode = tempCode.Insert(randomPosition, aGroup1[randomValue].ToString());
            }
            // Anzahl Buchstaben 1 erzeugen und an Zufallsposition einfügen
            for (int i = 0; i < aGroup2Counter; i++)
            {
                randomValue = random.Next(aGroup2.Length);
                randomPosition = random.Next(tempCode.Length);

                tempCode = tempCode.Insert(randomPosition, aGroup2[randomValue].ToString());
            }
            // Anzahl Buchstaben 2 erzeugen und an Zufallsposition einfügen
            for (int i = 0; i < aGroup3Counter; i++)
            {
                randomValue = random.Next(aGroup3.Length);
                randomPosition = random.Next(tempCode.Length);

                tempCode = tempCode.Insert(randomPosition, aGroup3[randomValue].ToString());
            }
            // Anzahl Sonderzeichen erzeugen und an Zufallsposition einfügen
            for (int i = 0; i < aGroup4Counter; i++)
            {
                randomValue = random.Next(aGroup4.Length);
                randomPosition = random.Next(tempCode.Length);

                tempCode = tempCode.Insert(randomPosition, aGroup4[randomValue].ToString());
            }
            SqlDB dataReader = new SqlDB("select accesscode from orgmanager_employee WHERE accesscode='" + tempCode + "'", projectID);
            if (dataReader.read())
            {
                codeExists = true;
            }
            dataReader.close();
        }
        while (codeExists || (tempCode[0] == '0') || (tempCode[0] == '+') || (tempCode[0] == '-'));

        return tempCode;
    }
    public string createAndSaveAccessCode(int aEmployeeID)
    { 
        string accessCode = createAccessCode();

        // Code eintragen
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("employeeID", "int", aEmployeeID.ToString());
        parameterList.addParameter("accesscode", "string", accessCode);
        SqlDB dataReader = new SqlDB(projectID);
        dataReader.execSQLwithParameter("UPDATE orgmanager_employee SET accesscode=@accesscode WHERE employeeID=@employeeID", parameterList);

        return accessCode;
    }
    public void setDefintion(int aCounter1, string aValue1, int aCounter2, string aValue2, int aCounter3, string aValue3, int aCounter4, string aValue4)
    {
        aGroup1Counter = aCounter1;
        aGroup2Counter = aCounter2;
        aGroup3Counter = aCounter3;
        aGroup4Counter = aCounter4;
        aGroup1 = aValue1;
        aGroup2 = aValue2;
        aGroup3 = aValue3;
        aGroup4 = aValue4;
        setValue(aCounter1, "AccessCodeCounter1");
        setValue(aCounter2, "AccessCodeCounter2");
        setValue(aCounter3, "AccessCodeCounter3");
        setValue(aCounter4, "AccessCodeCounter4");
        setValue(aValue1, "AccessCodeTextGroup1");
        setValue(aValue2, "AccessCodeTextGroup2");
        setValue(aValue3, "AccessCodeTextGroup3");
        setValue(aValue4, "AccessCodeTextGroup4");
    }
    protected void setValue(int aValue, string aKey)
    {
        SqlDB dataReader;
        dataReader = new SqlDB(projectID);
        dataReader.execSQL("DELETE FROM orgmanager_config WHERE ID='" + aKey + "'");
        dataReader = new SqlDB(projectID);
        dataReader.execSQL("INSERT INTO orgmanager_config (ID, value) VALUES ('" + aKey + "', '" + aValue + "')");
    }
    protected void setValue(string aValue, string aKey)
    {
        SqlDB dataReader;
        dataReader = new SqlDB(projectID);
        dataReader.execSQL("DELETE FROM orgmanager_config WHERE ID='" + aKey + "'");
        dataReader = new SqlDB(projectID);
        dataReader.execSQL("INSERT INTO orgmanager_config (ID, text) VALUES ('" + aKey + "', '" + aValue + "')");
    }
}

