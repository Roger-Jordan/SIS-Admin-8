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
/// Zusammenfassungsbeschreibung für TUser
/// </summary>
public class TEmployee
{
    public class TKeyValue
    {
        public string key;
        public string value;
    }

    public int EmployeeID;				// eindeutige Kennung des Mitarbeiters
    public string Name;
    public string Firstname;
    public string Email;
    public string Mobile;
    public int surveyType;
    public int orgID;
    public int orgIDAdd1;
    public string orgIDAdd1Text;
    public int orgIDAdd2;
    public string orgIDAdd2Text;
    public int orgIDAdd3;
    public string orgIDAdd3Text;
    public bool isSupervisor;
    public bool isSupervisor1;
    public bool isSupervisor2;
    public bool isSupervisor3;
    public int orgIDOrigin;
    public int orgIDAdd1Origin;
    public int orgIDAdd2Origin;
    public int orgIDAdd3Origin;
    public bool isBounce;
    public string accesscode;
    public ArrayList values;        // alle Werte der variablen Felder

    public TEmployee()
    {
        Name = "";
        Firstname = "";
        Email = "";
        Mobile = "";
        surveyType = 0;
        orgID = 0;
        orgIDAdd1 = 0;
        orgIDAdd2Text = ""; ;
        orgIDAdd2 = 0;
        orgIDAdd2Text = "";
        orgIDAdd3 = 0;
        orgIDAdd3Text = "";
        isSupervisor = false;
        isSupervisor1 = false;
        isSupervisor2 = false;
        isSupervisor3 = false;
        orgIDOrigin = 0;
        orgIDAdd1Origin = 0;
        orgIDAdd2Origin = 0;
        orgIDAdd3Origin = 0;
        isBounce = false;
        accesscode = "";
        values = new ArrayList();
    }
    public TEmployee(int aEmployeeID, string aProjectID)
    {
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("employeeID", "int", aEmployeeID.ToString());
        SqlDB dataReader;
        dataReader = new SqlDB("select employeeID, name, firstname, email, mobile, surveytype, orgID, orgIDAdd1, orgIDAdd2, orgIDAdd3, orgIDOrigin, orgIDAdd1Origin, orgIDAdd2Origin, orgIDAdd3Origin, issupervisor, issupervisor1, issupervisor2, issupervisor3, isbounce, accesscode from orgmanager_employee where employeeID=@employeeID", parameterList, aProjectID);
        if (dataReader.read())
        {
            EmployeeID = aEmployeeID;
            Name = dataReader.getString(1);
            Firstname = dataReader.getString(2);
            Email = dataReader.getString(3);
            Mobile = dataReader.getString(4);
            surveyType = dataReader.getInt32(5);
            orgID = dataReader.getInt32(6);
            orgIDAdd1 = dataReader.getInt32(7);
            orgIDAdd2 = dataReader.getInt32(8);
            orgIDAdd3 = dataReader.getInt32(9);
            orgIDOrigin = dataReader.getInt32(10);
            orgIDAdd1Origin = dataReader.getInt32(11);
            orgIDAdd2Origin = dataReader.getInt32(12);
            orgIDAdd3Origin = dataReader.getInt32(13);
            isSupervisor = dataReader.getBool(14);
            isSupervisor1 = dataReader.getBool(15);
            isSupervisor2 = dataReader.getBool(16);
            isSupervisor3 = dataReader.getBool(17);
            isBounce = dataReader.getBool(18);
            accesscode = dataReader.getString(19);

            SqlDB dataReader1;

            dataReader1 = new SqlDB("select displayName from structure1 where orgID='" + orgIDAdd1.ToString() + "'", aProjectID);
            if (dataReader1.read())
            {
                orgIDAdd1Text = dataReader1.getString(0);
            }
            dataReader1.close();

            dataReader1 = new SqlDB("select displayName from structure2 where orgID='" + orgIDAdd2.ToString() + "'", aProjectID);
            if (dataReader1.read())
            {
                orgIDAdd2Text = dataReader1.getString(0);
            }
            dataReader1.close();

            dataReader1 = new SqlDB("select displayName from structure3 where orgID='" + orgIDAdd3.ToString() + "'", aProjectID);
            if (dataReader1.read())
            {
                orgIDAdd3Text = dataReader1.getString(0);
            }
            dataReader1.close();

            values = new ArrayList();
            dataReader1 = new SqlDB("SELECT fieldID, value FROM orgmanager_employeefieldsvalues WHERE employeeID='" + aEmployeeID.ToString() + "'", aProjectID);
            while (dataReader1.read())
            {
                TKeyValue tempKeyValue = new TKeyValue();
                tempKeyValue.key = dataReader1.getString(0);
                tempKeyValue.value = dataReader1.getString(1);
                values.Add(tempKeyValue);
            }
            dataReader1.close();
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
    public static bool employeeExits(int aEmployeeID, string aProjectID)
    {
        bool result = false;

        SqlDB dataReader;
        dataReader = new SqlDB("select employeeID from orgmanager_emplyoee WHERE employeeID='" + aEmployeeID + "'", aProjectID);
        if (dataReader.read())
        {
            result = true;
        }
        return result;
    }
}

