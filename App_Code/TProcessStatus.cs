using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;

/// <summary>
/// Zusammenfassungsbeschreibung für TProcessStatus
/// </summary>
public class TProcessStatus
{
    public class TProcessInfo
    {
        public string ID;          // Primärschlüssel
        public int status;
        public int actValue;
        public int minValue;
        public int maxValue;       
    }
    public int indexCounter;

    public TProcessStatus()
    {
    }
    public static TProcessInfo getProcessInfo(string aID, string aProjectID)
    {
        TProcessInfo actProcessInfo = null;

        SqlDB dataReader = new SqlDB("SELECT ID, status, actValue, minValue, maxValue FROM processStatus WHERE ID='" + aID + "'", aProjectID);
        if (dataReader.read())
        {
            actProcessInfo = new TProcessInfo();
            actProcessInfo.ID = aID;
            actProcessInfo.status = dataReader.getInt32(1);
            actProcessInfo.actValue = dataReader.getInt32(2);
            actProcessInfo.minValue = dataReader.getInt32(3);
            actProcessInfo.maxValue = dataReader.getInt32(4);
        }
        dataReader.close();

        return actProcessInfo;
    }
    public static void delProcessInfo(string aID, string aProjectID)
    {
        SqlDB dataReader = new SqlDB(aProjectID);
        dataReader.execSQL("DELETE processStatus WHERE ID='" + aID + "'");
    }
    public static void insertProcessInfo(string aID, int aStatus, int aActValue, int aMinValue, int aMaxValue, string aProjectID)
    {
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("ID", "string", aID);
        parameterList.addParameter("status", "int", aStatus.ToString());
        parameterList.addParameter("actValue", "int", aActValue.ToString());
        parameterList.addParameter("minValue", "int", aMinValue.ToString());
        parameterList.addParameter("maxValue", "int", aMaxValue.ToString());
        SqlDB dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("INSERT INTO processStatus (ID, status, actValue, minValue, maxValue) values (@ID, @status, @actValue, @minValue, @maxValue)", parameterList);
    }
    public static void updateProcessInfo(string aID, int aStatus, int aActValue, int aMinValue, int aMaxValue, string aProjectID)
    {
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("ID", "string", aID);
        parameterList.addParameter("status", "int", aStatus.ToString());
        parameterList.addParameter("actValue", "int", aActValue.ToString());
        parameterList.addParameter("minValue", "int", aMinValue.ToString());
        parameterList.addParameter("maxValue", "int", aMaxValue.ToString());
        SqlDB dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("update processStatus SET status=@status,actValue=@actValue,minValue=@minValue,maxValue=@maxValue WHERE ID=@ID", parameterList);
    }
}
