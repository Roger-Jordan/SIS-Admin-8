using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Data.SqlClient;

/// <summary>
/// Zusammenfassungsbeschreibung für SqlDB
/// </summary>
public class SqlDB
{
    /// <summary>
    /// Connectionstring, dieser muss auf die jeweils verwendete Datenbank angepasst werden
    /// </summary>
    string sConnection = ConfigurationManager.ConnectionStrings["SISDB"].ConnectionString;
    private SqlConnection sqlConnection;
    private SqlDataReader dataReader;
    private string project;
    private string sqlCommandString;

    //=================================
    /// <summary>
    /// Konstruktor für Queries
    /// </summary>
    /// <param name="projekt"></param>
    public SqlDB(string aProject)
    {
        this.project = aProject;
        sConnection = sConnection.Replace("[ProjectID]",aProject);
    }

    //=================================
    /// <summary>
    /// Konstruktor für Queries, die Abfrage wird ausgeführt, damit wird der Reader gefüllt
    /// </summary>
    /// <param name="projekt">Name des Projektes</param>
    /// <param name="q">query string</param>
    public SqlDB(String aSqlSelectCommand, string aProject)
    {
        this.project = aProject;
        sConnection = sConnection.Replace("[ProjectID]",aProject);
        sqlCommandString = aSqlSelectCommand;

        if ((sqlConnection == null) || (sqlConnection.State != ConnectionState.Open))
        {
            openConnection();
        }
        SqlCommand sqlCommand = new SqlCommand(aSqlSelectCommand, sqlConnection);
        try
        {
            sqlCommand.Prepare();
            dataReader = sqlCommand.ExecuteReader();
        }
        catch
        {
//            schreibeFehler(project, "Fehler beim Reader " + aSqlSelectCommand, ex);
        }
    }
    public SqlDB(String aSqlSelectCommand, TParameterList aParameterList, string aProject)
    {
        this.project = aProject;
        sConnection = sConnection.Replace("[ProjectID]", aProject);
        sqlCommandString = aSqlSelectCommand;

        if ((sqlConnection == null) || (sqlConnection.State != ConnectionState.Open))
        {
            openConnection();
        }
        SqlCommand sqlCommand = new SqlCommand(aSqlSelectCommand, sqlConnection);
        for (int n = 0; n < aParameterList.parameter.Count; n++)
        {
            string name = ((TParameterList.TEntry)aParameterList.parameter[n]).name;
            string value = ((TParameterList.TEntry)aParameterList.parameter[n]).value;
            if (value == null)
                value = "";
            string typ = ((TParameterList.TEntry)aParameterList.parameter[n]).type;
            SqlParameter tempParameter = new SqlParameter(name, value);
            if (typ == "string")
            {
                tempParameter.SqlDbType = System.Data.SqlDbType.NVarChar;
                tempParameter.Size = value.Length;
                if (tempParameter.Size == 0)
                {
                    tempParameter.Size = 1;
                }
            }
            if (typ == "int")
            {
                tempParameter.SqlDbType = System.Data.SqlDbType.Int;
            }
            if (typ == "float")
            {
                tempParameter.SqlDbType = System.Data.SqlDbType.Float;
            }
            if (typ == "datetime")
            {
                tempParameter.SqlDbType = System.Data.SqlDbType.DateTime;
            }
            if (typ == "text")
            {
                tempParameter.SqlDbType = System.Data.SqlDbType.NText;
                tempParameter.Size = value.Length;
                if (tempParameter.Size == 0)
                {
                    tempParameter.Size = 1;
                }
            }
            sqlCommand.Parameters.Add(tempParameter);
        }

        try
        {
            sqlCommand.Prepare();
            dataReader = sqlCommand.ExecuteReader();
        }
        catch
        {
            //            schreibeFehler(project, "Fehler beim Reader " + aSqlSelectCommand, ex);
        }
    }

    //=================================
    /// <summary>
    /// öffnet eine Connection zur Datenbank
    /// </summary>
    private void openConnection()
    {
        try
        {
            sqlConnection = new System.Data.SqlClient.SqlConnection(sConnection);
            sqlConnection.Open();
        }
        catch (Exception ex)
        {
            //Fehlerbehandlung
            schreibeFehler(project, "Fehler beim Aufbau der Connection", ex);
        }
    }

    //=================================
    /// <summary>
    /// schließt die Connection zur Datenbank
    /// </summary>
    public void closeConnection() 
    {
        if (sqlConnection!=null)
        {
            sqlConnection.Close();
            sqlConnection.Dispose();
            sqlConnection = null;
        }
    }

    //=================================
    /// <summary>
    /// setzt eine Query, die die Datenbank ändert (z.B. insert, update, delete, ...)
    /// </summary>
    /// <param name="q"></param>
    public int execSQL(String aSqlCommand)
    {
        sqlCommandString = aSqlCommand;
        if ((sqlConnection==null) || (sqlConnection.State!=ConnectionState.Open))
        {
            openConnection();
        }
        SqlCommand sqlCommand = new SqlCommand(aSqlCommand,sqlConnection);
        int rowsAffected = 0;
        try
        {
            rowsAffected = sqlCommand.ExecuteNonQuery();
        }
//		catch (Exception ex)
//		{
//            schreibeFehler(project,"Fehler bei execSQL", ex);
//        }
        finally
        {
                closeConnection();
        }
        return rowsAffected;
    }

    //=================================
    /// <summary>
    /// setzt eine Insert query, fragt dann die last_insert_id ab. Die Connection bleibt dabei offen...
    /// so findet man die zuletzt in die DB eingetragene ID
    /// </summary>
    /// <param name="q"></param>
    /// <returns></returns>
    public int execSQLAndReadID(String aSqlCommand)
    {
        sqlCommandString = aSqlCommand;
        int result;
        if ((sqlConnection==null) || (sqlConnection.State!=ConnectionState.Open))
        {
            openConnection();
        }

        System.Data.SqlClient.SqlCommand sqlCommand1 = new System.Data.SqlClient.SqlCommand(aSqlCommand,sqlConnection);			
        int rowsAffected = 0;
        try 
        {
            rowsAffected = sqlCommand1.ExecuteNonQuery();

            System.Data.SqlClient.SqlCommand sqlCommand2 = new System.Data.SqlClient.SqlCommand("select @@IDENTITY as id", sqlConnection);
            try
            {
                dataReader = sqlCommand2.ExecuteReader();
                dataReader.Read();
                result = (int)dataReader.GetDecimal(0);
                dataReader.Close();
            }
            catch (Exception ex)
            {
                schreibeFehler(project, "Fehler bei last_insert_id " + sqlCommandString, ex);
                return 0;
            }
            finally
            {
                closeConnection();
            }
            return result;
        }
        catch (Exception ex)
        {
            schreibeFehler(project, "Fehler beim Insert: " + sqlCommandString, ex);
            return 0;
        }
        finally 
        {
            closeConnection();
        }
    }
    
    /// <summary>
    /// schließt den Reader, dann sie Connection
    /// </summary>
    public void close()
    {
        if (dataReader!=null) 
        {
            dataReader.Close();
        }
        closeConnection();
    }

    /// <summary>
    /// liest einen Datensatz aus dem Reader
    /// </summary>
    /// <returns></returns>
    public bool read()
    {
        if (dataReader!=null)
        {
            return dataReader.Read();
        }
        else
        {
            return false;
        }
    }

    /// <summary>
    /// prüft, ob der reader geschlossen ist
    /// </summary>
    /// <returns></returns>
    public bool isClosed()
    {
        if (dataReader!=null)
        {
            return dataReader.IsClosed;
        }
        else
        {
            return false;
        }
    }

    /// <summary>
    /// liefert, ob eine spallte NULL ist
    /// </summary>
    /// <param name="i"></param>
    /// <returns></returns>
    private bool isNull(int i) 
    {
        return dataReader.IsDBNull(i);
    }
    
    /// <summary>
    /// holte einen int-wert aus dem reader
    /// </summary>
    /// <param name="i"></param>
    /// <returns></returns>
    public int getInt32(int i)
    {
        int value;
        try
        {
            value = dataReader.GetInt32(i);
        }
            catch (Exception ex)
            {
                schreibeFehler(project, "Fehler bei getInt " + sqlCommandString + " Nummer:" + i, ex);
                value = 0;
            }
        return value;
    }

    /// <summary>
    /// holt einen datetime-wert aus dem reader
    /// </summary>
    /// <param name="i"></param>
    /// <returns></returns>
    public DateTime getDateTime(int i) 
    {
        DateTime value;
        try 
        {
            value=dataReader.GetDateTime(i);
        }
        catch (Exception ex)
        {
            schreibeFehler(project, "Fehler bei getDateTeim " + sqlCommandString + " Nummer:" + i, ex);
            value = DateTime.MinValue;
        }
        return value;
    }

    /// <summary>
    /// holt einen string aus dem reader
    /// </summary>
    /// <param name="i"></param>
    /// <returns></returns>
    public string getString(int i)
    {
        string value;
        try
        {
            value = dataReader.GetString(i);
        }
        catch (Exception ex)
        {
            schreibeFehler(project, "Fehler bei getString " + sqlCommandString + " Nummer:" + i, ex);
            value = "";
        }
        return value;
    }

    /// <summary>
    /// holt einen bool aus dem reader
    /// </summary>
    /// <param name="i"></param>
    /// <returns></returns>
    public bool getBool(int i)
    {
        bool value;
        try
        {
            value = dataReader.GetInt32(i) == 1;
        }
        catch (Exception ex)
        {
            schreibeFehler(project, "Fehler bei getBool " + sqlCommandString + " Nummer:" + i, ex);
            value = false;
        }
        return value;
    }

    /// <summary>
    /// holt einen einzelnen String aus der DB
    /// </summary>
    /// <param name="q">SQL Query</param>
    /// <returns>ergebnis der abfrage</returns>
    public TSkalar getScalarString(string aSqlCommandString) 
    {
        object o;
        TSkalar skalar = new TSkalar();
        if ((sqlConnection==null) || (sqlConnection.State!=ConnectionState.Open))
        {
            openConnection();
        }
        System.Data.SqlClient.SqlCommand sqlCommand = new System.Data.SqlClient.SqlCommand(aSqlCommandString, sqlConnection);			
        try 
        {
            o = sqlCommand.ExecuteScalar();
            if (o.Equals(null)) 
            {
                skalar.valid=false;	
            }
            else 
            {
                skalar.valid=true;
                skalar.stringValue=(string)o;
            }
        }
        catch (Exception ex)
        {
            schreibeFehler(project, "Fehler bei getSkalarString " + aSqlCommandString, ex);
            skalar.valid = false;
        }
        finally 
        {
            closeConnection();
        }
        return skalar;
    }

    /// <summary>
    /// holt einen einzelnen integer aus der DB
    /// </summary>
    /// <param name="q">SQL Query</param>
    /// <returns>ergebnis der abfrage</returns>
    public TSkalar getScalarInt32(string aSqlCommandString) 
    {
        object o;
        TSkalar skalar = new TSkalar();
        if ((sqlConnection==null) || (sqlConnection.State!=ConnectionState.Open))
        {
            openConnection();
        }
        System.Data.SqlClient.SqlCommand sqlCommand = new System.Data.SqlClient.SqlCommand(aSqlCommandString, sqlConnection);			
        try 
        {
            o = sqlCommand.ExecuteScalar();
            if (o==null) 
            {
                skalar.valid=false;	
            }
            else 
            {
                skalar.valid=true;
                skalar.intValue=(int)o;
            }
        }
        catch (Exception ex)
        {
            schreibeFehler(project, "Fehler bei getSkalarInt32 " + aSqlCommandString, ex);
            skalar.valid=false;
        }
        finally 
        {
            closeConnection();
        }
        return skalar;
    }			

    public static void schreibeFehler(string aProject, string message, Exception e)
    {
        SqlDB dat = new SqlDB(aProject);
        string q = "INSERT INTO LOGS (message,ex) VALUES ('" + message.Replace("'", "") + "','" + e.Message.Replace("'", "") + "')";
        dat.execSQL(q);
    }
    public static void schreibeLog(string aProject, string message)
    {
        SqlDB dat = new SqlDB(aProject);
        string q = "INSERT INTO LOGS (message) VALUES ('" + message.Replace("'", "") + "')";
        dat.execSQL(q);
    }

    /// <summary>
    /// führt ein parametrisiertes Kommando aus
    /// </summary>
    /// <param name="q"></param>
    /// <param name="paraNamen"></param>
    /// <param name="paraTypen"></param>
    /// <param name="paraWerte"></param>
    public int execSQLwithParameter(string aSqlCommand, TParameterList aParameterList)
    {
        //int erg = 0;
        sqlCommandString = aSqlCommand;
        if ((sqlConnection == null) || (sqlConnection.State != ConnectionState.Open))
        {
            openConnection();
        }
        SqlCommand sqlCommand = new SqlCommand(aSqlCommand, sqlConnection);
        for (int n = 0; n < aParameterList.parameter.Count; n++)
        {
            string name = ((TParameterList.TEntry)aParameterList.parameter[n]).name;
            string value = ((TParameterList.TEntry)aParameterList.parameter[n]).value;
            string typ = ((TParameterList.TEntry)aParameterList.parameter[n]).type;

            SqlParameter tempParameter = new SqlParameter(name, value);
            if (typ == "string")
            {
                tempParameter.SqlDbType = System.Data.SqlDbType.NVarChar;
                tempParameter.Size = value.Length;
                if (tempParameter.Size == 0)
                {
                    tempParameter.Size = 1;
                }
            }
            if (typ == "int")
            {
                tempParameter.SqlDbType = System.Data.SqlDbType.Int;
            }
            if (typ == "float")
            {
                tempParameter.SqlDbType = System.Data.SqlDbType.Float;
            }
            if (typ == "datetime")
            {
                tempParameter.SqlDbType = System.Data.SqlDbType.DateTime;
            }
            if (typ == "text")
            {
                tempParameter.SqlDbType = System.Data.SqlDbType.NText;
                tempParameter.Size = value.Length;
                if (tempParameter.Size == 0)
                {
                    tempParameter.Size = 1;
                }
            }
            sqlCommand.Parameters.Add(tempParameter);
        }

        sqlCommand.Prepare();
        int rowsAffected = 0;
        try
        {
            rowsAffected = sqlCommand.ExecuteNonQuery();
        }
        //catch (Exception ex)
        //{
        //    schreibeFehler(projekt, "Fehler bei der Abfrage " + aSqlCommand, ex);
        //    closeConnection();
        //    return 0;
        //}
        finally
        {
            closeConnection();
        }
        return rowsAffected;
    }
    public int execSQLwithParameterAndReadID(string aSqlCommand, TParameterList aParameterList)
    {
        sqlCommandString = aSqlCommand;
        if ((sqlConnection == null) || (sqlConnection.State != ConnectionState.Open))
        {
            openConnection();
        }
        SqlCommand sqlCommand = new SqlCommand(aSqlCommand, sqlConnection);
        for (int n = 0; n < aParameterList.parameter.Count; n++)
        {
            string name = ((TParameterList.TEntry)aParameterList.parameter[n]).name;
            string value = ((TParameterList.TEntry)aParameterList.parameter[n]).value;
            string typ = ((TParameterList.TEntry)aParameterList.parameter[n]).type;

            SqlParameter tempParameter = new SqlParameter(name, value);
            if (typ == "string")
            {
                tempParameter.SqlDbType = System.Data.SqlDbType.NVarChar;
                tempParameter.Size = value.Length;
                if (tempParameter.Size == 0)
                {
                    tempParameter.Size = 1;
                }
            }
            if (typ == "int")
            {
                tempParameter.SqlDbType = System.Data.SqlDbType.Int;
            }
            if (typ == "float")
            {
                tempParameter.SqlDbType = System.Data.SqlDbType.Float;
            }
            if (typ == "datetime")
            {
                tempParameter.SqlDbType = System.Data.SqlDbType.DateTime;
            }
            if (typ == "text")
            {
                tempParameter.SqlDbType = System.Data.SqlDbType.NText;
                tempParameter.Size = value.Length;
                if (tempParameter.Size == 0)
                {
                    tempParameter.Size = 1;
                }
            }
            sqlCommand.Parameters.Add(tempParameter);
        }
        sqlCommand.Prepare();
        int rowsAffected = 0;
        try
        {
            rowsAffected = sqlCommand.ExecuteNonQuery();

            int result;
            SqlCommand sqlCommand2 = new SqlCommand("select @@IDENTITY as id", sqlConnection);
            try
            {
                dataReader = sqlCommand2.ExecuteReader();
                dataReader.Read();
                result = (int)dataReader.GetDecimal(0);
                dataReader.Close();
            }
            catch (Exception ex)
            {
                schreibeFehler(project, "Fehler bei last_insert_id " + sqlCommandString, ex);
                return 0;
            }
            finally
            {
                closeConnection();
            }
            return result;
        }
        //catch (Exception ex)
        //{
        //    schreibeFehler(projekt, "Fehler bei der Abfrage " + aSqlCommand, ex);
        //    closeConnection();
        //    return 0;
        //}
        finally
        {
            closeConnection();
        }
    }

    ///// <summary>
    ///// liefert für die übergebene Query einen DataAdapter zurück
    ///// </summary>
    ///// <param name="q"></param>
    ///// <returns></returns>
    //public System.Data.Common.DataAdapter getAdapter(string q, bool connectionSchliessen)
    //{
    //    qString = q;
    //    if ((objConn == null) || (objConn.State != ConnectionState.Open))
    //    {
    //        openConnection();
    //    }
    //    System.Data.SqlClient.SqlDataAdapter da;
    //    try
    //    {
    //        da = new System.Data.SqlClient.SqlDataAdapter(q, objConn);
    //        System.Data.SqlClient.SqlCommandBuilder cmdBuiler = new System.Data.SqlClient.SqlCommandBuilder(da);
    //        string s = cmdBuiler.GetUpdateCommand().CommandText;
    //    }
    //    catch (Exception ex)
    //    {
    //        schreibeFehler(projekt, "Fehler beim Anfordern des Adapters " + q, ex);
    //        da = null;
    //    }
    //    if (connectionSchliessen)
    //    {
    //        closeConnection();
    //    }
    //    return da;
    //}
}
