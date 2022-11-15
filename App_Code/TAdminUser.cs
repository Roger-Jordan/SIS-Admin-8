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
/// Klasse für Daten des aktuell angemeldetet Users/Administrators
/// </summary>

public class TAdminUser
{
    public string UserID;				// eindeutige Kennung des aktuellen Benutzers
    public string DisplayName;		// Anzeigename des aktuellen Benutzers
    public string Password;			// Kennwort des aktuellen Benutzers
    public string Email;
    public DateTime LastLogin;			// letzter erfolgreicher Login des aktuellen Benutzers
    public DateTime LastLoginError;		// letzter nicht erfolgreicher Login des aktuellen Benutzers
    public int LastLoginErrorCount;		// Anzahl aktueller nicht erfolgreicher Login des aktuellen Benutzers
    public bool ForceChange;			// Benutzer muss bei Anmeldung Kennwort ändern
    public bool Active;					// Aktueller Benutzer ist aktiviert
    public bool Exists;
    public bool SuperAdmin;             // User darf Admins anlegen, bearbeiten, löschen, Projekte anlegen, archivieren, löschen
    public ArrayList projectRights;
    
    /// <summary>
    /// leeres TUser-Objekt erzeugen
    /// </summary>
    public TAdminUser()
    {
        projectRights = new ArrayList();
    }
    /// <summary>
    /// User mit allen Eigenschaften aus DB laden
    /// </summary>
    /// <param name="aUserID">ID des Benutzers</param>
    public TAdminUser(string aUserID)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        UserID = aUserID;
        Exists = false;

        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("userID", "string", aUserID);
        SqlDB dataReader;
        dataReader = new SqlDB("SELECT userID, displayName, password, email, superAdmin, lastLogin, lastLoginError, lastLoginErrorCount, forceChange, active FROM admins where userID=@userID", parameterList, "");
        if (dataReader.read())
        {
            Exists = true;
            DisplayName = dataReader.getString(1);
            Password = dataReader.getString(2);
            Email = dataReader.getString(3);
            SuperAdmin = dataReader.getBool(4);
            LastLogin = dataReader.getDateTime(5);
            LastLoginError = dataReader.getDateTime(6);
            LastLoginErrorCount = dataReader.getInt32(7);
            ForceChange = dataReader.getBool(8);
            Active = dataReader.getBool(9);
        }
        dataReader.close();
        projectRights = new ArrayList();
        dataReader = new SqlDB("SELECT projectID FROM mapadminprojects where userID=@userID", parameterList, "");
        while (dataReader.read())
        {
            // Version prüfen
            TParameterList parameterList1 = new TParameterList();
            parameterList1.addParameter("projectID", "string", dataReader.getString(0));
            SqlDB dataReader1 = new SqlDB("SELECT version FROM projects where projectID=@projectID", parameterList1, "");
            if (dataReader1.read())
            {
                if (dataReader1.getInt32(0)==SessionObj.AppVersion)
                    projectRights.Add(dataReader.getString(0));
            }
            dataReader1.close();
        }
        dataReader.close();
    }
    /// <summary>
    /// Erzeugen eines Kennwortes für denBenutzer; Läge 10 Zeichen; numerisch aus den Ziffern 1 bis9
    /// </summary>
    /// <returns>erzeugtes Kennwort</returns>
    private string generatePassword()
    {
        string result = "";
        Random randomObject = new Random();
        for (int Index = 0; Index < 10; Index++)
        {
            result = result + randomObject.Next(1, 9).ToString();
        }
        return result;
    }
    /// <summary>
    /// Versendung eines Emails mit Link und Zugangsdaten an den Benutzer
    /// </summary>
    /// <param name="aServer">IP Adresse des zu verwendenden Mail Servers</param>
    public void sendMail(string aServer, string aUser, string aPassword, string aMailSenderAdress, string aMailSenderName)
    {
        System.Net.Mail.MailAddress From = new System.Net.Mail.MailAddress(aMailSenderAdress, aMailSenderName);
        System.Net.Mail.MailAddress To = new System.Net.Mail.MailAddress(Email);
        System.Net.Mail.MailMessage Message = new System.Net.Mail.MailMessage(From, To);

        Message.Subject = "";
        Message.BodyEncoding = System.Text.Encoding.UTF8;
        Message.Body = Message.Body + (char)10 + (char)13;
        Message.Body = Password;
        Message.Body = Message.Body + (char)10 + (char)13;
        System.Net.NetworkCredential mailClientCredentials = new System.Net.NetworkCredential(aUser, aPassword);
        System.Net.Mail.SmtpClient Client = new System.Net.Mail.SmtpClient(aServer);
        Client.Credentials = mailClientCredentials;
        Client.Send(Message);
    }
    /// <summary>
    /// Speichern eines neuen Admin
    /// </summary>
    public bool save(int aAppVersion)
    {
        // userID auf eindeutigkeit prüfen
        bool exists = false;
        SqlDB dataReader = new SqlDB("SELECT userID FROM admins WHERE userID='" + this.UserID + "'", "");
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            Password = generatePassword();
            string tempForceChange = "0";
            if (ForceChange)
                tempForceChange = "1";
            string tempSuperAdmin = "0";
            if (SuperAdmin)
                tempSuperAdmin = "1";
            string tempActive = "0";
            if (Active)
                tempActive = "1";
            TParameterList parameterList = new TParameterList();
            parameterList.addParameter("userID", "string", UserID);
            parameterList.addParameter("displayName", "string", DisplayName);
            parameterList.addParameter("password", "string", Password);
            parameterList.addParameter("email", "string", Email);
            parameterList.addParameter("forceChange", "int", tempForceChange);
            parameterList.addParameter("superAdmin", "int", tempSuperAdmin);
            parameterList.addParameter("active", "int", tempActive);
            parameterList.addParameter("lastLogin", "datetime", DateTime.Now.ToString());
            parameterList.addParameter("lastLoginError", "datetime", DateTime.Now.ToString());
            parameterList.addParameter("lastLoginErrorCount", "int", LastLoginErrorCount.ToString());
            dataReader = new SqlDB("");
            dataReader.execSQLwithParameter("INSERT INTO admins (userID, displayname, password, email, superAdmin, lastLogin, lastLoginError, lastLoginErrorCount, forceChange, active) VALUES (@userID, @displayName, @password, @email, @superAdmin, @lastLogin, @lastLoginError, @lastLoginErrorCount, @forceChange, @active)", parameterList);
            // projectRights
            foreach (string tempValue in projectRights)
            {
                parameterList = new TParameterList();
                parameterList.addParameter("userID", "string", UserID);
                parameterList.addParameter("projectID", "string", tempValue);
                dataReader = new SqlDB("");
                dataReader.execSQLwithParameter("INSERT INTO mapadminprojects (userID, projectID) VALUES (@userID, @projectID)", parameterList);
            }
            return true;
        }
        else
        {
            return false;
        }
    }
    public void update(int aAppVersion)
    {
        string tempForceChange = "0";
        if (ForceChange)
            tempForceChange = "1";
        string tempSuperAdmin = "0";
        if (SuperAdmin)
            tempSuperAdmin = "1";
        string tempActive = "0";
        if (Active)
            tempActive = "1";
        SqlDB dataReader;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("displayName", "string", DisplayName);
        parameterList.addParameter("password", "string", Password);
        parameterList.addParameter("email", "string", Email);
        parameterList.addParameter("forceChange", "int", tempForceChange);
        parameterList.addParameter("superAdmin", "int", tempSuperAdmin);
        parameterList.addParameter("active", "int", tempActive);
        parameterList.addParameter("lastLoginErrorCount", "int", LastLoginErrorCount.ToString());
        parameterList.addParameter("userID", "string", UserID);
        dataReader = new SqlDB("");
        dataReader.execSQLwithParameter("UPDATE admins SET displayName=@displayName, password=@password, email=@email, forceChange=@forceChange, superAdmin=@superAdmin, active=@active, lastLoginErrorCount=@lastLoginErrorCount WHERE userID=@userID", parameterList);
        // projectRights löschen
        dataReader = new SqlDB("SELECT projectID FROM projects WHERE version='" + aAppVersion.ToString() + "' ORDER by projectID", "");
        while (dataReader.read())
        {
            SqlDB dataReader1 = new SqlDB("");
            dataReader1.execSQL("DELETE FROM mapadminprojects WHERE userID='" + UserID + "' AND projectID='" + dataReader.getString(0) + "'");
        }
        dataReader.close();
        // projectRights speichern
        foreach (string tempValue in projectRights)
        {
            parameterList = new TParameterList();
            parameterList.addParameter("userID", "string", UserID);
            parameterList.addParameter("projectID", "string", tempValue);
            dataReader = new SqlDB("");
            dataReader.execSQLwithParameter("INSERT INTO mapadminprojects (userID, projectID) VALUES (@userID, @projectID)", parameterList);
        }
    }
    /// <summary>
    /// Login-Eigenschaften des Benutzers in der Db aktualisieren
    /// </summary>
    public void updatePassword()
    {
        string tempForceChange = "0";
        if (ForceChange)
            tempForceChange = "1";
        SqlDB dataReader;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("password", "string", Password);
        parameterList.addParameter("forceChange", "int", tempForceChange);
        parameterList.addParameter("lastLogin", "datetime", LastLogin.ToString());
        parameterList.addParameter("lastLoginError", "datetime", LastLoginError.ToString());
        parameterList.addParameter("lastLoginErrorCount", "int", LastLoginErrorCount.ToString());
        parameterList.addParameter("userID", "string", UserID);
        dataReader = new SqlDB("");
        dataReader.execSQLwithParameter("UPDATE admins SET password=@password, forceChange=@forceChange, lastLogin=@lastLogin, lastLoginError=@lastLoginError, lastLoginErrorCount=@lastLoginErrorCount WHERE userID=@userID",parameterList);
    }
    /// <summary>
    /// Löschen eines Admins aus der DB
    /// </summary>
    /// <param name="aUserID">ID des Admins</param>
    public static void delete(string aUserID)
    {
        SqlDB dataReader;
        // aus Datenbank löschen
        dataReader = new SqlDB("");
        dataReader.execSQL("DELETE FROM admins WHERE userID='" + aUserID + "'");
        dataReader = new SqlDB("");
        dataReader.execSQL("DELETE FROM mapadminprojects WHERE userID='" + aUserID + "'");
    }
}

