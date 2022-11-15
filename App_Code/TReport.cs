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
using System.IO;

/// <summary>
/// Klasse zur Verwaltung der Benutzers der Tools
/// </summary>
public class TReport
{
    public int ID;
    public string Filename;				
    public string DisplayName;	    	// Anzeigename des aktuellen Benutzers
    public int OrgID;
    public int FileType;
    public int FileLogo;

    /// <summary>
    /// Objekt erzeugen
    /// </summary>
    public TReport()
    {
    }
    /// <summary>
    /// Berechtigung aus DB lesen
    /// </summary>
    /// <param name="aID">ID der Berechtigung</param>
    /// <param name="aProjectID">ID des Projektes</param>
    public TReport(int aID, string aProjectID)
    {
        ID = aID;
        SqlDB dataReader;
        dataReader = new SqlDB("select filename, displayName, orgID, fileType, fileLogo from download_reports where ID='" + aID + "'", aProjectID);
        if (dataReader.read())
        {
            Filename = dataReader.getString(0);
            DisplayName = dataReader.getString(1);
            OrgID = dataReader.getInt32(2);
            FileType = dataReader.getInt32(3);
            FileLogo = dataReader.getInt32(4);
        }
        dataReader.close();
    }
    /// <summary>
    /// neue Berechtigung erzeugen
    /// </summary>
    /// <param name="aUserID">ID des Bentuzers</param>
    /// <param name="aProjectID">ID des Projektes</param>
    public TReport(string aFilename, string aDisplayname, int aOrgID, string aProjectID, int aFileType, int aFileLogo)
    {
        Filename = aFilename;
        DisplayName = aDisplayname;
        OrgID = aOrgID;
        FileType = aFileType;
        FileLogo = aFileLogo;
    }
    /// <summary>
    /// Berechtigung in DB schreiben
    /// </summary>
    /// <param name="aProjectID"></param>
    public void save(string aProjectID)
    {
        SqlDB dataReader;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("filename", "string", Filename);
        parameterList.addParameter("displayName", "string", DisplayName);
        parameterList.addParameter("orgID", "int", OrgID.ToString());
        parameterList.addParameter("fileType", "int", FileType.ToString());
        parameterList.addParameter("fileLogo", "int", FileLogo.ToString());
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("INSERT INTO download_reports (filename, displayName, orgID, fileType, fileLogo) VALUES (@filename, @displayName, @orgID, @fileType, @fileLogo)", parameterList);
    }
    /// <summary>
    /// Berechtigung in DB aktunalisieren
    /// </summary>
    /// <param name="aProjectID">ID des Projektes</param>
    public void update(string aProjectID)
    {
        SqlDB dataReader;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("filename", "string", Filename);
        parameterList.addParameter("displayName", "string", DisplayName);
        parameterList.addParameter("orgID", "int", OrgID.ToString());
        parameterList.addParameter("fileType", "int", FileType.ToString());
        parameterList.addParameter("fileLogo", "int", FileLogo.ToString());
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("UPDATE download_reports SET displayName=@displayName, fileType=@fileType, fileLogo=@fileLogo WHERE filename=@filename & orgID=@orgID", parameterList);
    }
    /// <summary>
    /// Löschen einer Berechtigung aus der DB
    /// </summary>
    /// <param name="aUserID">ID des Benutzers</param>
    /// <param name="aProjectID">ID des Projektes</param>
    public static void deleteRight(string aID, string aProjectID)
    {
        SqlDB dataReader;
        // aus Datenbank löschen
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQL("DELETE FROM download_reports WHERE ID='" + aID + "'");
    }
    /// <summary>
    /// Löschen einer Berichtsdatei aus der Dateiablage und aller Berechtigungen darauf aus der DB
    /// </summary>
    /// <param name="aUserID">ID des Benutzers</param>
    /// <param name="aProjectID">ID des Projektes</param>
    public static void deleteFile(string aFilename, string aDatapath, string aProjectID)
    {
        SqlDB dataReader;
        // Berechtigung aus Datenbank löschen
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQL("DELETE FROM download_reports WHERE filename='" + aFilename + "'");
        // Datei löschen
        File.Delete(aDatapath + aProjectID + "\\Download\\Files\\" + aFilename);

    }
    public static string getDisplayName(string aFilename, string aProjectID)
    {
        string result = "";
        SqlDB dataReader;
        dataReader = new SqlDB("select displayName from download_reports where filename='" + aFilename + "'", aProjectID);
        if (dataReader.read())
        {
            result = dataReader.getString(0);
        }
        dataReader.close();
        return result;
    }
}

