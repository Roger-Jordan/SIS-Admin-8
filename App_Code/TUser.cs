using System;
using System.Collections;
using System.Collections.Specialized;
using System.IO;
using System.Text;
using System.Web;

/// <summary>
/// Klasse zur Verwaltung der Benutzers der Tools
/// </summary>
public class TUser
{
    public class TToolRight
    {
        public string toolID;
    }
    public class TOrgRight
    {
        public int ID;
        public int OrgID;
        public string OProfile = "";
        public int OReadLevel;
        public int OEditLevel;
        public bool OAllowDelegate;
        public bool OAllowLock;
        public bool OAllowEmployeeRead;
        public bool OAllowEmployeeEdit;
        public bool OAllowEmployeeImport;
        public bool OAllowEmployeeExport;
        public bool OAllowUnitAdd;
        public bool OAllowUnitMove;
        public bool OAllowUnitDelete;
        public bool OAllowUnitProperty;
        public bool OAllowReportRecipient;
        public bool OAllowStructureImport;
        public bool OAllowStructureExport;
        public bool OAllowBouncesRead;
        public bool OAllowBouncesEdit;
        public bool OAllowBouncesDelete;
        public bool OAllowBouncesExport;
        public bool OAllowBouncesImport;
        public bool OAllowReInvitation;
        public bool OAllowPostNomination;
        public bool OAllowPostNominationImport;
        public string FProfile = "";
        public int FReadLevel;
        public int FEditLevel;
        public int FSumLevel;
        public bool FAllowCommunications;
        public bool FAllowMeasures;
        public bool FAllowDelete;
        public bool FAllowExcelExport;
        public bool FAllowReminderOnMeasure;
        public bool FAllowReminderOnUnit;
        public bool FAllowUpUnitsMeasures;
        public bool FAllowColleaguesMeasures;
        public bool FAllowThemes;
        public bool FAllowReminderOnThemes;
        public bool FAllowReminderOnThemeUnit;
        public bool FAllowTopUnitsThemes;
        public bool FAllowColleaguesThemes;
        public bool FAllowOrder;
        public bool FAllowReminderOnOrderUnit;
        public string RProfile = "";
        public int RReadLevel;
        public string DProfile = "";
        public int DReadLevel;
        public bool DAllowMail;
        public bool DAllowDownload;
        public bool DAllowLog;
        public bool DAllowType1;
        public bool DAllowType2;
        public bool DAllowType3;
        public bool DAllowType4;
        public bool DAllowType5;
        public int AEditLevelWorkshop;
        public int AReadLevelStructure;
    }
    public class TBackRight
    {
        public int ID;
        public int OrgID;
        public string OProfile = "";
        public int OReadLevel;
        public int AReadLevel;
    }
    public class TOrg1Right
    {
        public int ID;
        public int OrgID;
        public string OProfile = "";
        public int OReadLevel;
        public int OEditLevel;
        public bool OAllowDelegate;
        public bool OAllowLock;
        public bool OAllowEmployeeRead;
        public bool OAllowEmployeeEdit;
        public bool OAllowEmployeeImport;
        public bool OAllowEmployeeExport;
        public bool OAllowUnitAdd;
        public bool OAllowUnitMove;
        public bool OAllowUnitDelete;
        public bool OAllowUnitProperty;
        public bool OAllowReportRecipient;
        public bool OAllowStructureImport;
        public bool OAllowStructureExport;
        public string FProfile = "";
        public int FReadLevel;
        public int FEditLevel;
        public int FSumLevel;
        public bool FAllowCommunications;
        public bool FAllowMeasures;
        public bool FAllowDelete;
        public bool FAllowExcelExport;
        public bool FAllowReminderOnMeasure;
        public bool FAllowReminderOnUnit;
        public bool FAllowUpUnitsMeasures;
        public bool FAllowColleaguesMeasures;
        public bool FAllowThemes;
        public bool FAllowReminderOnThemes;
        public bool FAllowReminderOnThemeUnit;
        public bool FAllowTopUnitsThemes;
        public bool FAllowColleaguesThemes;
        public bool FAllowOrder;
        public bool FAllowReminderOnOrderUnit;
        public string RProfile = "";
        public int RReadLevel;
        public string DProfile = "";
        public int DReadLevel;
        public bool DAllowMail;
        public bool DAllowDownload;
        public bool DAllowLog;
        public bool DAllowType1;
        public bool DAllowType2;
        public bool DAllowType3;
        public bool DAllowType4;
        public bool DAllowType5;
        public int AEditLevelWorkshop;
        public int AReadLevelStructure;
    }
    public class TBack1Right
    {
        public int ID;
        public int OrgID;
        public string OProfile = "";
        public int OReadLevel;
        public int AReadLevel;
    }
    public class TOrg2Right
    {
        public int ID;
        public int OrgID;
        public string OProfile = "";
        public int OReadLevel;
        public int OEditLevel;
        public bool OAllowLock;
        public bool OAllowDelegate;
        public bool OAllowEmployeeRead;
        public bool OAllowEmployeeEdit;
        public bool OAllowEmployeeImport;
        public bool OAllowEmployeeExport;
        public bool OAllowUnitAdd;
        public bool OAllowUnitMove;
        public bool OAllowUnitDelete;
        public bool OAllowUnitProperty;
        public bool OAllowReportRecipient;
        public bool OAllowStructureImport;
        public bool OAllowStructureExport;
        public string FProfile = "";
        public int FReadLevel;
        public int FEditLevel;
        public int FSumLevel;
        public bool FAllowCommunications;
        public bool FAllowMeasures;
        public bool FAllowDelete;
        public bool FAllowExcelExport;
        public bool FAllowReminderOnMeasure;
        public bool FAllowReminderOnUnit;
        public bool FAllowUpUnitsMeasures;
        public bool FAllowColleaguesMeasures;
        public bool FAllowThemes;
        public bool FAllowReminderOnThemes;
        public bool FAllowReminderOnThemeUnit;
        public bool FAllowTopUnitsThemes;
        public bool FAllowColleaguesThemes;
        public bool FAllowOrder;
        public bool FAllowReminderOnOrderUnit;
        public string RProfile = "";
        public int RReadLevel;
        public string DProfile = "";
        public int DReadLevel;
        public bool DAllowMail;
        public bool DAllowDownload;
        public bool DAllowLog;
        public bool DAllowType1;
        public bool DAllowType2;
        public bool DAllowType3;
        public bool DAllowType4;
        public bool DAllowType5;
        public int AEditLevelWorkshop;
        public int AReadLevelStructure;
    }
    public class TBack2Right
    {
        public int ID;
        public int OrgID;
        public string OProfile = "";
        public int OReadLevel;
        public int AReadLevel;
    }
    public class TOrg3Right
    {
        public int ID;
        public int OrgID;
        public string OProfile = "";
        public int OReadLevel;
        public int OEditLevel;
        public bool OAllowLock;
        public bool OAllowDelegate;
        public bool OAllowEmployeeRead;
        public bool OAllowEmployeeEdit;
        public bool OAllowEmployeeImport;
        public bool OAllowEmployeeExport;
        public bool OAllowUnitAdd;
        public bool OAllowUnitMove;
        public bool OAllowUnitDelete;
        public bool OAllowUnitProperty;
        public bool OAllowReportRecipient;
        public bool OAllowStructureImport;
        public bool OAllowStructureExport;
        public string FProfile = "";
        public int FReadLevel;
        public int FEditLevel;
        public int FSumLevel;
        public bool FAllowCommunications;
        public bool FAllowMeasures;
        public bool FAllowDelete;
        public bool FAllowExcelExport;
        public bool FAllowReminderOnMeasure;
        public bool FAllowReminderOnUnit;
        public bool FAllowUpUnitsMeasures;
        public bool FAllowColleaguesMeasures;
        public bool FAllowThemes;
        public bool FAllowReminderOnThemes;
        public bool FAllowReminderOnThemeUnit;
        public bool FAllowTopUnitsThemes;
        public bool FAllowColleaguesThemes;
        public bool FAllowOrder;
        public bool FAllowReminderOnOrderUnit;
        public string RProfile = "";
        public int RReadLevel;
        public string DProfile = "";
        public int DReadLevel;
        public bool DAllowMail;
        public bool DAllowDownload;
        public bool DAllowLog;
        public bool DAllowType1;
        public bool DAllowType2;
        public bool DAllowType3;
        public bool DAllowType4;
        public bool DAllowType5;
        public int AEditLevelWorkshop;
        public int AReadLevelStructure;
    }
    public class TBack3Right
    {
        public int ID;
        public int OrgID;
        public string OProfile = "";
        public int OReadLevel;
        public int AReadLevel;
    }
    public class TUserRight
    {
        public int ID;
        public string TTaskProfile = "";
        public bool TAllowTaskRead;
        public bool TAllowTaskCreateEditDelete;
        public bool TAllowTaskFileUpload;
        public bool TAllowTaskExport;
        public bool TAllowTaskImport;
        public string TContactProfile = "";
        public bool TAllowContactRead;
        public bool TAllowContactCreateEditDelete;
        public bool TAllowContactExport;
        public bool TAllowContactImport;
        public string TTranslateProfile = "";
        public string TTranslatefile = "";
        public bool TAllowTranslateRead;
        public bool TAllowTranslateEdit;
        public bool TAllowTranslateRelease;
        public bool TAllowTranslateNew;
        public bool TAllowTranslateDelete;
        public bool AAllowAccountTransfer;
        public bool AAllowAccountManagement;
    }
    public class TContainerRight
    {
        public int ID;
        public int ContainerID;
        public string TContainerProfile = "";
        public bool TAllowUpload;
        public bool TAllowDownload;
        public bool TAllowDeleteOwnFiles;
        public bool TAllowDeleteAllFiles;
        public bool TAllowAccessOwnFilesWithoutPassword;
        public bool TAllowAccessAllFilesWithoutPassword;
        public bool TAllowCreateFolder;
        public bool TAllowDeleteOwnFolder;
        public bool TAllowDeleteAllFolder;
        public bool TAllowResetPassword;
        public bool TAllowTakeOwnership;
    }
    public string UserID;				// eindeutige Kennung des aktuellen Benutzers
    public string DisplayName;	    	// Anzeigename des aktuellen Benutzers
    public string Password;		    	// Kennwort des aktuellen Benutzers
    public string Email;		    	// Email-Adresse der aktuellen Benutzers
    public string Title;
    public string Name;
    public string Firstname;
    public string Street;
    public string Zipcode;
    public string City;
    public string State;
    public DateTime LastLogin;			// letzter erfolgreicher Login des aktuellen Benutzers
    public DateTime LastLoginError;		// letzter nicht erfolgreicher Login des aktuellen Benutzers
    public int LastLoginErrorCount;		// Anzahl aktueller nicht erfolgreicher Login des aktuellen Benutzers
    public bool ForceChange;			// Benutzer muss bei Anmeldung Kennwort ändern
    public bool Active;					// Aktueller Benutzer ist aktiviert
    public bool SystemGeneratedPassword;// Aktuelles Passwort ist vom System generiert
    public DateTime PasswordValidDate;			// Datem der Gültigkeit des Passsworts
    public int Language;                // Sprache des Benutzers
    public string Comment;                 // Kommentar des Benutzers
    public bool IsDelegate;               // Kennzeichen, ob Benutzer durch Delegieren im OrgManager erzeugt worden ist
    public bool IsInternalUser;         // Kennzeichen, ob es ein Benutzer ist, die keine Lizenz benötigt
    public string GroupID;                 // Benutzergruppe des Benutzers
    public bool Exists;
    // Profilberechtigungen
    public ArrayList profileRights;
    // Einheiten-Berechtigungen
    public ArrayList orgRights;
    public string orgRightsSelectedValue;
    public int maxOrgRightID;
    public ArrayList backRights;
    public string backRightsSelectedValue;
    public int maxBackRightID;
    public ArrayList org1Rights;
    public string org1RightsSelectedValue;
    public int maxOrg1RightID;
    public ArrayList back1Rights;
    public string back1RightsSelectedValue;
    public int maxBack1RightID;
    public ArrayList org2Rights;
    public string org2RightsSelectedValue;
    public int maxOrg2RightID;
    public ArrayList back2Rights;
    public string back2RightsSelectedValue;
    public int maxBack2RightID;
    public ArrayList org3Rights;
    public string org3RightsSelectedValue;
    public int maxOrg3RightID;
    public ArrayList back3Rights;
    public string back3RightsSelectedValue;
    public int maxBack3RightID;
    // Tool Berechtigungen
    public ArrayList toolRights;
    // Container Berechtigungen
    public ArrayList containerRights;
    public string containerRightsSelectedValue;
    public int maxContainerRightID;
    // TaskContactTranslate Berechtigungen
    public TUserRight userRights;
    // Taskroles Berechtigungen
    public ArrayList taskroleRights;
    // Gruppen Berechtigungen
    public ArrayList usergroupRights;
    // TranslateLanguages Berechtigungen
    public ArrayList translateLanguagesRights;

    public TUser()
    {
        UserID = "";
        DisplayName = "";
        Password = "";
        Email = "";
        Title = "";
        Name = "";
    Firstname = "";
    Street = "";
    Zipcode = "";
    City = "";
        State = ""; ;
    ForceChange = false;			// Benutzer muss bei Anmeldung Kennwort ändern
    Active = false;					// Aktueller Benutzer ist aktiviert
    Language = 0;                // Sprache des Benutzers
    Comment = "";                 // Kommentar des Benutzers
    IsDelegate = false;               // Kennzeichen, ob Benutzer durch Delegieren im OrgManager erzeugt worden ist
    GroupID = "";                 // Benutzergruppe des Benutzers
    Exists = false;


    orgRights = new ArrayList();
        backRights = new ArrayList();
        org1Rights = new ArrayList();
        back1Rights = new ArrayList();
        org2Rights = new ArrayList();
        back2Rights = new ArrayList();
        org3Rights = new ArrayList();
        back3Rights = new ArrayList();
        toolRights = new ArrayList();
        containerRights = new ArrayList();
        userRights = new TUserRight();
        taskroleRights = new ArrayList();
        usergroupRights = new ArrayList();
        translateLanguagesRights = new ArrayList();
    }
    /// <summary>
    /// Eigenschaften eines Benutzer aus der DB lesen
    /// </summary>
    /// <param name="aUserID">ID des Bentuzers</param>
    /// <param name="aProjectID">ID des Projektes</param>
    public TUser(string aUserID, string aProjectID)
    {
        UserID = "";
        DisplayName = "";
        Password = "";
        Email = "";
        Title = "";
        Name = "";
        Firstname = "";
        Street = "";
        Zipcode = "";
        City = "";
        State = ""; ;
        ForceChange = false;            // Benutzer muss bei Anmeldung Kennwort ändern
        Active = false;                 // Aktueller Benutzer ist aktiviert
        Language = 0;                // Sprache des Benutzers
        Comment = "";                 // Kommentar des Benutzers
        IsDelegate = false;               // Kennzeichen, ob Benutzer durch Delegieren im OrgManager erzeugt worden ist
        GroupID = "";                 // Benutzergruppe des Benutzers
        Exists = false;

        SqlDB dataReader;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("userID", "string", aUserID);
        dataReader = new SqlDB("select userID, displayName, password, email, title, name, firstname, street, zipcode, city, state, lastLogin, lastLoginError, lastLoginErrorCount, isinternaluser, forceChange, active, systemgeneratedpassword, passwordValidDate, language, comment, groupID, isdelegate from loginuser where userID=@userID", parameterList, aProjectID);
        if (dataReader.read())
        {
            TCrypt cryptKey = new TCrypt();
            Exists = true;
            UserID = dataReader.getString(0);
            DisplayName = dataReader.getString(1);
            Password = dataReader.getString(2);
            Email = dataReader.getString(3);
            Title = dataReader.getString(4);
            Name = dataReader.getString(5);
            Firstname = dataReader.getString(6);
            Street = dataReader.getString(7);
            Zipcode = dataReader.getString(8);
            City = dataReader.getString(9);
            State = dataReader.getString(10);
            LastLogin = dataReader.getDateTime(11);
            LastLoginError = dataReader.getDateTime(12);
            LastLoginErrorCount = dataReader.getInt32(13);
            IsInternalUser = dataReader.getBool(14);
            ForceChange = dataReader.getBool(15);
            Active = dataReader.getBool(16);
            SystemGeneratedPassword = dataReader.getBool(17);
            PasswordValidDate = dataReader.getDateTime(18);
            Language = dataReader.getInt32(19);
            Comment = dataReader.getString(20);
            GroupID = dataReader.getString(21);
            IsDelegate = dataReader.getBool(22);
        }
        dataReader.close();

        maxOrgRightID = 0;
        maxBackRightID = 0;
        maxOrg1RightID = 0;
        maxBack1RightID = 0;
        maxOrg2RightID = 0;
        maxBack2RightID = 0;
        maxContainerRightID = 0;
        // Unit-Rechte auslesen
        orgRights = new ArrayList();
        dataReader = new SqlDB("select ID, orgID from mapuserorgid where userID=@userID ORDER BY orgID", parameterList, aProjectID);
        while (dataReader.read())
        {
            int actID = dataReader.getInt32(0);
            int actOrgID = dataReader.getInt32(1);
            orgRights.Add(getOrgRight(aUserID, actOrgID, aProjectID));
            maxOrgRightID = Math.Max(maxOrgRightID, actID);
        }
        dataReader.close();
        backRights = new ArrayList();
        dataReader = new SqlDB("select ID, orgID from mapuserback where userID=@userID ORDER BY orgID", parameterList, aProjectID);
        while (dataReader.read())
        {
            int actID = dataReader.getInt32(0);
            int actOrgID = dataReader.getInt32(1);
            backRights.Add(getBackRight(aUserID, actOrgID, aProjectID));
            maxBackRightID = Math.Max(maxBackRightID, actID);
        }
        dataReader.close();
        org1Rights = new ArrayList();
        dataReader = new SqlDB("select ID, orgID from mapuserorgid1 where userID=@userID ORDER BY orgID", parameterList, aProjectID);
        while (dataReader.read())
        {
            int actID = dataReader.getInt32(0);
            int actOrgID = dataReader.getInt32(1);
            org1Rights.Add(getOrg1Right(aUserID, actOrgID, aProjectID));
            maxOrg1RightID = Math.Max(maxOrg1RightID, actID);
        }
        dataReader.close();
        back1Rights = new ArrayList();
        dataReader = new SqlDB("select ID, orgID from mapuserback1 where userID=@userID ORDER BY orgID", parameterList, aProjectID);
        while (dataReader.read())
        {
            int actID = dataReader.getInt32(0);
            int actOrgID = dataReader.getInt32(1);
            back1Rights.Add(getBack1Right(aUserID, actOrgID, aProjectID));
            maxBack1RightID = Math.Max(maxBack1RightID, actID);
        }
        dataReader.close();
        org2Rights = new ArrayList();
        dataReader = new SqlDB("select ID, orgID from mapuserorgid2 where userID=@userID ORDER BY orgID", parameterList, aProjectID);
        while (dataReader.read())
        {
            int actID = dataReader.getInt32(0);
            int actOrgID = dataReader.getInt32(1);
            org2Rights.Add(getOrg2Right(aUserID, actOrgID, aProjectID));
            maxOrg2RightID = Math.Max(maxOrg2RightID, actID);
        }
        dataReader.close();
        back2Rights = new ArrayList();
        dataReader = new SqlDB("select ID, orgID from mapuserback2 where userID=@userID ORDER BY orgID", parameterList, aProjectID);
        while (dataReader.read())
        {
            int actID = dataReader.getInt32(0);
            int actOrgID = dataReader.getInt32(1);
            back2Rights.Add(getBack2Right(aUserID, actOrgID, aProjectID));
            maxBack2RightID = Math.Max(maxBack2RightID, actID);
        }
        dataReader.close();
        org3Rights = new ArrayList();
        dataReader = new SqlDB("select ID, orgID from mapuserorgid3 where userID=@userID ORDER BY orgID", parameterList, aProjectID);
        while (dataReader.read())
        {
            int actID = dataReader.getInt32(0);
            int actOrgID = dataReader.getInt32(1);
            org3Rights.Add(getOrg3Right(aUserID, actOrgID, aProjectID));
            maxOrg3RightID = Math.Max(maxOrg3RightID, actID);
        }
        dataReader.close();
        back3Rights = new ArrayList();
        dataReader = new SqlDB("select ID, orgID from mapuserback3 where userID=@userID ORDER BY orgID", parameterList, aProjectID);
        while (dataReader.read())
        {
            int actID = dataReader.getInt32(0);
            int actOrgID = dataReader.getInt32(1);
            back3Rights.Add(getBack3Right(aUserID, actOrgID, aProjectID));
            maxBack3RightID = Math.Max(maxBack3RightID, actID);
        }
        dataReader.close();

        // Container-Rechte auslesen
        containerRights = new ArrayList();
        dataReader = new SqlDB("select ID, containerID from mapusercontainer where userID=@userID ORDER BY containerID", parameterList, aProjectID);
        while (dataReader.read())
        {
            int actID = dataReader.getInt32(0);
            int actContainerID = dataReader.getInt32(1);
            containerRights.Add(getContainerRight(aUserID, actContainerID, aProjectID));
            maxContainerRightID = Math.Max(maxContainerRightID, actContainerID);
        }
        dataReader.close();

        // Projektplan, Kontakte- und Übersetzungsrechte auslesen
        userRights = getUserRight(aUserID, aProjectID); new TUserRight();

        // Projektplanrollen auslesen
        taskroleRights = new ArrayList();
        dataReader = new SqlDB("select ID, roleID from mapusertaskrole where userID=@userID ORDER BY roleID", parameterList, aProjectID);
        while (dataReader.read())
        {
            taskroleRights.Add(dataReader.getInt32(1));
        }
        dataReader.close();

        // Gruppen auslesen
        usergroupRights = new ArrayList();
        dataReader = new SqlDB("select ID, groupID from mapusergroup where userID=@userID ORDER BY groupID", parameterList, aProjectID);
        while (dataReader.read())
        {
            usergroupRights.Add(dataReader.getString(1));
        }
        dataReader.close();

        // Übersetzungssprachen auslesen
        translateLanguagesRights = new ArrayList();
        dataReader = new SqlDB("select ID, translatelanguageID from mapusertranslatelanguage where userID=@userID ORDER BY translatelanguageID", parameterList, aProjectID);
        while (dataReader.read())
        {
            translateLanguagesRights.Add(dataReader.getInt32(1));
        }
        dataReader.close();

        // Tool Berechtigungen
        toolRights = new ArrayList();
        dataReader = new SqlDB("select ID, userID, toolID from mapusertool where userID=@userID", parameterList, aProjectID);
        while (dataReader.read())
        {
            TToolRight tempTool = new TToolRight();
            tempTool.toolID = dataReader.getString(2);
            toolRights.Add(tempTool);
        }
        dataReader.close();
    }
    public void deleteRight(int aID)
    {
        TOrgRight actUnitRight = null;
        foreach (TOrgRight tempUnitRight in orgRights)
        {
            if (tempUnitRight.ID == aID)
                actUnitRight = tempUnitRight;
        }
        if (actUnitRight != null)
            orgRights.Remove(actUnitRight);
    }
    public void deleteBackRight(int aID)
    {
        TBackRight actUnitRight = null;
        foreach (TBackRight tempUnitRight in backRights)
        {
            if (tempUnitRight.ID == aID)
                actUnitRight = tempUnitRight;
        }
        if (actUnitRight != null)
            backRights.Remove(actUnitRight);
    }
    public void deleteOrg1Right(int aID)
    {
        TOrg1Right actUnitRight = null;
        foreach (TOrg1Right tempUnitRight in org1Rights)
        {
            if (tempUnitRight.ID == aID)
                actUnitRight = tempUnitRight;
        }
        if (actUnitRight != null)
            org1Rights.Remove(actUnitRight);
    }
    public void deleteBack1Right(int aID)
    {
        TBack1Right actUnitRight = null;
        foreach (TBack1Right tempUnitRight in back1Rights)
        {
            if (tempUnitRight.ID == aID)
                actUnitRight = tempUnitRight;
        }
        if (actUnitRight != null)
            back1Rights.Remove(actUnitRight);
    }
    public void deleteOrg2Right(int aID)
    {
        TOrg2Right actUnitRight = null;
        foreach (TOrg2Right tempUnitRight in org2Rights)
        {
            if (tempUnitRight.ID == aID)
                actUnitRight = tempUnitRight;
        }
        if (actUnitRight != null)
            org2Rights.Remove(actUnitRight);
    }
    public void deleteBack2Right(int aID)
    {
        TBack2Right actUnitRight = null;
        foreach (TBack2Right tempUnitRight in back2Rights)
        {
            if (tempUnitRight.ID == aID)
                actUnitRight = tempUnitRight;
        }
        if (actUnitRight != null)
            back2Rights.Remove(actUnitRight);
    }
    public void deleteOrg3Right(int aID)
    {
        TOrg3Right actUnitRight = null;
        foreach (TOrg3Right tempUnitRight in org3Rights)
        {
            if (tempUnitRight.ID == aID)
                actUnitRight = tempUnitRight;
        }
        if (actUnitRight != null)
            org3Rights.Remove(actUnitRight);
    }
    public void deleteBack3Right(int aID)
    {
        TBack3Right actUnitRight = null;
        foreach (TBack3Right tempUnitRight in back3Rights)
        {
            if (tempUnitRight.ID == aID)
                actUnitRight = tempUnitRight;
        }
        if (actUnitRight != null)
            back3Rights.Remove(actUnitRight);
    }
    public void deleteContainerRight(int aID)
    {
        TContainerRight actContainerRight = null;
        foreach (TContainerRight tempContainerRight in containerRights)
        {
            if (tempContainerRight.ID == aID)
                actContainerRight = tempContainerRight;
        }
        if (actContainerRight != null)
            containerRights.Remove(actContainerRight);
    }
    public TOrgRight getUnitRight(int aID)
    {
        TOrgRight actUnitRight = null;
        foreach (TOrgRight tempUnitRight in orgRights)
        {
            if (tempUnitRight.ID == aID)
                actUnitRight = tempUnitRight;
        }
        return actUnitRight;
    }
    public TBackRight getBackRight(int aID)
    {
        TBackRight actUnitRight = null;
        foreach (TBackRight tempUnitRight in backRights)
        {
            if (tempUnitRight.ID == aID)
                actUnitRight = tempUnitRight;
        }
        return actUnitRight;
    }
    public TOrg1Right getUnitRightOrg1(int aID)
    {
        TOrg1Right actUnitRight = null;
        foreach (TOrg1Right tempUnitRight in org1Rights)
        {
            if (tempUnitRight.ID == aID)
                actUnitRight = tempUnitRight;
        }
        return actUnitRight;
    }
    public TBack1Right getBack1Right(int aID)
    {
        TBack1Right actUnitRight = null;
        foreach (TBack1Right tempUnitRight in back1Rights)
        {
            if (tempUnitRight.ID == aID)
                actUnitRight = tempUnitRight;
        }
        return actUnitRight;
    }
    public TOrg2Right getUnitRightOrg2(int aID)
    {
        TOrg2Right actUnitRight = null;
        foreach (TOrg2Right tempUnitRight in org2Rights)
        {
            if (tempUnitRight.ID == aID)
                actUnitRight = tempUnitRight;
        }
        return actUnitRight;
    }
    public TBack2Right getBack2Right(int aID)
    {
        TBack2Right actUnitRight = null;
        foreach (TBack2Right tempUnitRight in back2Rights)
        {
            if (tempUnitRight.ID == aID)
                actUnitRight = tempUnitRight;
        }
        return actUnitRight;
    }
    public TOrg3Right getUnitRightOrg3(int aID)
    {
        TOrg3Right actUnitRight = null;
        foreach (TOrg3Right tempUnitRight in org3Rights)
        {
            if (tempUnitRight.ID == aID)
                actUnitRight = tempUnitRight;
        }
        return actUnitRight;
    }
    public TBack3Right getBack3Right(int aID)
    {
        TBack3Right actUnitRight = null;
        foreach (TBack3Right tempUnitRight in back3Rights)
        {
            if (tempUnitRight.ID == aID)
                actUnitRight = tempUnitRight;
        }
        return actUnitRight;
    }
    public TContainerRight getContainerRight(int aID)
    {
        TContainerRight actContainerRight = null;
        foreach (TContainerRight tempContainerRight in containerRights)
        {
            if (tempContainerRight.ID == aID)
                actContainerRight = tempContainerRight;
        }
        return actContainerRight;
    }
    public int getUnitRightIndex(int aID)
    {
        int result = -1;
        for (int index = 0; index < orgRights.Count; index++)
        {
            if (((TUser.TOrgRight)orgRights[index]).ID == aID)
                result = index;
        }
        return result;
    }
    public int getBackRightIndex(int aID)
    {
        int result = -1;
        for (int index = 0; index < backRights.Count; index++)
        {
            if (((TUser.TBackRight)backRights[index]).ID == aID)
                result = index;
        }
        return result;
    }
    public int getUnitRightIndexOrg1(int aID)
    {
        int result = -1;
        for (int index = 0; index < org1Rights.Count; index++)
        {
            if (((TUser.TOrg1Right)org1Rights[index]).ID == aID)
                result = index;
        }
        return result;
    }
    public int getBack1RightIndex(int aID)
    {
        int result = -1;
        for (int index = 0; index < back1Rights.Count; index++)
        {
            if (((TUser.TBack1Right)back1Rights[index]).ID == aID)
                result = index;
        }
        return result;
    }
    public int getUnitRightIndexOrg2(int aID)
    {
        int result = -1;
        for (int index = 0; index < org2Rights.Count; index++)
        {
            if (((TUser.TOrg2Right)org2Rights[index]).ID == aID)
                result = index;
        }
        return result;
    }
    public int getBack2RightIndex(int aID)
    {
        int result = -1;
        for (int index = 0; index < back2Rights.Count; index++)
        {
            if (((TUser.TBack2Right)back2Rights[index]).ID == aID)
                result = index;
        }
        return result;
    }
    public int getUnitRightIndexOrg3(int aID)
    {
        int result = -1;
        for (int index = 0; index < org3Rights.Count; index++)
        {
            if (((TUser.TOrg2Right)org3Rights[index]).ID == aID)
                result = index;
        }
        return result;
    }
    public int getBack3RightIndex(int aID)
    {
        int result = -1;
        for (int index = 0; index < back3Rights.Count; index++)
        {
            if (((TUser.TBack3Right)back3Rights[index]).ID == aID)
                result = index;
        }
        return result;
    }
    public int getContainerRightIndex(int aID)
    {
        int result = -1;
        for (int index = 0; index < containerRights.Count; index++)
        {
            if (((TUser.TContainerRight)containerRights[index]).ID == aID)
                result = index;
        }
        return result;
    }
    /// <summary>
    /// Ermittlung der Verfügbarkeit eines Tools für den Benutzer
    /// </summary>
    /// <param name="aToolName">ID des Tools</param>
    /// <returns>true wenn verfügbar, sonst false</returns>
    public bool getToolAvailable(string aToolName)
    {
        bool available = false;
        foreach (TToolRight tempTool in toolRights)
        {
            if (tempTool.toolID == aToolName)
                available = true;
        }
        return available;
    }
    /// <summary>
    /// Speichern eines neuen Benutzers in der DB
    /// </summary>
    /// <param name="aProjectID">ID des Projektes</param>
    /// <returns></returns>
    public bool save(string aProjectID, TUserPassword aUserPassword, bool aSaveUserRights)
    {
        //TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];
        SqlDB dataReader;

        // userID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("userID", "string", this.UserID);
        dataReader = new SqlDB("SELECT userID FROM loginuser WHERE userID=@userID", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            if (Password == "")
                Password = TCrypt.StringVerschluesseln(aUserPassword.generatePassword(UserID, aProjectID));
                //Password = TCrypt.StringVerschluesseln(Login.generatePassword(UserID, aPasswordLength, aPasswordComplexity, aPssswordHistory, aProjectID));
            //Password = TCrypt.StringVerschluesseln(Password);
            string tempActive = "0";
            if (Active)
                tempActive = "1";
            string tempInternalUser = "0";
            if (IsInternalUser)
                tempInternalUser = "1";
            string tempForce = "0";
            if (ForceChange)
                tempForce = "1";

            TCrypt cryptKey = new TCrypt();
            parameterList = new TParameterList();
            parameterList.addParameter("userID", "string", UserID);
            parameterList.addParameter("displayName", "string", DisplayName);
            parameterList.addParameter("password", "string", Password);
            parameterList.addParameter("email", "string", Email);
            parameterList.addParameter("title", "string", Title);
            parameterList.addParameter("name", "string", Name);
            parameterList.addParameter("firstname", "string", Firstname);
            parameterList.addParameter("street", "string", Street);
            parameterList.addParameter("zipcode", "string", Zipcode);
            parameterList.addParameter("city", "string", City);
            parameterList.addParameter("state", "string", State);
            parameterList.addParameter("lastLogin", "datetime", DateTime.Now.ToShortDateString());
            parameterList.addParameter("lastLoginError", "datetime", DateTime.Now.ToShortDateString());
            parameterList.addParameter("lastLoginErrorCount", "int", "0");
            parameterList.addParameter("isinternaluser", "int", tempInternalUser);
            parameterList.addParameter("forceChange", "int", tempForce);
            parameterList.addParameter("active", "int", tempActive);
            parameterList.addParameter("systemgeneratedpassword", "int", "1");
            parameterList.addParameter("passwordValidDate", "datetime", PasswordValidDate.ToString());
            parameterList.addParameter("language", "int", Language.ToString());
            parameterList.addParameter("comment", "string", Comment);
            parameterList.addParameter("groupID", "string", GroupID);
            parameterList.addParameter("isdelegate", "string", "0");
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO loginuser (userID, displayName, password, email, title, name, firstname, street, zipcode, city, state, lastLogin, lastLoginError, lastLoginErrorCount, isinternaluser, forceChange, active, systemgeneratedpassword, passwordValidDate, language, comment, groupID, isdelegate) VALUES (@userID, @displayName, @password, @email, @title, @name, @firstname, @street, @zipcode, @city, @state, @lastLogin, @lastLoginError, @lastLoginErrorCount, @isinternaluser, @forceChange, @active, @systemgeneratedpassword, @passwordValidDate, @language, @comment, @groupID, @isdelegate)", parameterList);
            // Tool erlauben
            foreach (TToolRight tempTool in toolRights)
            {
                parameterList = new TParameterList();
                parameterList.addParameter("userID", "string", UserID);
                parameterList.addParameter("toolID", "string", tempTool.toolID);
                dataReader = new SqlDB(aProjectID);
                dataReader.execSQLwithParameter("INSERT INTO mapusertool (userID, toolID) VALUES (@userID, @toolID)", parameterList);
            }
            // Unit-Berechtigungen
            foreach (TOrgRight tempUnitRight in orgRights)
            {
                TUser.saveUnitRight(this.UserID, tempUnitRight, aProjectID);
            }
            // Back-Berechtigungen
            foreach (TBackRight tempUnitRight in backRights)
            {
                TUser.saveUnitRightBack(this.UserID, tempUnitRight, aProjectID);
            }
            // Org1-Berechtigungen
            foreach (TOrg1Right tempUnitRight in org1Rights)
            {
                TUser.saveUnitRightOrg1(this.UserID, tempUnitRight, aProjectID);
            }
            // Org1-Back-Berechtigungen
            foreach (TBack1Right tempUnitRight in back1Rights)
            {
                TUser.saveUnitRightBack1(this.UserID, tempUnitRight, aProjectID);
            }
            // Org2-Berechtigungen
            foreach (TOrg2Right tempUnitRight in org2Rights)
            {
                TUser.saveUnitRightOrg2(this.UserID, tempUnitRight, aProjectID);
            }
            // Org2-Back-Berechtigungen
            foreach (TBack2Right tempUnitRight in back2Rights)
            {
                TUser.saveUnitRightBack2(this.UserID, tempUnitRight, aProjectID);
            }
            // Org3-Berechtigungen
            foreach (TOrg3Right tempUnitRight in org3Rights)
            {
                TUser.saveUnitRightOrg3(this.UserID, tempUnitRight, aProjectID);
            }
            // Org2-Back-Berechtigungen
            foreach (TBack3Right tempUnitRight in back3Rights)
            {
                TUser.saveUnitRightBack3(this.UserID, tempUnitRight, aProjectID);
            }
            // Container-Berechtigungen
            foreach (TContainerRight tempContainerRight in containerRights)
            {
                TUser.saveContainerRight(this.UserID, tempContainerRight, aProjectID);
            }
            // Rollenberechtigungen
            foreach (string tempTaskRoleRight in taskroleRights)
            {
                TUser.saveTaskRoleRight(this.UserID, Convert.ToInt32(tempTaskRoleRight), aProjectID);
            }
            // Gruppenberechtigungen
            foreach (string tempGroupRight in usergroupRights)
            {
                TUser.saveGroupRight(this.UserID, tempGroupRight, aProjectID);
            }
            // Übersetzungssprachenberechtigungen
            foreach (string tempTranslateLanguagesRight in translateLanguagesRights)
            {
                TUser.saveTranslateLanguageRight(this.UserID, Convert.ToInt32(tempTranslateLanguagesRight), aProjectID);
            }
            if (aSaveUserRights)
                saveUserRight(this.UserID, userRights, aProjectID);
            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveUnitRight(string aUserID, TOrgRight aUnitRight, string aProjectID)
    {
        SqlDB dataReader;

        // userID/orgID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("userID", "string", aUserID);
        dataReader = new SqlDB("SELECT ID FROM mapuserorgid WHERE userID=@userID AND orgID='" + aUnitRight.OrgID.ToString() + "'", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            string tempDelegate = "0";
            if (aUnitRight.OAllowDelegate)
                tempDelegate = "1";
            string tempLock = "0";
            if (aUnitRight.OAllowLock)
                tempLock = "1";
            string tempEmployeeRead = "0";
            if (aUnitRight.OAllowEmployeeRead)
                tempEmployeeRead = "1";
            string tempEmployeeEdit = "0";
            if (aUnitRight.OAllowEmployeeEdit)
                tempEmployeeEdit = "1";
            string tempEmployeeImport = "0";
            if (aUnitRight.OAllowEmployeeImport)
                tempEmployeeImport = "1";
            string tempEmployeeExport = "0";
            if (aUnitRight.OAllowEmployeeExport)
                tempEmployeeExport = "1";
            string tempUnitAdd = "0";
            if (aUnitRight.OAllowUnitAdd)
                tempUnitAdd = "1";
            string tempUnitMove = "0";
            if (aUnitRight.OAllowUnitMove)
                tempUnitMove = "1";
            string tempUnitDelete = "0";
            if (aUnitRight.OAllowUnitDelete)
                tempUnitDelete = "1";
            string tempUnitProperty = "0";
            if (aUnitRight.OAllowUnitProperty)
                tempUnitProperty = "1";
            string tempReportRecipient = "0";
            if (aUnitRight.OAllowReportRecipient)
                tempReportRecipient = "1";
            string tempStructureImport = "0";
            if (aUnitRight.OAllowStructureImport)
                tempStructureImport = "1";
            string tempStructureExport = "0";
            if (aUnitRight.OAllowStructureExport)
                tempStructureExport = "1";
            string tempBouncesRead = "0";
            if (aUnitRight.OAllowBouncesRead)
                tempBouncesRead = "1";
            string tempBouncesEdit = "0";
            if (aUnitRight.OAllowBouncesEdit)
                tempBouncesEdit = "1";
            string tempBouncesDelete = "0";
            if (aUnitRight.OAllowBouncesDelete)
                tempBouncesDelete = "1";
            string tempBouncesExport = "0";
            if (aUnitRight.OAllowBouncesExport)
                tempBouncesExport = "1";
            string tempBouncesImport = "0";
            if (aUnitRight.OAllowBouncesImport)
                tempBouncesImport = "1";
            string tempAllowReInvition = "0";
            if (aUnitRight.OAllowReInvitation)
                tempAllowReInvition = "1";
            string tempAllowPostNomination = "0";
            if (aUnitRight.OAllowPostNomination)
                tempAllowPostNomination = "1";
            string tempAllowPostNominationImport = "0";
            if (aUnitRight.OAllowPostNominationImport)
                tempAllowPostNominationImport = "1";
            string tempAllowCommunications = "0";
            if (aUnitRight.FAllowCommunications)
                tempAllowCommunications = "1";
            string tempAllowMeasures = "0";
            if (aUnitRight.FAllowMeasures)
                tempAllowMeasures = "1";
            string tempDelete = "0";
            if (aUnitRight.FAllowDelete)
                tempDelete = "1";
            string tempExcelSummary = "0";
            if (aUnitRight.FAllowExcelExport)
                tempExcelSummary = "1";
            string tempAllowReminderOnMeasure = "0";
            if (aUnitRight.FAllowReminderOnMeasure)
                tempAllowReminderOnMeasure = "1";
            string tempAllowReminderOnUnit = "0";
            if (aUnitRight.FAllowReminderOnUnit)
                tempAllowReminderOnUnit = "1";
            string tempAllowUpUnitsMeasures = "0";
            if (aUnitRight.FAllowUpUnitsMeasures)
                tempAllowUpUnitsMeasures = "1";
            string tempAllowColleaguesMeasures = "0";
            if (aUnitRight.FAllowColleaguesMeasures)
                tempAllowColleaguesMeasures = "1";

            string tempAllowThemes = "0";
            if (aUnitRight.FAllowThemes)
                tempAllowThemes = "1";
            string tempAllowReminderOnThemes = "0";
            if (aUnitRight.FAllowReminderOnThemes)
                tempAllowReminderOnThemes = "1";
            string tempAllowReminderOnThemesUnit = "0";
            if (aUnitRight.FAllowReminderOnThemeUnit)
                tempAllowReminderOnThemesUnit = "1";
            string tempAllowTopUnitsThemes = "0";
            if (aUnitRight.FAllowTopUnitsThemes)
                tempAllowTopUnitsThemes = "1";
            string tempAllowColleguesThemes = "0";
            if (aUnitRight.FAllowColleaguesThemes)
                tempAllowColleguesThemes = "1";
            string tempAllowOrders = "0";
            if (aUnitRight.FAllowOrder)
                tempAllowOrders = "1";
            string tempAllowReminderOnOrdersUnit = "0";
            if (aUnitRight.FAllowReminderOnOrderUnit)
                tempAllowReminderOnOrdersUnit = "1";
            
            string tempMail = "0";
            if (aUnitRight.DAllowMail)
                tempMail = "1";
            string tempDownload = "0";
            if (aUnitRight.DAllowDownload)
                tempDownload = "1";
            string tempLog = "0";
            if (aUnitRight.DAllowLog)
                tempLog = "1";
            string tempFileType1 = "0";
            if (aUnitRight.DAllowType1)
                tempFileType1 = "1";
            string tempFileType2 = "0";
            if (aUnitRight.DAllowType2)
                tempFileType2 = "1";
            string tempFileType3 = "0";
            if (aUnitRight.DAllowType3)
                tempFileType3 = "1";
            string tempFileType4 = "0";
            if (aUnitRight.DAllowType4)
                tempFileType4 = "1";
            string tempFileType5 = "0";
            if (aUnitRight.DAllowType5)
                tempFileType5 = "1";
            parameterList = new TParameterList();
            parameterList.addParameter("userID", "string", aUserID);
            parameterList.addParameter("orgID", "int", aUnitRight.OrgID.ToString());
            parameterList.addParameter("OProfile", "string", aUnitRight.OProfile);
            parameterList.addParameter("OReadLevel", "int", aUnitRight.OReadLevel.ToString());
            parameterList.addParameter("OEditLevel", "int", aUnitRight.OEditLevel.ToString());
            parameterList.addParameter("OAllowDelegate", "int", tempDelegate);
            parameterList.addParameter("OAllowLock", "int", tempLock);
            parameterList.addParameter("OAllowEmployeeRead", "int", tempEmployeeRead);
            parameterList.addParameter("OAllowEmployeeEdit", "int", tempEmployeeEdit);
            parameterList.addParameter("OAllowEmployeeImport", "int", tempEmployeeImport);
            parameterList.addParameter("OAllowEmployeeExport", "int", tempEmployeeExport);
            parameterList.addParameter("OAllowUnitAdd", "int", tempUnitAdd);
            parameterList.addParameter("OAllowUnitMove", "int", tempUnitMove);
            parameterList.addParameter("OAllowUnitDelete", "int", tempUnitDelete);
            parameterList.addParameter("OAllowUnitProperty", "int", tempUnitProperty);
            parameterList.addParameter("OAllowReportRecipient", "int", tempReportRecipient);
            parameterList.addParameter("OAllowStructureImport", "int", tempStructureImport);
            parameterList.addParameter("OAllowStructureExport", "int", tempStructureExport);
            parameterList.addParameter("OAllowBouncesRead", "int", tempBouncesRead);
            parameterList.addParameter("OAllowBouncesEdit", "int", tempBouncesEdit);
            parameterList.addParameter("OAllowBouncesDelete", "int", tempBouncesDelete);
            parameterList.addParameter("OAllowBouncesExport", "int", tempBouncesExport);
            parameterList.addParameter("OAllowBouncesImport", "int", tempBouncesImport);
            parameterList.addParameter("OAllowReInvitation", "int", tempAllowReInvition);
            parameterList.addParameter("OAllowPostNomination", "int", tempAllowPostNomination);
            parameterList.addParameter("OAllowPostNominationImport", "int", tempAllowPostNominationImport);
            parameterList.addParameter("FProfile", "string", aUnitRight.FProfile);
            parameterList.addParameter("FReadLevel", "int", aUnitRight.FReadLevel.ToString());
            parameterList.addParameter("FEditLevel", "int", aUnitRight.FEditLevel.ToString());
            parameterList.addParameter("FSumLevel", "int", aUnitRight.FSumLevel.ToString());
            parameterList.addParameter("FAllowCommunications", "int", tempAllowCommunications.ToString());
            parameterList.addParameter("FAllowMeasures", "int", tempAllowMeasures.ToString());
            parameterList.addParameter("FAllowDelete", "int", tempDelete.ToString());
            parameterList.addParameter("FAllowExcelExport", "int", tempExcelSummary.ToString());
            parameterList.addParameter("FAllowReminderOnMeasure", "int", tempAllowReminderOnMeasure.ToString());
            parameterList.addParameter("FAllowReminderOnUnit", "int", tempAllowReminderOnUnit.ToString());
            parameterList.addParameter("FAllowUpUnitsMeasures", "int", tempAllowUpUnitsMeasures.ToString());
            parameterList.addParameter("FAllowColleaguesMeasures", "int", tempAllowColleaguesMeasures.ToString());
            parameterList.addParameter("FAllowThemes", "int", tempAllowThemes.ToString());
            parameterList.addParameter("FAllowReminderOnTheme", "int", tempAllowReminderOnThemes.ToString());
            parameterList.addParameter("FAllowReminderOnThemeUnit", "int", tempAllowReminderOnThemesUnit.ToString());
            parameterList.addParameter("FAllowTopUnitsThemes", "int", tempAllowTopUnitsThemes.ToString());
            parameterList.addParameter("FAllowColleaguesThemes", "int", tempAllowColleguesThemes.ToString());
            parameterList.addParameter("FAllowOrders", "int", tempAllowOrders.ToString());
            parameterList.addParameter("FAllowReminderOnOrderUnit", "int", tempAllowReminderOnOrdersUnit.ToString());

            parameterList.addParameter("RProfile", "string", aUnitRight.RProfile);
            parameterList.addParameter("RReadLevel", "int", aUnitRight.RReadLevel.ToString());
            parameterList.addParameter("DProfile", "string", aUnitRight.DProfile);
            parameterList.addParameter("DReadLevel", "int", aUnitRight.DReadLevel.ToString());
            parameterList.addParameter("DAllowMail", "int", tempMail);
            parameterList.addParameter("DAllowDownload", "int", tempDownload);
            parameterList.addParameter("DAllowLog", "int", tempLog);
            parameterList.addParameter("DAllowType1", "int", tempFileType1);
            parameterList.addParameter("DAllowType2", "int", tempFileType2);
            parameterList.addParameter("DAllowType3", "int", tempFileType3);
            parameterList.addParameter("DAllowType4", "int", tempFileType4);
            parameterList.addParameter("DAllowType5", "int", tempFileType5);
            parameterList.addParameter("AEditLevelWorkshop", "int", aUnitRight.AEditLevelWorkshop.ToString());
            parameterList.addParameter("AReadLevelStructure", "int", aUnitRight.AReadLevelStructure.ToString());
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO mapuserorgid (userID, orgID, OProfile, OReadLevel, OEditLevel, OAllowDelegate, OAllowLock, OAllowEmployeeRead, OAllowEmployeeEdit, OAllowEmployeeImport, OAllowEmployeeExport, OAllowUnitAdd, OAllowUnitMove, OAllowUnitDelete, OAllowUnitProperty, OAllowReportRecipient, OAllowStructureImport, OAllowStructureExport, OAllowBouncesRead, OAllowBouncesEdit, OAllowBouncesDelete, OAllowBouncesExport, OAllowBouncesImport, OAllowReInvitation, OAllowPostNomination, OAllowPostNominationImport, FProfile, FReadLevel, FEditLevel, FSumLevel, FAllowCommunications, FAllowMeasures, FAllowDelete, FAllowExcelExport, FAllowReminderOnMeasure, FAllowReminderOnUnit, FAllowUpUnitsMeasures, FAllowColleaguesMeasures, FAllowThemes, FAllowReminderOnTheme, FAllowReminderOnThemeUnit, FAllowTopUnitsThemes, FAllowColleaguesThemes, FAllowOrders, FAllowReminderOnOrderUnit, RProfile, RReadLevel, DProfile, DReadLevel, DAllowMail, DAllowDownload, DAllowLog, DAllowType1, DAllowType2, DAllowType3, DAllowType4, DAllowType5, AEditLevelWorkshop, AReadLevelStructure) VALUES (@userID, @orgID, @OProfile, @OReadLevel, @OEditLevel, @OAllowDelegate, @OAllowLock, @OAllowEmployeeRead, @OAllowEmployeeEdit, @OAllowEmployeeImport, @OAllowEmployeeExport, @OAllowUnitAdd, @OAllowUnitMove, @OAllowUnitDelete, @OAllowUnitProperty, @OAllowReportRecipient, @OAllowStructureImport, @OAllowStructureExport, @OAllowBouncesRead, @OAllowBouncesEdit, @OAllowBouncesDelete, @OAllowBouncesExport, @OAllowBouncesImport, @OAllowReInvitation, @OAllowPostNomination, @OAllowPostNominationImport, @FProfile, @FReadLevel, @FEditLevel, @FSumLevel, @FAllowCommunications, @FAllowMeasures, @FAllowDelete, @FAllowExcelExport, @FAllowReminderOnMeasure, @FAllowReminderOnUnit, @FAllowUpUnitsMeasures, @FAllowColleaguesMeasures, @FAllowThemes, @FAllowReminderOnTheme, @FAllowReminderOnThemeUnit, @FAllowTopUnitsThemes, @FAllowColleaguesThemes, @FAllowOrders, @FAllowReminderOnOrderUnit, @RProfile, @RReadLevel, @DProfile, @DReadLevel, @DAllowMail, @DAllowDownload, @DAllowLog, @DAllowType1, @DAllowType2, @DAllowType3, @DAllowType4, @DAllowType5, @AEditLevelWorkshop, @AReadLevelStructure)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveUnitRightBack(string aUserID, TBackRight aUnitRight, string aProjectID)
    {
        SqlDB dataReader;

        // userID/orgID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("userID", "string", aUserID);
        dataReader = new SqlDB("SELECT ID FROM mapuserback WHERE userID=@userID AND orgID='" + aUnitRight.OrgID.ToString() + "'", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            parameterList = new TParameterList();
            parameterList.addParameter("userID", "string", aUserID);
            parameterList.addParameter("orgID", "int", aUnitRight.OrgID.ToString());
            parameterList.addParameter("OProfile", "string", aUnitRight.OProfile);
            parameterList.addParameter("OReadLevel", "int", aUnitRight.OReadLevel.ToString());
            parameterList.addParameter("AReadLevel", "int", aUnitRight.AReadLevel.ToString());
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO mapuserback (userID, orgID, OProfile, OReadLevel, AReadLevel) VALUES (@userID, @orgID, @OProfile, @OReadLevel, @AReadLevel)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveUnitRightOrg1(string aUserID, TOrg1Right aUnitRight, string aProjectID)
    {
        SqlDB dataReader;

        // userID/orgID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("userID", "string", aUserID);
        dataReader = new SqlDB("SELECT ID FROM mapuserorgid1 WHERE userID=@userID AND orgID='" + aUnitRight.OrgID.ToString() + "'", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            string tempDelegate = "0";
            if (aUnitRight.OAllowDelegate)
                tempDelegate = "1";
            string tempLock = "0";
            if (aUnitRight.OAllowLock)
                tempLock = "1";
            string tempEmployeeRead = "0";
            if (aUnitRight.OAllowEmployeeRead)
                tempEmployeeRead = "1";
            string tempEmployeeEdit = "0";
            if (aUnitRight.OAllowEmployeeEdit)
                tempEmployeeEdit = "1";
            string tempEmployeeImport = "0";
            if (aUnitRight.OAllowEmployeeImport)
                tempEmployeeImport = "1";
            string tempEmployeeExport = "0";
            if (aUnitRight.OAllowEmployeeExport)
                tempEmployeeExport = "1";
            string tempUnitAdd = "0";
            if (aUnitRight.OAllowUnitAdd)
                tempUnitAdd = "1";
            string tempUnitMove = "0";
            if (aUnitRight.OAllowUnitMove)
                tempUnitMove = "1";
            string tempUnitDelete = "0";
            if (aUnitRight.OAllowUnitDelete)
                tempUnitDelete = "1";
            string tempUnitProperty = "0";
            if (aUnitRight.OAllowUnitProperty)
                tempUnitProperty = "1";
            string tempReportRecipient = "0";
            if (aUnitRight.OAllowReportRecipient)
                tempReportRecipient = "1";
            string tempStructureImport = "0";
            if (aUnitRight.OAllowStructureImport)
                tempStructureImport = "1";
            string tempStructureExport = "0";
            if (aUnitRight.OAllowStructureExport)
                tempStructureExport = "1";
            string tempAllowCommunications = "0";
            if (aUnitRight.FAllowCommunications)
                tempAllowCommunications = "1";
            string tempAllowMeasures = "0";
            if (aUnitRight.FAllowMeasures)
                tempAllowMeasures = "1";
            string tempDelete = "0";
            if (aUnitRight.FAllowDelete)
                tempDelete = "1";
            string tempExcelSummary = "0";
            if (aUnitRight.FAllowExcelExport)
                tempExcelSummary = "1";
            string tempAllowReminderOnMeasure = "0";
            if (aUnitRight.FAllowReminderOnMeasure)
                tempAllowReminderOnMeasure = "1";
            string tempAllowReminderOnUnit = "0";
            if (aUnitRight.FAllowReminderOnUnit)
                tempAllowReminderOnUnit = "1";
            string tempAllowUpUnitsMeasures = "0";
            if (aUnitRight.FAllowUpUnitsMeasures)
                tempAllowUpUnitsMeasures = "1";
            string tempAllowColleaguesMeasures = "0";
            if (aUnitRight.FAllowColleaguesMeasures)
                tempAllowColleaguesMeasures = "1";

            string tempAllowThemes = "0";
            if (aUnitRight.FAllowThemes)
                tempAllowThemes = "1";
            string tempAllowReminderOnThemes = "0";
            if (aUnitRight.FAllowReminderOnThemes)
                tempAllowReminderOnThemes = "1";
            string tempAllowReminderOnThemesUnit = "0";
            if (aUnitRight.FAllowReminderOnThemeUnit)
                tempAllowReminderOnThemesUnit = "1";
            string tempAllowTopUnitsThemes = "0";
            if (aUnitRight.FAllowTopUnitsThemes)
                tempAllowTopUnitsThemes = "1";
            string tempAllowColleguesThemes = "0";
            if (aUnitRight.FAllowColleaguesThemes)
                tempAllowColleguesThemes = "1";
            string tempAllowOrders = "0";
            if (aUnitRight.FAllowOrder)
                tempAllowOrders = "1";
            string tempAllowReminderOnOrdersUnit = "0";
            if (aUnitRight.FAllowReminderOnOrderUnit)
                tempAllowReminderOnOrdersUnit = "1";

            string tempMail = "0";
            if (aUnitRight.DAllowMail)
                tempMail = "1";
            string tempDownload = "0";
            if (aUnitRight.DAllowDownload)
                tempDownload = "1";
            string tempLog = "0";
            if (aUnitRight.DAllowLog)
                tempLog = "1";
            string tempFileType1 = "0";
            if (aUnitRight.DAllowType1)
                tempFileType1 = "1";
            string tempFileType2 = "0";
            if (aUnitRight.DAllowType2)
                tempFileType2 = "1";
            string tempFileType3 = "0";
            if (aUnitRight.DAllowType3)
                tempFileType3 = "1";
            string tempFileType4 = "0";
            if (aUnitRight.DAllowType4)
                tempFileType4 = "1";
            string tempFileType5 = "0";
            if (aUnitRight.DAllowType5)
                tempFileType5 = "1";
            parameterList = new TParameterList();
            parameterList.addParameter("userID", "string", aUserID);
            parameterList.addParameter("orgID", "int", aUnitRight.OrgID.ToString());
            parameterList.addParameter("OProfile", "string", aUnitRight.OProfile);
            parameterList.addParameter("OReadLevel", "int", aUnitRight.OReadLevel.ToString());
            parameterList.addParameter("OEditLevel", "int", aUnitRight.OEditLevel.ToString());
            parameterList.addParameter("OAllowDelegate", "int", tempDelegate);
            parameterList.addParameter("OAllowLock", "int", tempLock);
            parameterList.addParameter("OAllowEmployeeRead", "int", tempEmployeeRead);
            parameterList.addParameter("OAllowEmployeeEdit", "int", tempEmployeeEdit);
            parameterList.addParameter("OAllowEmployeeImport", "int", tempEmployeeImport);
            parameterList.addParameter("OAllowEmployeeExport", "int", tempEmployeeExport);
            parameterList.addParameter("OAllowUnitAdd", "int", tempUnitAdd);
            parameterList.addParameter("OAllowUnitMove", "int", tempUnitMove);
            parameterList.addParameter("OAllowUnitDelete", "int", tempUnitDelete);
            parameterList.addParameter("OAllowUnitProperty", "int", tempUnitProperty);
            parameterList.addParameter("OAllowReportRecipient", "int", tempReportRecipient);
            parameterList.addParameter("OAllowStructureImport", "int", tempStructureImport);
            parameterList.addParameter("OAllowStructureExport", "int", tempStructureExport);
            parameterList.addParameter("FProfile", "string", aUnitRight.FProfile);
            parameterList.addParameter("FReadLevel", "int", aUnitRight.FReadLevel.ToString());
            parameterList.addParameter("FEditLevel", "int", aUnitRight.FEditLevel.ToString());
            parameterList.addParameter("FSumLevel", "int", aUnitRight.FSumLevel.ToString());
            parameterList.addParameter("FAllowCommunications", "int", tempAllowCommunications.ToString());
            parameterList.addParameter("FAllowMeasures", "int", tempAllowMeasures.ToString());
            parameterList.addParameter("FAllowDelete", "int", tempDelete.ToString());
            parameterList.addParameter("FAllowExcelExport", "int", tempExcelSummary.ToString());
            parameterList.addParameter("FAllowReminderOnMeasure", "int", tempAllowReminderOnMeasure.ToString());
            parameterList.addParameter("FAllowReminderOnUnit", "int", tempAllowReminderOnUnit.ToString());
            parameterList.addParameter("FAllowUpUnitsMeasures", "int", tempAllowUpUnitsMeasures.ToString());
            parameterList.addParameter("FAllowColleaguesMeasures", "int", tempAllowColleaguesMeasures.ToString());
            parameterList.addParameter("FAllowThemes", "int", tempAllowThemes.ToString());
            parameterList.addParameter("FAllowReminderOnTheme", "int", tempAllowReminderOnThemes.ToString());
            parameterList.addParameter("FAllowReminderOnThemeUnit", "int", tempAllowReminderOnThemesUnit.ToString());
            parameterList.addParameter("FAllowTopUnitsThemes", "int", tempAllowTopUnitsThemes.ToString());
            parameterList.addParameter("FAllowColleaguesThemes", "int", tempAllowColleguesThemes.ToString());
            parameterList.addParameter("FAllowOrders", "int", tempAllowOrders.ToString());
            parameterList.addParameter("FAllowReminderOnOrderUnit", "int", tempAllowReminderOnOrdersUnit.ToString());

            parameterList.addParameter("RProfile", "string", aUnitRight.RProfile);
            parameterList.addParameter("RReadLevel", "int", aUnitRight.RReadLevel.ToString());
            parameterList.addParameter("DProfile", "string", aUnitRight.DProfile);
            parameterList.addParameter("DReadLevel", "int", aUnitRight.DReadLevel.ToString());
            parameterList.addParameter("DAllowMail", "int", tempMail);
            parameterList.addParameter("DAllowDownload", "int", tempDownload);
            parameterList.addParameter("DAllowLog", "int", tempLog);
            parameterList.addParameter("DAllowType1", "int", tempFileType1);
            parameterList.addParameter("DAllowType2", "int", tempFileType2);
            parameterList.addParameter("DAllowType3", "int", tempFileType3);
            parameterList.addParameter("DAllowType4", "int", tempFileType4);
            parameterList.addParameter("DAllowType5", "int", tempFileType5);
            parameterList.addParameter("AEditLevelWorkshop", "int", aUnitRight.AEditLevelWorkshop.ToString());
            parameterList.addParameter("AReadLevelStructure", "int", aUnitRight.AReadLevelStructure.ToString());
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO mapuserorgid1 (userID, orgID, OProfile, OReadLevel, OEditLevel, OAllowDelegate, OAllowLock, OAllowEmployeeRead, OAllowEmployeeEdit, OAllowEmployeeImport, OAllowEmployeeExport, OAllowUnitAdd, OAllowUnitMove, OAllowUnitDelete, OAllowUnitProperty, OAllowReportRecipient, OAllowStructureImport, OAllowStructureExport, FProfile, FReadLevel, FEditLevel, FSumLevel, FAllowCommunications, FAllowMeasures, FAllowDelete, FAllowExcelExport, FAllowReminderOnMeasure, FAllowReminderOnUnit, FAllowUpUnitsMeasures, FAllowColleaguesMeasures, FAllowThemes, FAllowReminderOnTheme, FAllowReminderOnThemeUnit, FAllowTopUnitsThemes, FAllowColleaguesThemes, FAllowOrders, FAllowReminderOnOrderUnit, RProfile, RReadLevel, DProfile, DReadLevel, DAllowMail, DAllowDownload, DAllowLog, DAllowType1, DAllowType2, DAllowType3, DAllowType4, DAllowType5, AEditLevelWorkshop, AReadLevelStructure) VALUES (@userID, @orgID, @OProfile, @OReadLevel, @OEditLevel, @OAllowDelegate, @OAllowLock, @OAllowEmployeeRead, @OAllowEmployeeEdit, @OAllowEmployeeImport, @OAllowEmployeeExport, @OAllowUnitAdd, @OAllowUnitMove, @OAllowUnitDelete, @OAllowUnitProperty, @OAllowReportRecipient, @OAllowStructureImport, @OAllowStructureExport, @FProfile, @FReadLevel, @FEditLevel, @FSumLevel, @FAllowCommunications, @FAllowMeasures, @FAllowDelete, @FAllowExcelExport, @FAllowReminderOnMeasure, @FAllowReminderOnUnit, @FAllowUpUnitsMeasures, @FAllowColleaguesMeasures, @FAllowThemes, @FAllowReminderOnTheme, @FAllowReminderOnThemeUnit, @FAllowTopUnitsThemes, @FAllowColleaguesThemes, @FAllowOrders, @FAllowReminderOnOrderUnit, @RProfile, @RReadLevel, @DProfile, @DReadLevel, @DAllowMail, @DAllowDownload, @DAllowLog, @DAllowType1, @DAllowType2, @DAllowType3, @DAllowType4, @DAllowType5, @AEditLevelWorkshop, @AReadLevelStructure)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveUnitRightBack1(string aUserID, TBack1Right aUnitRight, string aProjectID)
    {
        SqlDB dataReader;

        // userID/orgID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("userID", "string", aUserID);
        dataReader = new SqlDB("SELECT ID FROM mapuserback1 WHERE userID=@userID AND orgID='" + aUnitRight.OrgID.ToString() + "'", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            parameterList = new TParameterList();
            parameterList.addParameter("userID", "string", aUserID);
            parameterList.addParameter("orgID", "int", aUnitRight.OrgID.ToString());
            parameterList.addParameter("OProfile", "string", aUnitRight.OProfile);
            parameterList.addParameter("OReadLevel", "int", aUnitRight.OReadLevel.ToString());
            parameterList.addParameter("AReadLevel", "int", aUnitRight.AReadLevel.ToString());
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO mapuserback1 (userID, orgID, OProfile, OReadLevel, AReadLevel) VALUES (@userID, @orgID, @OProfile, @OReadLevel, @AReadLevel)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveUnitRightOrg2(string aUserID, TOrg2Right aUnitRight, string aProjectID)
    {
        SqlDB dataReader;

        // userID/orgID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("userID", "string", aUserID);
        dataReader = new SqlDB("SELECT ID FROM mapuserorgid2 WHERE userID=@userID AND orgID='" + aUnitRight.OrgID.ToString() + "'", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            string tempDelegate = "0";
            if (aUnitRight.OAllowDelegate)
                tempDelegate = "1";
            string tempLock = "0";
            if (aUnitRight.OAllowLock)
                tempLock = "1";
            string tempEmployeeRead = "0";
            if (aUnitRight.OAllowEmployeeRead)
                tempEmployeeRead = "1";
            string tempEmployeeEdit = "0";
            if (aUnitRight.OAllowEmployeeEdit)
                tempEmployeeEdit = "1";
            string tempEmployeeImport = "0";
            if (aUnitRight.OAllowEmployeeImport)
                tempEmployeeImport = "1";
            string tempEmployeeExport = "0";
            if (aUnitRight.OAllowEmployeeExport)
                tempEmployeeExport = "1";
            string tempUnitAdd = "0";
            if (aUnitRight.OAllowUnitAdd)
                tempUnitAdd = "1";
            string tempUnitMove = "0";
            if (aUnitRight.OAllowUnitMove)
                tempUnitMove = "1";
            string tempUnitDelete = "0";
            if (aUnitRight.OAllowUnitDelete)
                tempUnitDelete = "1";
            string tempUnitProperty = "0";
            if (aUnitRight.OAllowUnitProperty)
                tempUnitProperty = "1";
            string tempReportRecipient = "0";
            if (aUnitRight.OAllowReportRecipient)
                tempReportRecipient = "1";
            string tempStructureImport = "0";
            if (aUnitRight.OAllowStructureImport)
                tempStructureImport = "1";
            string tempStructureExport = "0";
            if (aUnitRight.OAllowStructureExport)
                tempStructureExport = "1";
            string tempAllowCommunications = "0";
            if (aUnitRight.FAllowCommunications)
                tempAllowCommunications = "1";
            string tempAllowMeasures = "0";
            if (aUnitRight.FAllowMeasures)
                tempAllowMeasures = "1";
            string tempDelete = "0";
            if (aUnitRight.FAllowDelete)
                tempDelete = "1";
            string tempExcelSummary = "0";
            if (aUnitRight.FAllowExcelExport)
                tempExcelSummary = "1";
            string tempAllowReminderOnMeasure = "0";
            if (aUnitRight.FAllowReminderOnMeasure)
                tempAllowReminderOnMeasure = "1";
            string tempAllowReminderOnUnit = "0";
            if (aUnitRight.FAllowReminderOnUnit)
                tempAllowReminderOnUnit = "1";
            string tempAllowUpUnitsMeasures = "0";
            if (aUnitRight.FAllowUpUnitsMeasures)
                tempAllowUpUnitsMeasures = "1";
            string tempAllowColleaguesMeasures = "0";
            if (aUnitRight.FAllowColleaguesMeasures)
                tempAllowColleaguesMeasures = "1";

            string tempAllowThemes = "0";
            if (aUnitRight.FAllowThemes)
                tempAllowThemes = "1";
            string tempAllowReminderOnThemes = "0";
            if (aUnitRight.FAllowReminderOnThemes)
                tempAllowReminderOnThemes = "1";
            string tempAllowReminderOnThemesUnit = "0";
            if (aUnitRight.FAllowReminderOnThemeUnit)
                tempAllowReminderOnThemesUnit = "1";
            string tempAllowTopUnitsThemes = "0";
            if (aUnitRight.FAllowTopUnitsThemes)
                tempAllowTopUnitsThemes = "1";
            string tempAllowColleguesThemes = "0";
            if (aUnitRight.FAllowColleaguesThemes)
                tempAllowColleguesThemes = "1";
            string tempAllowOrders = "0";
            if (aUnitRight.FAllowOrder)
                tempAllowOrders = "1";
            string tempAllowReminderOnOrdersUnit = "0";
            if (aUnitRight.FAllowReminderOnOrderUnit)
                tempAllowReminderOnOrdersUnit = "1";

            string tempMail = "0";
            if (aUnitRight.DAllowMail)
                tempMail = "1";
            string tempDownload = "0";
            if (aUnitRight.DAllowDownload)
                tempDownload = "1";
            string tempLog = "0";
            if (aUnitRight.DAllowLog)
                tempLog = "1";
            string tempFileType1 = "0";
            if (aUnitRight.DAllowType1)
                tempFileType1 = "1";
            string tempFileType2 = "0";
            if (aUnitRight.DAllowType2)
                tempFileType2 = "1";
            string tempFileType3 = "0";
            if (aUnitRight.DAllowType3)
                tempFileType3 = "1";
            string tempFileType4 = "0";
            if (aUnitRight.DAllowType4)
                tempFileType4 = "1";
            string tempFileType5 = "0";
            if (aUnitRight.DAllowType5)
                tempFileType5 = "1";
            parameterList = new TParameterList();
            parameterList.addParameter("userID", "string", aUserID);
            parameterList.addParameter("orgID", "int", aUnitRight.OrgID.ToString());
            parameterList.addParameter("OProfile", "string", aUnitRight.OProfile);
            parameterList.addParameter("OReadLevel", "int", aUnitRight.OReadLevel.ToString());
            parameterList.addParameter("OEditLevel", "int", aUnitRight.OEditLevel.ToString());
            parameterList.addParameter("OAllowDelegate", "int", tempDelegate);
            parameterList.addParameter("OAllowLock", "int", tempLock);
            parameterList.addParameter("OAllowEmployeeRead", "int", tempEmployeeRead);
            parameterList.addParameter("OAllowEmployeeEdit", "int", tempEmployeeEdit);
            parameterList.addParameter("OAllowEmployeeImport", "int", tempEmployeeImport);
            parameterList.addParameter("OAllowEmployeeExport", "int", tempEmployeeExport);
            parameterList.addParameter("OAllowUnitAdd", "int", tempUnitAdd);
            parameterList.addParameter("OAllowUnitMove", "int", tempUnitMove);
            parameterList.addParameter("OAllowUnitDelete", "int", tempUnitDelete);
            parameterList.addParameter("OAllowUnitProperty", "int", tempUnitProperty);
            parameterList.addParameter("OAllowReportRecipient", "int", tempReportRecipient);
            parameterList.addParameter("OAllowStructureImport", "int", tempStructureImport);
            parameterList.addParameter("OAllowStructureExport", "int", tempStructureExport);
            parameterList.addParameter("FProfile", "string", aUnitRight.FProfile);
            parameterList.addParameter("FReadLevel", "int", aUnitRight.FReadLevel.ToString());
            parameterList.addParameter("FEditLevel", "int", aUnitRight.FEditLevel.ToString());
            parameterList.addParameter("FSumLevel", "int", aUnitRight.FSumLevel.ToString());
            parameterList.addParameter("FAllowCommunications", "int", tempAllowCommunications.ToString());
            parameterList.addParameter("FAllowMeasures", "int", tempAllowMeasures.ToString());
            parameterList.addParameter("FAllowDelete", "int", tempDelete.ToString());
            parameterList.addParameter("FAllowExcelExport", "int", tempExcelSummary.ToString());
            parameterList.addParameter("FAllowReminderOnMeasure", "int", tempAllowReminderOnMeasure.ToString());
            parameterList.addParameter("FAllowReminderOnUnit", "int", tempAllowReminderOnUnit.ToString());
            parameterList.addParameter("FAllowUpUnitsMeasures", "int", tempAllowUpUnitsMeasures.ToString());
            parameterList.addParameter("FAllowColleaguesMeasures", "int", tempAllowColleaguesMeasures.ToString());
            parameterList.addParameter("FAllowThemes", "int", tempAllowThemes.ToString());
            parameterList.addParameter("FAllowReminderOnTheme", "int", tempAllowReminderOnThemes.ToString());
            parameterList.addParameter("FAllowReminderOnThemeUnit", "int", tempAllowReminderOnThemesUnit.ToString());
            parameterList.addParameter("FAllowTopUnitsThemes", "int", tempAllowTopUnitsThemes.ToString());
            parameterList.addParameter("FAllowColleaguesThemes", "int", tempAllowColleguesThemes.ToString());
            parameterList.addParameter("FAllowOrders", "int", tempAllowOrders.ToString());
            parameterList.addParameter("FAllowReminderOnOrderUnit", "int", tempAllowReminderOnOrdersUnit.ToString());

            parameterList.addParameter("RProfile", "string", aUnitRight.RProfile);
            parameterList.addParameter("RReadLevel", "int", aUnitRight.RReadLevel.ToString());
            parameterList.addParameter("DProfile", "string", aUnitRight.DProfile);
            parameterList.addParameter("DReadLevel", "int", aUnitRight.DReadLevel.ToString());
            parameterList.addParameter("DAllowMail", "int", tempMail);
            parameterList.addParameter("DAllowDownload", "int", tempDownload);
            parameterList.addParameter("DAllowLog", "int", tempLog);
            parameterList.addParameter("DAllowType1", "int", tempFileType1);
            parameterList.addParameter("DAllowType2", "int", tempFileType2);
            parameterList.addParameter("DAllowType3", "int", tempFileType3);
            parameterList.addParameter("DAllowType4", "int", tempFileType4);
            parameterList.addParameter("DAllowType5", "int", tempFileType5);
            parameterList.addParameter("AEditLevelWorkshop", "int", aUnitRight.AEditLevelWorkshop.ToString());
            parameterList.addParameter("AReadLevelStructure", "int", aUnitRight.AReadLevelStructure.ToString());
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO mapuserorgid2 (userID, orgID, OProfile, OReadLevel, OEditLevel, OAllowDelegate, OAllowLock, OAllowEmployeeRead, OAllowEmployeeEdit, OAllowEmployeeImport, OAllowEmployeeExport, OAllowUnitAdd, OAllowUnitMove, OAllowUnitDelete, OAllowUnitProperty, OAllowReportRecipient, OAllowStructureImport, OAllowStructureExport, FProfile, FReadLevel, FEditLevel, FSumLevel, FAllowCommunications, FAllowMeasures, FAllowDelete, FAllowExcelExport, FAllowReminderOnMeasure, FAllowReminderOnUnit, FAllowUpUnitsMeasures, FAllowColleaguesMeasures, FAllowThemes, FAllowReminderOnTheme, FAllowReminderOnThemeUnit, FAllowTopUnitsThemes, FAllowColleaguesThemes, FAllowOrders, FAllowReminderOnOrderUnit, RProfile, RReadLevel, DProfile, DReadLevel, DAllowMail, DAllowDownload, DAllowLog, DAllowType1, DAllowType2, DAllowType3, DAllowType4, DAllowType5, AEditLevelWorkshop, AReadLevelStructure) VALUES (@userID, @orgID, @OProfile, @OReadLevel, @OEditLevel, @OAllowDelegate, @OAllowLock, @OAllowEmployeeRead, @OAllowEmployeeEdit, @OAllowEmployeeImport, @OAllowEmployeeExport, @OAllowUnitAdd, @OAllowUnitMove, @OAllowUnitDelete, @OAllowUnitProperty, @OAllowReportRecipient, @OAllowStructureImport, @OAllowStructureExport, @FProfile, @FReadLevel, @FEditLevel, @FSumLevel, @FAllowCommunications, @FAllowMeasures, @FAllowDelete, @FAllowExcelExport, @FAllowReminderOnMeasure, @FAllowReminderOnUnit, @FAllowUpUnitsMeasures, @FAllowColleaguesMeasures, @FAllowThemes, @FAllowReminderOnTheme, @FAllowReminderOnThemeUnit, @FAllowTopUnitsThemes, @FAllowColleaguesThemes, @FAllowOrders, @FAllowReminderOnOrderUnit, @RProfile, @RReadLevel, @DProfile, @DReadLevel, @DAllowMail, @DAllowDownload, @DAllowLog, @DAllowType1, @DAllowType2, @DAllowType3, @DAllowType4, @DAllowType5, @AEditLevelWorkshop, @AReadLevelStructure)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveUnitRightBack2(string aUserID, TBack2Right aUnitRight, string aProjectID)
    {
        SqlDB dataReader;

        // userID/orgID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("userID", "string", aUserID);
        dataReader = new SqlDB("SELECT ID FROM mapuserback2 WHERE userID=@userID AND orgID='" + aUnitRight.OrgID.ToString() + "'", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            parameterList = new TParameterList();
            parameterList.addParameter("userID", "string", aUserID);
            parameterList.addParameter("orgID", "int", aUnitRight.OrgID.ToString());
            parameterList.addParameter("OProfile", "string", aUnitRight.OProfile);
            parameterList.addParameter("OReadLevel", "int", aUnitRight.OReadLevel.ToString());
            parameterList.addParameter("AReadLevel", "int", aUnitRight.AReadLevel.ToString());
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO mapuserback2 (userID, orgID, OProfile, OReadLevel, AReadLevel) VALUES (@userID, @orgID, @OProfile, @OReadLevel, @AReadLevel)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveUnitRightOrg3(string aUserID, TOrg3Right aUnitRight, string aProjectID)
    {
        SqlDB dataReader;

        // userID/orgID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("userID", "string", aUserID);
        dataReader = new SqlDB("SELECT ID FROM mapuserorgid3 WHERE userID=@userID AND orgID='" + aUnitRight.OrgID.ToString() + "'", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            string tempDelegate = "0";
            if (aUnitRight.OAllowDelegate)
                tempDelegate = "1";
            string tempLock = "0";
            if (aUnitRight.OAllowLock)
                tempLock = "1";
            string tempEmployeeRead = "0";
            if (aUnitRight.OAllowEmployeeRead)
                tempEmployeeRead = "1";
            string tempEmployeeEdit = "0";
            if (aUnitRight.OAllowEmployeeEdit)
                tempEmployeeEdit = "1";
            string tempEmployeeImport = "0";
            if (aUnitRight.OAllowEmployeeImport)
                tempEmployeeImport = "1";
            string tempEmployeeExport = "0";
            if (aUnitRight.OAllowEmployeeExport)
                tempEmployeeExport = "1";
            string tempUnitAdd = "0";
            if (aUnitRight.OAllowUnitAdd)
                tempUnitAdd = "1";
            string tempUnitMove = "0";
            if (aUnitRight.OAllowUnitMove)
                tempUnitMove = "1";
            string tempUnitDelete = "0";
            if (aUnitRight.OAllowUnitDelete)
                tempUnitDelete = "1";
            string tempUnitProperty = "0";
            if (aUnitRight.OAllowUnitProperty)
                tempUnitProperty = "1";
            string tempReportRecipient = "0";
            if (aUnitRight.OAllowReportRecipient)
                tempReportRecipient = "1";
            string tempStructureImport = "0";
            if (aUnitRight.OAllowStructureImport)
                tempStructureImport = "1";
            string tempStructureExport = "0";
            if (aUnitRight.OAllowStructureExport)
                tempStructureExport = "1";
            string tempAllowCommunications = "0";
            if (aUnitRight.FAllowCommunications)
                tempAllowCommunications = "1";
            string tempAllowMeasures = "0";
            if (aUnitRight.FAllowMeasures)
                tempAllowMeasures = "1";
            string tempDelete = "0";
            if (aUnitRight.FAllowDelete)
                tempDelete = "1";
            string tempExcelSummary = "0";
            if (aUnitRight.FAllowExcelExport)
                tempExcelSummary = "1";
            string tempAllowReminderOnMeasure = "0";
            if (aUnitRight.FAllowReminderOnMeasure)
                tempAllowReminderOnMeasure = "1";
            string tempAllowReminderOnUnit = "0";
            if (aUnitRight.FAllowReminderOnUnit)
                tempAllowReminderOnUnit = "1";
            string tempAllowUpUnitsMeasures = "0";
            if (aUnitRight.FAllowUpUnitsMeasures)
                tempAllowUpUnitsMeasures = "1";
            string tempAllowColleaguesMeasures = "0";
            if (aUnitRight.FAllowColleaguesMeasures)
                tempAllowColleaguesMeasures = "1";

            string tempAllowThemes = "0";
            if (aUnitRight.FAllowThemes)
                tempAllowThemes = "1";
            string tempAllowReminderOnThemes = "0";
            if (aUnitRight.FAllowReminderOnThemes)
                tempAllowReminderOnThemes = "1";
            string tempAllowReminderOnThemesUnit = "0";
            if (aUnitRight.FAllowReminderOnThemeUnit)
                tempAllowReminderOnThemesUnit = "1";
            string tempAllowTopUnitsThemes = "0";
            if (aUnitRight.FAllowTopUnitsThemes)
                tempAllowTopUnitsThemes = "1";
            string tempAllowColleguesThemes = "0";
            if (aUnitRight.FAllowColleaguesThemes)
                tempAllowColleguesThemes = "1";
            string tempAllowOrders = "0";
            if (aUnitRight.FAllowOrder)
                tempAllowOrders = "1";
            string tempAllowReminderOnOrdersUnit = "0";
            if (aUnitRight.FAllowReminderOnOrderUnit)
                tempAllowReminderOnOrdersUnit = "1";

            string tempMail = "0";
            if (aUnitRight.DAllowMail)
                tempMail = "1";
            string tempDownload = "0";
            if (aUnitRight.DAllowDownload)
                tempDownload = "1";
            string tempLog = "0";
            if (aUnitRight.DAllowLog)
                tempLog = "1";
            string tempFileType1 = "0";
            if (aUnitRight.DAllowType1)
                tempFileType1 = "1";
            string tempFileType2 = "0";
            if (aUnitRight.DAllowType2)
                tempFileType2 = "1";
            string tempFileType3 = "0";
            if (aUnitRight.DAllowType3)
                tempFileType3 = "1";
            string tempFileType4 = "0";
            if (aUnitRight.DAllowType4)
                tempFileType4 = "1";
            string tempFileType5 = "0";
            if (aUnitRight.DAllowType5)
                tempFileType5 = "1";
            parameterList = new TParameterList();
            parameterList.addParameter("userID", "string", aUserID);
            parameterList.addParameter("orgID", "int", aUnitRight.OrgID.ToString());
            parameterList.addParameter("OProfile", "string", aUnitRight.OProfile);
            parameterList.addParameter("OReadLevel", "int", aUnitRight.OReadLevel.ToString());
            parameterList.addParameter("OEditLevel", "int", aUnitRight.OEditLevel.ToString());
            parameterList.addParameter("OAllowDelegate", "int", tempDelegate);
            parameterList.addParameter("OAllowLock", "int", tempLock);
            parameterList.addParameter("OAllowEmployeeRead", "int", tempEmployeeRead);
            parameterList.addParameter("OAllowEmployeeEdit", "int", tempEmployeeEdit);
            parameterList.addParameter("OAllowEmployeeImport", "int", tempEmployeeImport);
            parameterList.addParameter("OAllowEmployeeExport", "int", tempEmployeeExport);
            parameterList.addParameter("OAllowUnitAdd", "int", tempUnitAdd);
            parameterList.addParameter("OAllowUnitMove", "int", tempUnitMove);
            parameterList.addParameter("OAllowUnitDelete", "int", tempUnitDelete);
            parameterList.addParameter("OAllowUnitProperty", "int", tempUnitProperty);
            parameterList.addParameter("OAllowReportRecipient", "int", tempReportRecipient);
            parameterList.addParameter("OAllowStructureImport", "int", tempStructureImport);
            parameterList.addParameter("OAllowStructureExport", "int", tempStructureExport);
            parameterList.addParameter("FProfile", "string", aUnitRight.FProfile);
            parameterList.addParameter("FReadLevel", "int", aUnitRight.FReadLevel.ToString());
            parameterList.addParameter("FEditLevel", "int", aUnitRight.FEditLevel.ToString());
            parameterList.addParameter("FSumLevel", "int", aUnitRight.FSumLevel.ToString());
            parameterList.addParameter("FAllowCommunications", "int", tempAllowCommunications.ToString());
            parameterList.addParameter("FAllowMeasures", "int", tempAllowMeasures.ToString());
            parameterList.addParameter("FAllowDelete", "int", tempDelete.ToString());
            parameterList.addParameter("FAllowExcelExport", "int", tempExcelSummary.ToString());
            parameterList.addParameter("FAllowReminderOnMeasure", "int", tempAllowReminderOnMeasure.ToString());
            parameterList.addParameter("FAllowReminderOnUnit", "int", tempAllowReminderOnUnit.ToString());
            parameterList.addParameter("FAllowUpUnitsMeasures", "int", tempAllowUpUnitsMeasures.ToString());
            parameterList.addParameter("FAllowColleaguesMeasures", "int", tempAllowColleaguesMeasures.ToString());
            parameterList.addParameter("FAllowThemes", "int", tempAllowThemes.ToString());
            parameterList.addParameter("FAllowReminderOnTheme", "int", tempAllowReminderOnThemes.ToString());
            parameterList.addParameter("FAllowReminderOnThemeUnit", "int", tempAllowReminderOnThemesUnit.ToString());
            parameterList.addParameter("FAllowTopUnitsThemes", "int", tempAllowTopUnitsThemes.ToString());
            parameterList.addParameter("FAllowColleaguesThemes", "int", tempAllowColleguesThemes.ToString());
            parameterList.addParameter("FAllowOrders", "int", tempAllowOrders.ToString());
            parameterList.addParameter("FAllowReminderOnOrderUnit", "int", tempAllowReminderOnOrdersUnit.ToString());

            parameterList.addParameter("RProfile", "string", aUnitRight.RProfile);
            parameterList.addParameter("RReadLevel", "int", aUnitRight.RReadLevel.ToString());
            parameterList.addParameter("DProfile", "string", aUnitRight.DProfile);
            parameterList.addParameter("DReadLevel", "int", aUnitRight.DReadLevel.ToString());
            parameterList.addParameter("DAllowMail", "int", tempMail);
            parameterList.addParameter("DAllowDownload", "int", tempDownload);
            parameterList.addParameter("DAllowLog", "int", tempLog);
            parameterList.addParameter("DAllowType1", "int", tempFileType1);
            parameterList.addParameter("DAllowType2", "int", tempFileType2);
            parameterList.addParameter("DAllowType3", "int", tempFileType3);
            parameterList.addParameter("DAllowType4", "int", tempFileType4);
            parameterList.addParameter("DAllowType5", "int", tempFileType5);
            parameterList.addParameter("AEditLevelWorkshop", "int", aUnitRight.AEditLevelWorkshop.ToString());
            parameterList.addParameter("AReadLevelStructure", "int", aUnitRight.AReadLevelStructure.ToString());
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO mapuserorgid3 (userID, orgID, OProfile, OReadLevel, OEditLevel, OAllowDelegate, OAllowLock, OAllowEmployeeRead, OAllowEmployeeEdit, OAllowEmployeeImport, OAllowEmployeeExport, OAllowUnitAdd, OAllowUnitMove, OAllowUnitDelete, OAllowUnitProperty, OAllowReportRecipient, OAllowStructureImport, OAllowStructureExport, FProfile, FReadLevel, FEditLevel, FSumLevel, FAllowCommunications, FAllowMeasures, FAllowDelete, FAllowExcelExport, FAllowReminderOnMeasure, FAllowReminderOnUnit, FAllowUpUnitsMeasures, FAllowColleaguesMeasures, FAllowThemes, FAllowReminderOnTheme, FAllowReminderOnThemeUnit, FAllowTopUnitsThemes, FAllowColleaguesThemes, FAllowOrders, FAllowReminderOnOrderUnit, RProfile, RReadLevel, DProfile, DReadLevel, DAllowMail, DAllowDownload, DAllowLog, DAllowType1, DAllowType2, DAllowType3, DAllowType4, DAllowType5, AEditLevelWorkshop, AReadLevelStructure) VALUES (@userID, @orgID, @OProfile, @OReadLevel, @OEditLevel, @OAllowDelegate, @OAllowLock, @OAllowEmployeeRead, @OAllowEmployeeEdit, @OAllowEmployeeImport, @OAllowEmployeeExport, @OAllowUnitAdd, @OAllowUnitMove, @OAllowUnitDelete, @OAllowUnitProperty, @OAllowReportRecipient, @OAllowStructureImport, @OAllowStructureExport, @FProfile, @FReadLevel, @FEditLevel, @FSumLevel, @FAllowCommunications, @FAllowMeasures, @FAllowDelete, @FAllowExcelExport, @FAllowReminderOnMeasure, @FAllowReminderOnUnit, @FAllowUpUnitsMeasures, @FAllowColleaguesMeasures, @FAllowThemes, @FAllowReminderOnTheme, @FAllowReminderOnThemeUnit, @FAllowTopUnitsThemes, @FAllowColleaguesThemes, @FAllowOrders, @FAllowReminderOnOrderUnit, @RProfile, @RReadLevel, @DProfile, @DReadLevel, @DAllowMail, @DAllowDownload, @DAllowLog, @DAllowType1, @DAllowType2, @DAllowType3, @DAllowType4, @DAllowType5, @AEditLevelWorkshop, @AReadLevelStructure)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveUnitRightBack3(string aUserID, TBack3Right aUnitRight, string aProjectID)
    {
        SqlDB dataReader;

        // userID/orgID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("userID", "string", aUserID);
        dataReader = new SqlDB("SELECT ID FROM mapuserback3 WHERE userID=@userID AND orgID='" + aUnitRight.OrgID.ToString() + "'", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            parameterList = new TParameterList();
            parameterList.addParameter("userID", "string", aUserID);
            parameterList.addParameter("orgID", "int", aUnitRight.OrgID.ToString());
            parameterList.addParameter("OProfile", "string", aUnitRight.OProfile);
            parameterList.addParameter("OReadLevel", "int", aUnitRight.OReadLevel.ToString());
            parameterList.addParameter("AReadLevel", "int", aUnitRight.AReadLevel.ToString());
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO mapuserback3 (userID, orgID, OProfile, OReadLevel, AReadLevel) VALUES (@userID, @orgID, @OProfile, @OReadLevel, @AReadLevel)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveUserRight(string aUserID, TUserRight aTaskContactRight, string aProjectID)
    {
        SqlDB dataReader;
        // userID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("userID", "string", aUserID);
        dataReader = new SqlDB("SELECT ID FROM mapuserrights WHERE userID=@userID", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            string tempTaskRead = "0";
            if (aTaskContactRight.TAllowTaskRead)
                tempTaskRead = "1";
            string tempTaskCreateEditDelete = "0";
            if (aTaskContactRight.TAllowTaskCreateEditDelete)
                tempTaskCreateEditDelete = "1";
            string tempTaskFileUpload = "0";
            if (aTaskContactRight.TAllowTaskFileUpload)
                tempTaskFileUpload = "1";
            string tempTaskExport = "0";
            if (aTaskContactRight.TAllowTaskExport)
                tempTaskExport = "1";
            string tempTaskImport = "0";
            if (aTaskContactRight.TAllowTaskImport)
                tempTaskImport = "1";
            string tempContactRead = "0";
            if (aTaskContactRight.TAllowContactRead)
                tempContactRead = "1";
            string tempContactCreatEditDelete = "0";
            if (aTaskContactRight.TAllowContactCreateEditDelete)
                tempContactCreatEditDelete = "1";
            string tempContactExport = "0";
            if (aTaskContactRight.TAllowContactExport)
                tempContactExport = "1";
            string tempContactImport = "0";
            if (aTaskContactRight.TAllowContactImport)
                tempContactImport = "1";
            string tempTranslateRead = "0";
            if (aTaskContactRight.TAllowTranslateRead)
                tempTranslateRead = "1";
            string tempTranslateEdit = "0";
            if (aTaskContactRight.TAllowTranslateEdit)
                tempTranslateEdit = "1";
            string tempTranslateRelease = "0";
            if (aTaskContactRight.TAllowTranslateRelease)
                tempTranslateRelease = "1";
            string tempTranslateNew = "0";
            if (aTaskContactRight.TAllowTranslateNew)
                tempTranslateNew = "1";
            string tempTranslateDelete = "0";
            if (aTaskContactRight.TAllowTranslateDelete)
                tempTranslateDelete = "1";
            string tempAccountTransfer = "0";
            if (aTaskContactRight.AAllowAccountTransfer)
                tempAccountTransfer = "1";
            string tempAccountMangement = "0";
            if (aTaskContactRight.AAllowAccountManagement)
                tempAccountMangement = "1";

            parameterList = new TParameterList();
            parameterList.addParameter("userID", "string", aUserID);
            parameterList.addParameter("TTaskProfile", "string", aTaskContactRight.TTaskProfile);
            parameterList.addParameter("AllowTaskRead", "int", tempTaskRead);
            parameterList.addParameter("AllowTaskCreateEditDelete", "int", tempTaskCreateEditDelete);
            parameterList.addParameter("AllowTaskFileUpload", "int", tempTaskFileUpload);
            parameterList.addParameter("AllowTaskExport", "int", tempTaskExport);
            parameterList.addParameter("AllowTaskImport", "int", tempTaskImport);
            parameterList.addParameter("TContactProfile", "string", aTaskContactRight.TContactProfile);
            parameterList.addParameter("AllowContactRead", "int", tempContactRead);
            parameterList.addParameter("AllowContactCreateEditDelete", "int", tempContactCreatEditDelete);
            parameterList.addParameter("AllowContactExport", "int", tempContactExport);
            parameterList.addParameter("AllowContactImport", "int", tempContactImport);
            parameterList.addParameter("TTranslateProfile", "string", aTaskContactRight.TTranslateProfile);
            parameterList.addParameter("AllowTranslateRead", "int", tempTranslateRead);
            parameterList.addParameter("AllowTranslateEdit", "int", tempTranslateEdit);
            parameterList.addParameter("AllowTranslateRelease", "int", tempTranslateRelease);
            parameterList.addParameter("AllowTranslateNew", "int", tempTranslateNew);
            parameterList.addParameter("AllowTranslateDelete", "int", tempTranslateDelete);
            parameterList.addParameter("AllowAccountTransfer", "int", tempAccountTransfer);
            parameterList.addParameter("AllowAccountManagement", "int", tempAccountMangement);

            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO mapuserrights (userID, TTaskProfile, AllowTaskRead, AllowTaskCreateEditDelete, AllowTaskFileUpload, AllowTaskExport, AllowTaskImport, TContactProfile, AllowContactRead, AllowContactCreateEditDelete, AllowContactExport, AllowContactImport, TTranslateProfile, AllowTranslateRead, AllowTranslateEdit, AllowTranslateRelease, AllowTranslateNew, AllowTranslateDelete, AllowAccountTransfer, AllowAccountManagement) VALUES (@userID, @TTaskProfile, @AllowTaskRead, @AllowTaskCreateEditDelete, @AllowTaskFileUpload, @AllowTaskExport, @AllowTaskImport, @TContactProfile, @AllowContactRead, @AllowContactCreateEditDelete, @AllowContactExport, @AllowContactImport, @TTranslateProfile, @AllowTranslateRead, @AllowTranslateEdit, @AllowTranslateRelease, @AllowTranslateNew, @AllowTranslateDelete, @AllowAccountTransfer, @AllowAccountManagement)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveTaskRoleRight(string aUserID, int aTaskRoleID, string aProjectID)
    {
        SqlDB dataReader;

        // userID/roleID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("userID", "string", aUserID);
        dataReader = new SqlDB("SELECT ID FROM mapusertaskrole WHERE userID=@userID AND roleID='" + aTaskRoleID.ToString() + "'", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            parameterList = new TParameterList();
            parameterList.addParameter("userID", "string", aUserID);
            parameterList.addParameter("roleID", "int", aTaskRoleID.ToString());
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO mapusertaskrole (userID, roleID) VALUES (@userID, @roleID)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveGroupRight(string aUserID, string aGroupID, string aProjectID)
    {
        SqlDB dataReader;

        // userID/groupID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("userID", "string", aUserID);
        dataReader = new SqlDB("SELECT ID FROM mapusergroup WHERE userID=@userID AND groupID='" + aGroupID.ToString() + "'", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            parameterList = new TParameterList();
            parameterList.addParameter("userID", "string", aUserID);
            parameterList.addParameter("groupID", "string", aGroupID);
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO mapusergroup (userID, groupID) VALUES (@userID, @groupID)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveTranslateLanguageRight(string aUserID, int aTranslateLanguageID, string aProjectID)
    {
        SqlDB dataReader;

        // userID/tranalatelangauageID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("userID", "string", aUserID);
        dataReader = new SqlDB("SELECT ID FROM mapusertranslatelanguage WHERE userID=@userID AND translateLanguageID='" + aTranslateLanguageID.ToString() + "'", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            parameterList = new TParameterList();
            parameterList.addParameter("userID", "string", aUserID);
            parameterList.addParameter("translateLanguageID", "int", aTranslateLanguageID.ToString());
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO mapusertranslatelanguage (userID, translateLanguageID) VALUES (@userID, @translateLanguageID)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveContainerRight(string aUserID, TContainerRight aContainerRight, string aProjectID)
    {
        SqlDB dataReader;

        // userID/orgID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("userID", "string", aUserID);
        dataReader = new SqlDB("SELECT ID FROM mapusercontainer WHERE userID=@userID AND containerID='" + aContainerRight.ContainerID.ToString() + "'", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            string tempUpload = "0";
            if (aContainerRight.TAllowUpload)
                tempUpload = "1";
            string tempDownload = "0";
            if (aContainerRight.TAllowDownload)
                tempDownload = "1";
            string tempDeleteOwn = "0";
            if (aContainerRight.TAllowDeleteOwnFiles)
                tempDeleteOwn = "1";
            string tempDeleteAll = "0";
            if (aContainerRight.TAllowDeleteAllFiles)
                tempDeleteAll = "1";
            string tempAccessOwnFilesWithoutPassword = "0";
            if (aContainerRight.TAllowAccessOwnFilesWithoutPassword)
                tempAccessOwnFilesWithoutPassword = "1";
            string tempAccessAllFilesWithoutPassword = "0";
            if (aContainerRight.TAllowAccessAllFilesWithoutPassword)
                tempAccessAllFilesWithoutPassword = "1";
            string tempCreateFolder = "0";
            if (aContainerRight.TAllowCreateFolder)
                tempCreateFolder = "1";
            string tempDeleteOwnFolder = "0";
            if (aContainerRight.TAllowDeleteOwnFolder)
                tempDeleteOwnFolder = "1";
            string tempDeleteAllFolder = "0";
            if (aContainerRight.TAllowDeleteAllFolder)
                tempDeleteAllFolder = "1";
            string tempRestPassword = "0";
            if (aContainerRight.TAllowResetPassword)
                tempRestPassword = "1";
            string tempTakeOwnership = "0";
            if (aContainerRight.TAllowTakeOwnership)
                tempTakeOwnership = "1";
            parameterList = new TParameterList();
            parameterList.addParameter("userID", "string", aUserID);
            parameterList.addParameter("containerID", "int", aContainerRight.ContainerID.ToString());
            parameterList.addParameter("TContainerProfile", "string", aContainerRight.TContainerProfile);
            parameterList.addParameter("TAllowUpload", "int", tempUpload);
            parameterList.addParameter("TAllowDownload", "int", tempDownload);
            parameterList.addParameter("TAllowDeleteOwn", "int", tempDeleteOwn);
            parameterList.addParameter("TAllowDeleteAll", "int", tempDeleteAll);
            parameterList.addParameter("TAllowAccessOwnFilesWithoutPassword", "int", tempAccessOwnFilesWithoutPassword);
            parameterList.addParameter("TAllowAccessAllFilesWithoutPassword", "int", tempAccessAllFilesWithoutPassword);
            parameterList.addParameter("TAllowCreateFolder", "int", tempCreateFolder);
            parameterList.addParameter("TAllowDeleteOwnFolder", "int", tempDeleteOwnFolder);
            parameterList.addParameter("TAllowDeleteAllFolder", "int", tempDeleteAllFolder);
            parameterList.addParameter("TAllowResetPassword", "int", tempRestPassword);
            parameterList.addParameter("TAllowTakeOwnership", "int", tempTakeOwnership);
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO mapusercontainer (userID, containerID, TContainerProfile, allowUpload, allowDownload, allowDeleteOwnFiles, allowDeleteAllFiles, allowAccessOwnFilesWithoutPassword, allowAccessAllFilesWithoutPassword, allowCreateFolder, allowDeleteOwnFolder, allowDeleteAllFolder, allowResetPassword, allowTakeOwnership) VALUES (@userID, @containerID, @TContainerProfile, @TAllowUpload, @TAllowDownload, @TAllowDeleteOwn, @TAllowDeleteAll, @TAllowAccessOwnFilesWithoutPassword, @TAllowAccessAllFilesWithoutPassword, @TAllowCreateFolder, @TAllowDeleteOwnFolder, @TAllowDeleteAllFolder, @TAllowResetPassword, @TAllowTakeOwnership)", parameterList);

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
        string tempSystemGeneratedPassword = "0";
        if (SystemGeneratedPassword)
            tempSystemGeneratedPassword = "1";
        string tempInternalUser = "0";
        if (IsInternalUser)
            tempInternalUser = "1";
        string tempForce = "0";
        if (ForceChange)
            tempForce = "1";
        string tempActive = "0";
        if (Active)
            tempActive = "1";

        TCrypt cryptKey = new TCrypt();
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("displayName", "string", DisplayName);
        parameterList.addParameter("password", "string", Password);
        parameterList.addParameter("email", "string", Email);
        parameterList.addParameter("title", "string", Title);
        parameterList.addParameter("name", "string", Name);
        parameterList.addParameter("firstname", "string", Firstname);
        parameterList.addParameter("street", "string", Street);
        parameterList.addParameter("zipcode", "string", Zipcode);
        parameterList.addParameter("city", "string", City);
        parameterList.addParameter("state", "string", State);
        parameterList.addParameter("lastlogin", "datetime", LastLogin.ToString());
        parameterList.addParameter("lastloginerror", "datetime", LastLoginError.ToString());
        parameterList.addParameter("lastloginerrorcount", "int", LastLoginErrorCount.ToString());
        parameterList.addParameter("isinternaluser", "int", tempInternalUser);
        parameterList.addParameter("forcechange", "int", tempForce);
        parameterList.addParameter("active", "int", tempActive);
        parameterList.addParameter("systemgeneratedpassword", "int", tempSystemGeneratedPassword);
        parameterList.addParameter("passwordValidDate", "datetime", PasswordValidDate.ToString());
        parameterList.addParameter("language", "int", Language.ToString());
        parameterList.addParameter("comment", "string", Comment);
        parameterList.addParameter("groupID", "string", GroupID);
        parameterList.addParameter("userID", "string", UserID);
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("UPDATE loginuser SET displayName=@displayName, password=@password, email=@email, title=@title, name=@name, firstname=@firstname, street=@street, zipcode=@zipcode, city=@city, state=@state, lastlogin=@lastlogin, lastloginerror=@lastloginerror, lastloginerrorcount=@lastloginerrorcount, isinternaluser=@isinternaluser, forcechange=@forcechange, active=@active, systemgeneratedpassword=@systemgeneratedpassword, passwordValidDate=@passwordValidDate, language=@language, comment=@comment, groupID=@groupID WHERE userID=@userID", parameterList);
        // bisherige Tool-Berechtigungen löschen
        parameterList = new TParameterList();
        parameterList.addParameter("userID", "string", UserID);
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("DELETE FROM mapusertool WHERE userID=@userID", parameterList);
        // neue Tool-Berechtigungen eintragen
        foreach (TToolRight tempTool in toolRights)
        {
            parameterList = new TParameterList();
            parameterList.addParameter("userID", "string", UserID);
            parameterList.addParameter("toolID", "string", tempTool.toolID);
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO mapusertool (userID, toolID) VALUES (@userID, @toolID)", parameterList);
        }
    }
    public void updateUserRight(TOrgRight aUnitRight, string aProjectID)
    {
        string tempDelegate = "0";
        if (aUnitRight.OAllowDelegate)
            tempDelegate = "1";
        string tempLock = "0";
        if (aUnitRight.OAllowLock)
            tempLock = "1";
        string tempEmployeeRead = "0";
        if (aUnitRight.OAllowEmployeeRead)
            tempEmployeeRead = "1";
        string tempEmployeeEdit = "0";
        if (aUnitRight.OAllowEmployeeEdit)
            tempEmployeeEdit = "1";
        string tempEmployeeImport = "0";
        if (aUnitRight.OAllowEmployeeImport)
            tempEmployeeImport = "1";
        string tempEmployeeExport = "0";
        if (aUnitRight.OAllowEmployeeExport)
            tempEmployeeExport = "1";
        string tempUnitAdd = "0";
        if (aUnitRight.OAllowUnitAdd)
            tempUnitAdd = "1";
        string tempUnitMove = "0";
        if (aUnitRight.OAllowUnitMove)
            tempUnitMove = "1";
        string tempUnitDelete = "0";
        if (aUnitRight.OAllowUnitDelete)
            tempUnitDelete = "1";
        string tempUnitProperty = "0";
        if (aUnitRight.OAllowUnitProperty)
            tempUnitProperty = "1";
        string tempReportRecipient = "0";
        if (aUnitRight.OAllowReportRecipient)
            tempReportRecipient = "1";
        string tempStructureImport = "0";
        if (aUnitRight.OAllowStructureImport)
            tempStructureImport = "1";
        string tempStructureExport = "0";
        if (aUnitRight.OAllowStructureExport)
            tempStructureExport = "1";
        string tempBouncesRead = "0";
        if (aUnitRight.OAllowBouncesRead)
            tempBouncesRead = "1";
        string tempBouncesEdit = "0";
        if (aUnitRight.OAllowBouncesEdit)
            tempBouncesEdit = "1";
        string tempBouncesDelete = "0";
        if (aUnitRight.OAllowBouncesDelete)
            tempBouncesDelete = "1";
        string tempBouncesExport = "0";
        if (aUnitRight.OAllowBouncesExport)
            tempBouncesExport = "1";
        string tempBouncesImport = "0";
        if (aUnitRight.OAllowBouncesImport)
            tempBouncesImport = "1";
        string tempAllowReInvitation = "0";
        if (aUnitRight.OAllowReInvitation)
            tempAllowReInvitation = "1";
        string tempAllowPostNomination = "0";
        if (aUnitRight.OAllowPostNomination)
            tempAllowPostNomination = "1";
        string tempAllowPostNominationImport = "0";
        if (aUnitRight.OAllowPostNominationImport)
            tempAllowPostNominationImport = "1";
        string tempAllowCommunications = "0";
        if (aUnitRight.FAllowCommunications)
            tempAllowCommunications = "1";
        string tempAllowMeasures = "0";
        if (aUnitRight.FAllowMeasures)
            tempAllowMeasures = "1";
        string tempDelete = "0";
        if (aUnitRight.FAllowDelete)
            tempDelete = "1";
        string tempExcelSummary = "0";
        if (aUnitRight.FAllowExcelExport)
            tempExcelSummary = "1";
        string tempAllowReminderOnMeasure = "0";
        if (aUnitRight.FAllowReminderOnMeasure)
            tempAllowReminderOnMeasure = "1";
        string tempAllowReminderOnUnit = "0";
        if (aUnitRight.FAllowReminderOnUnit)
            tempAllowReminderOnUnit = "1";
        string tempAllowUpUnitsMeasures = "0";
        if (aUnitRight.FAllowUpUnitsMeasures)
            tempAllowUpUnitsMeasures = "1";
        string tempAllowColleaguesMeasures = "0";
        if (aUnitRight.FAllowColleaguesMeasures)
            tempAllowColleaguesMeasures = "1";

        string tempAllowThemes = "0";
        if (aUnitRight.FAllowThemes)
            tempAllowThemes = "1";
        string tempAllowReminderOnThemes = "0";
        if (aUnitRight.FAllowReminderOnThemes)
            tempAllowReminderOnThemes = "1";
        string tempAllowReminderOnThemesUnit = "0";
        if (aUnitRight.FAllowReminderOnThemeUnit)
            tempAllowReminderOnThemesUnit = "1";
        string tempAllowTopUnitsThemes = "0";
        if (aUnitRight.FAllowTopUnitsThemes)
            tempAllowTopUnitsThemes = "1";
        string tempAllowColleguesThemes = "0";
        if (aUnitRight.FAllowColleaguesThemes)
            tempAllowColleguesThemes = "1";
        string tempAllowOrders = "0";
        if (aUnitRight.FAllowOrder)
            tempAllowOrders = "1";
        string tempAllowReminderOnOrdersUnit = "0";
        if (aUnitRight.FAllowReminderOnOrderUnit)
            tempAllowReminderOnOrdersUnit = "1";

        string tempMail = "0";
        if (aUnitRight.DAllowMail)
            tempMail = "1";
        string tempDownload = "0";
        if (aUnitRight.DAllowDownload)
            tempDownload = "1";
        string tempLog = "0";
        if (aUnitRight.DAllowLog)
            tempLog = "1";
        string tempFileType1 = "0";
        if (aUnitRight.DAllowType1)
            tempFileType1 = "1";
        string tempFileType2 = "0";
        if (aUnitRight.DAllowType2)
            tempFileType2 = "1";
        string tempFileType3 = "0";
        if (aUnitRight.DAllowType3)
            tempFileType3 = "1";
        string tempFileType4 = "0";
        if (aUnitRight.DAllowType4)
            tempFileType4 = "1";
        string tempFileType5 = "0";
        if (aUnitRight.DAllowType5)
            tempFileType5 = "1";

        SqlDB dataReader;
        TParameterList parameterList = new TParameterList();
        parameterList = new TParameterList();
        parameterList.addParameter("ID", "int", aUnitRight.ID.ToString());
        parameterList.addParameter("userID", "string", UserID);
        parameterList.addParameter("orgID", "int", aUnitRight.OrgID.ToString());
        parameterList.addParameter("OProfile", "string", aUnitRight.OProfile);
        parameterList.addParameter("OReadLevel", "int", aUnitRight.OReadLevel.ToString());
        parameterList.addParameter("OEditLevel", "int", aUnitRight.OEditLevel.ToString());
        parameterList.addParameter("OAllowDelegate", "int", tempDelegate);
        parameterList.addParameter("OAllowLock", "int", tempLock);
        parameterList.addParameter("OAllowEmployeeRead", "int", tempEmployeeRead);
        parameterList.addParameter("OAllowEmployeeEdit", "int", tempEmployeeEdit);
        parameterList.addParameter("OAllowEmployeeImport", "int", tempEmployeeImport);
        parameterList.addParameter("OAllowEmployeeExport", "int", tempEmployeeExport);
        parameterList.addParameter("OAllowUnitAdd", "int", tempUnitAdd);
        parameterList.addParameter("OAllowUnitMove", "int", tempUnitMove);
        parameterList.addParameter("OAllowUnitDelete", "int", tempUnitDelete);
        parameterList.addParameter("OAllowUnitProperty", "int", tempUnitProperty);
        parameterList.addParameter("OAllowReportRecipient", "int", tempReportRecipient);
        parameterList.addParameter("OAllowStructureImport", "int", tempStructureImport);
        parameterList.addParameter("OAllowStructureExport", "int", tempStructureExport);
        parameterList.addParameter("OAllowBouncesRead", "int", tempBouncesRead);
        parameterList.addParameter("OAllowBouncesEdit", "int", tempBouncesEdit);
        parameterList.addParameter("OAllowBouncesDelete", "int", tempBouncesDelete);
        parameterList.addParameter("OAllowBouncesExport", "int", tempBouncesExport);
        parameterList.addParameter("OAllowBouncesImport", "int", tempBouncesImport);
        parameterList.addParameter("OAllowReInvitation", "int", tempAllowReInvitation);
        parameterList.addParameter("OAllowPostNomination", "int", tempAllowPostNomination);
        parameterList.addParameter("OAllowPostNominationImport", "int", tempAllowPostNominationImport);
        parameterList.addParameter("FProfile", "string", aUnitRight.FProfile);
        parameterList.addParameter("FReadLevel", "int", aUnitRight.FReadLevel.ToString());
        parameterList.addParameter("FEditLevel", "int", aUnitRight.FEditLevel.ToString());
        parameterList.addParameter("FSumLevel", "int", aUnitRight.FSumLevel.ToString());
        parameterList.addParameter("FAllowCommunications", "int", tempAllowCommunications.ToString());
        parameterList.addParameter("FAllowMeasures", "int", tempAllowMeasures.ToString());
        parameterList.addParameter("FAllowDelete", "int", tempDelete.ToString());
        parameterList.addParameter("FAllowExcelExport", "int", tempExcelSummary.ToString());
        parameterList.addParameter("FAllowReminderOnMeasure", "int", tempAllowReminderOnMeasure.ToString());
        parameterList.addParameter("FAllowReminderOnUnit", "int", tempAllowReminderOnUnit.ToString());
        parameterList.addParameter("FAllowUpUnitsMeasures", "int", tempAllowUpUnitsMeasures.ToString());
        parameterList.addParameter("FAllowColleaguesMeasures", "int", tempAllowColleaguesMeasures.ToString());
        parameterList.addParameter("FAllowThemes", "int", tempAllowThemes.ToString());
        parameterList.addParameter("FAllowReminderOnTheme", "int", tempAllowReminderOnThemes.ToString());
        parameterList.addParameter("FAllowReminderOnThemeUnit", "int", tempAllowReminderOnThemesUnit.ToString());
        parameterList.addParameter("FAllowTopUnitsThemes", "int", tempAllowTopUnitsThemes.ToString());
        parameterList.addParameter("FAllowColleaguesThemes", "int", tempAllowColleguesThemes.ToString());
        parameterList.addParameter("FAllowOrders", "int", tempAllowOrders.ToString());
        parameterList.addParameter("FAllowReminderOnOrderUnit", "int", tempAllowReminderOnOrdersUnit.ToString());
        parameterList.addParameter("RProfile", "string", aUnitRight.RProfile);
        parameterList.addParameter("RReadLevel", "int", aUnitRight.RReadLevel.ToString());
        parameterList.addParameter("DProfile", "string", aUnitRight.DProfile);
        parameterList.addParameter("DReadLevel", "int", aUnitRight.DReadLevel.ToString());
        parameterList.addParameter("DAllowMail", "int", tempMail);
        parameterList.addParameter("DAllowDownload", "int", tempDownload);
        parameterList.addParameter("DAllowLog", "int", tempLog);
        parameterList.addParameter("DAllowType1", "int", tempFileType1);
        parameterList.addParameter("DAllowType2", "int", tempFileType2);
        parameterList.addParameter("DAllowType3", "int", tempFileType3);
        parameterList.addParameter("DAllowType4", "int", tempFileType4);
        parameterList.addParameter("DAllowType5", "int", tempFileType5);
        parameterList.addParameter("AEditLevelWorkshop", "int", aUnitRight.AEditLevelWorkshop.ToString());
        parameterList.addParameter("AReadLevelStructure", "int", aUnitRight.AReadLevelStructure.ToString());
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("UPDATE mapuserorgid SET OProfile=@OProfile, OReadLevel=@OReadLevel, OEditLevel=@OEditLevel, OAllowDelegate=@OAllowDelegate, OAllowLock=@OAllowLock, OAllowEmployeeRead=@OAllowEmployeeRead, OAllowEmployeeEdit=@OAllowEmployeeEdit, OAllowEmployeeImport=@OAllowEmployeeImport, OAllowEmployeeExport=@OAllowEmployeeExport, OAllowUnitAdd=@OAllowUnitAdd, OAllowUnitMove=@OAllowUnitMove, OAllowUnitDelete=@OAllowUnitDelete, OAllowUnitProperty=@OAllowUnitProperty, OAllowReportRecipient=@OAllowReportRecipient, OAllowStructureImport=@OAllowStructureImport, OAllowStructureExport=@OAllowStructureExport, OAllowBouncesRead=@OAllowBouncesRead, OAllowBouncesEdit=@OAllowBouncesEdit, OAllowBouncesDelete=@OAllowBouncesDelete, OAllowBouncesExport=@OAllowBouncesExport, OAllowBouncesImport=@OAllowBouncesImport, OAllowReInvitation=@OAllowReInvitation, OAllowPostNomination=@OAllowPostNomination, OAllowPostNominationImport=@OAllowPostNominationImport, FProfile=@FProfile, FReadLevel=@FReadLevel, FEditLevel=@FEditLevel, FSumLevel=@FSumLevel, FAllowCommunications=@FAllowCommunications, FAllowMeasures=@FAllowMeasures, FAllowDelete=@FAllowDelete, FAllowExcelExport=@FAllowExcelExport, FAllowReminderOnMeasure=@FAllowReminderOnMeasure, FAllowReminderOnUnit=@FAllowReminderOnUnit, FAllowUpUnitsMeasures=@FAllowUpUnitsMeasures, FAllowColleaguesMeasure=@FAllowColleaguesMeasure, FAllowThemes=@FAllowThemes, FAllowReminderOnTheme=@FAllowReminderOnTheme, FAllowReminderOnThemeUnit=@FAllowReminderOnThemeUnit, FAllowTopUnitsThemes=@FAllowTopUnitsThemes, FAllowColleaguesThemes=@FAllowColleaguesThemes, FAllowOrders=@FAllowOrders, FAllowReminderOnOrderUnit=@FAllowReminderOnOrderUnit, RProfile=@RProfile, RReadLevel=@RReadLevel, DProfile=@DProfile, DReadLevel=@DReadLevel, DAllowMail=@DAllowMail, DAllowDownload=@DAllowDownload, DAllowLog=@DAllowLog, DAllowType1=@DAllowType1, DAllowType2=@DAllowType2, DAllowType3=@DAllowType3, DAllowType4=@DAllowType4, DAllowType5=@DAllowType5, AEditLevelWorkshop=@AEditLevelWorkshop, AReadLevelStructure=@AReadLevelStructure WHERE ID=@ID", parameterList);
    }
    public void updateUserRightBack(TBackRight aUnitRight, string aProjectID)
    {
        SqlDB dataReader;
        TParameterList parameterList = new TParameterList();
        parameterList = new TParameterList();
        parameterList.addParameter("ID", "int", aUnitRight.ID.ToString());
        parameterList.addParameter("userID", "string", UserID);
        parameterList.addParameter("OProfile", "string", aUnitRight.OProfile);
        parameterList.addParameter("orgID", "int", aUnitRight.OrgID.ToString());
        parameterList.addParameter("OReadLevel", "int", aUnitRight.OReadLevel.ToString());
        parameterList.addParameter("AReadLevel", "int", aUnitRight.AReadLevel.ToString());
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("UPDATE mapuserback SET OProfile=@OProfile, OReadLevel=@OReadLevel, AReadLevel=@AReadLevel WHERE ID=@ID", parameterList);
    }
    public void updateUserRightOrg1(TOrg1Right aUnitRight, string aProjectID)
    {
        string tempDelegate = "0";
        if (aUnitRight.OAllowDelegate)
            tempDelegate = "1";
        string tempLock = "0";
        if (aUnitRight.OAllowLock)
            tempLock = "1";
        string tempEmployeeRead = "0";
        if (aUnitRight.OAllowEmployeeRead)
            tempEmployeeRead = "1";
        string tempEmployeeEdit = "0";
        if (aUnitRight.OAllowEmployeeEdit)
            tempEmployeeEdit = "1";
        string tempEmployeeImport = "0";
        if (aUnitRight.OAllowEmployeeImport)
            tempEmployeeImport = "1";
        string tempEmployeeExport = "0";
        if (aUnitRight.OAllowEmployeeExport)
            tempEmployeeExport = "1";
        string tempUnitAdd = "0";
        if (aUnitRight.OAllowUnitAdd)
            tempUnitAdd = "1";
        string tempUnitMove = "0";
        if (aUnitRight.OAllowUnitMove)
            tempUnitMove = "1";
        string tempUnitDelete = "0";
        if (aUnitRight.OAllowUnitDelete)
            tempUnitDelete = "1";
        string tempUnitProperty = "0";
        if (aUnitRight.OAllowUnitProperty)
            tempUnitProperty = "1";
        string tempReportRecipient = "0";
        if (aUnitRight.OAllowReportRecipient)
            tempReportRecipient = "1";
        string tempStructureImport = "0";
        if (aUnitRight.OAllowStructureImport)
            tempStructureImport = "1";
        string tempStructureExport = "0";
        if (aUnitRight.OAllowStructureExport)
            tempStructureExport = "1";
        string tempAllowCommunications = "0";
        if (aUnitRight.FAllowCommunications)
            tempAllowCommunications = "1";
        string tempAllowMeasures = "0";
        if (aUnitRight.FAllowMeasures)
            tempAllowMeasures = "1";
        string tempDelete = "0";
        if (aUnitRight.FAllowDelete)
            tempDelete = "1";
        string tempExcelSummary = "0";
        if (aUnitRight.FAllowExcelExport)
            tempExcelSummary = "1";
        string tempAllowReminderOnMeasure = "0";
        if (aUnitRight.FAllowReminderOnMeasure)
            tempAllowReminderOnMeasure = "1";
        string tempAllowReminderOnUnit = "0";
        if (aUnitRight.FAllowReminderOnUnit)
            tempAllowReminderOnUnit = "1";
        string tempAllowUpUnitsMeasures = "0";
        if (aUnitRight.FAllowUpUnitsMeasures)
            tempAllowUpUnitsMeasures = "1";
        string tempAllowColleaguesMeasures = "0";
        if (aUnitRight.FAllowColleaguesMeasures)
            tempAllowColleaguesMeasures = "1";

        string tempAllowThemes = "0";
        if (aUnitRight.FAllowThemes)
            tempAllowThemes = "1";
        string tempAllowReminderOnThemes = "0";
        if (aUnitRight.FAllowReminderOnThemes)
            tempAllowReminderOnThemes = "1";
        string tempAllowReminderOnThemesUnit = "0";
        if (aUnitRight.FAllowReminderOnThemeUnit)
            tempAllowReminderOnThemesUnit = "1";
        string tempAllowTopUnitsThemes = "0";
        if (aUnitRight.FAllowTopUnitsThemes)
            tempAllowTopUnitsThemes = "1";
        string tempAllowColleguesThemes = "0";
        if (aUnitRight.FAllowColleaguesThemes)
            tempAllowColleguesThemes = "1";
        string tempAllowOrders = "0";
        if (aUnitRight.FAllowOrder)
            tempAllowOrders = "1";
        string tempAllowReminderOnOrdersUnit = "0";
        if (aUnitRight.FAllowReminderOnOrderUnit)
            tempAllowReminderOnOrdersUnit = "1";

        string tempMail = "0";
        if (aUnitRight.DAllowMail)
            tempMail = "1";
        string tempDownload = "0";
        if (aUnitRight.DAllowDownload)
            tempDownload = "1";
        string tempLog = "0";
        if (aUnitRight.DAllowLog)
            tempLog = "1";
        string tempFileType1 = "0";
        if (aUnitRight.DAllowType1)
            tempFileType1 = "1";
        string tempFileType2 = "0";
        if (aUnitRight.DAllowType2)
            tempFileType2 = "1";
        string tempFileType3 = "0";
        if (aUnitRight.DAllowType3)
            tempFileType3 = "1";
        string tempFileType4 = "0";
        if (aUnitRight.DAllowType4)
            tempFileType4 = "1";
        string tempFileType5 = "0";
        if (aUnitRight.DAllowType5)
            tempFileType5 = "1";

        SqlDB dataReader;
        TParameterList parameterList = new TParameterList();
        parameterList = new TParameterList();
        parameterList.addParameter("ID", "int", aUnitRight.ID.ToString());
        parameterList.addParameter("userID", "string", UserID);
        parameterList.addParameter("orgID", "int", aUnitRight.OrgID.ToString());
        parameterList.addParameter("OProfile", "string", aUnitRight.OProfile);
        parameterList.addParameter("OReadLevel", "int", aUnitRight.OReadLevel.ToString());
        parameterList.addParameter("OEditLevel", "int", aUnitRight.OEditLevel.ToString());
        parameterList.addParameter("OAllowDelegate", "int", tempDelegate);
        parameterList.addParameter("OAllowLock", "int", tempLock);
        parameterList.addParameter("OAllowEmployeeRead", "int", tempEmployeeRead);
        parameterList.addParameter("OAllowEmployeeEdit", "int", tempEmployeeEdit);
        parameterList.addParameter("OAllowEmployeeImport", "int", tempEmployeeImport);
        parameterList.addParameter("OAllowEmployeeExport", "int", tempEmployeeExport);
        parameterList.addParameter("OAllowUnitAdd", "int", tempUnitAdd);
        parameterList.addParameter("OAllowUnitMove", "int", tempUnitMove);
        parameterList.addParameter("OAllowUnitDelete", "int", tempUnitDelete);
        parameterList.addParameter("OAllowUnitProperty", "int", tempUnitProperty);
        parameterList.addParameter("OAllowReportRecipient", "int", tempReportRecipient);
        parameterList.addParameter("OAllowStructureImport", "int", tempStructureImport);
        parameterList.addParameter("OAllowStructureExport", "int", tempStructureExport);
        parameterList.addParameter("FProfile", "string", aUnitRight.FProfile);
        parameterList.addParameter("FReadLevel", "int", aUnitRight.FReadLevel.ToString());
        parameterList.addParameter("FEditLevel", "int", aUnitRight.FEditLevel.ToString());
        parameterList.addParameter("FSumLevel", "int", aUnitRight.FSumLevel.ToString());
        parameterList.addParameter("FAllowCommunications", "int", tempAllowCommunications.ToString());
        parameterList.addParameter("FAllowMeasures", "int", tempAllowMeasures.ToString());
        parameterList.addParameter("FAllowDelete", "int", tempDelete.ToString());
        parameterList.addParameter("FAllowExcelExport", "int", tempExcelSummary.ToString());
        parameterList.addParameter("FAllowReminderOnMeasure", "int", tempAllowReminderOnMeasure.ToString());
        parameterList.addParameter("FAllowReminderOnUnit", "int", tempAllowReminderOnUnit.ToString());
        parameterList.addParameter("FAllowUpUnitsMeasures", "int", tempAllowUpUnitsMeasures.ToString());
        parameterList.addParameter("FAllowColleaguesMeasures", "int", tempAllowColleaguesMeasures.ToString());
        parameterList.addParameter("FAllowThemes", "int", tempAllowThemes.ToString());
        parameterList.addParameter("FAllowReminderOnTheme", "int", tempAllowReminderOnThemes.ToString());
        parameterList.addParameter("FAllowReminderOnThemeUnit", "int", tempAllowReminderOnThemesUnit.ToString());
        parameterList.addParameter("FAllowTopUnitsThemes", "int", tempAllowTopUnitsThemes.ToString());
        parameterList.addParameter("FAllowColleaguesThemes", "int", tempAllowColleguesThemes.ToString());
        parameterList.addParameter("FAllowOrders", "int", tempAllowOrders.ToString());
        parameterList.addParameter("FAllowReminderOnOrderUnit", "int", tempAllowReminderOnOrdersUnit.ToString());
        parameterList.addParameter("RProfile", "string", aUnitRight.RProfile);
        parameterList.addParameter("RReadLevel", "int", aUnitRight.RReadLevel.ToString());
        parameterList.addParameter("DProfile", "string", aUnitRight.DProfile);
        parameterList.addParameter("DReadLevel", "int", aUnitRight.DReadLevel.ToString());
        parameterList.addParameter("DAllowMail", "int", tempMail);
        parameterList.addParameter("DAllowDownload", "int", tempDownload);
        parameterList.addParameter("DAllowLog", "int", tempLog);
        parameterList.addParameter("DAllowType1", "int", tempFileType1);
        parameterList.addParameter("DAllowType2", "int", tempFileType2);
        parameterList.addParameter("DAllowType3", "int", tempFileType3);
        parameterList.addParameter("DAllowType4", "int", tempFileType4);
        parameterList.addParameter("DAllowType5", "int", tempFileType5);
        parameterList.addParameter("AEditLevelWorkshop", "int", aUnitRight.AEditLevelWorkshop.ToString());
        parameterList.addParameter("AReadLevelStructure", "int", aUnitRight.AReadLevelStructure.ToString());
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("UPDATE mapuserorgid1 SET OProfile=@OProfile, OReadLevel=@OReadLevel, OEditLevel=@OEditLevel, OAllowDelegate=@OAllowDelegate, OAllowLock=@OAllowLock, OAllowEmployeeRead=@OAllowEmployeeRead, OAllowEmployeeEdit=@OAllowEmployeeEdit, OAllowEmployeeImport=@OAllowEmployeeImport, OAllowEmployeeExport=@OAllowEmployeeExport, OAllowUnitAdd=@OAllowUnitAdd, OAllowUnitMove=@OAllowUnitMove, OAllowUnitDelete=@OAllowUnitDelete, OAllowUnitProperty=@OAllowUnitProperty, OAllowReportRecipient=@OAllowReportRecipient, OAllowStructureImport=@OAllowStructureImport, OAllowStructureExport=@OAllowStructureExport, FProfile=@FProfile, FReadLevel=@FReadLevel, FEditLevel=@FEditLevel, FSumLevel=@FSumLevel, FAllowCommunications=@FAllowCommunications, FAllowMeasures=@FAllowMeasures, FAllowDelete=@FAllowDelete, FAllowExcelExport=@FAllowExcelExport, FAllowReminderOnMeasure=@FAllowReminderOnMeasure, FAllowReminderOnUnit=@FAllowReminderOnUnit, FAllowUpUnitsMeasures=@FAllowUpUnitsMeasures, FAllowColleaguesMeasure=@FAllowColleaguesMeasure, FAllowThemes=@FAllowThemes, FAllowReminderOnTheme=@FAllowReminderOnTheme, FAllowReminderOnThemeUnit=@FAllowReminderOnThemeUnit, FAllowTopUnitsThemes=@FAllowTopUnitsThemes, FAllowColleaguesThemes=@FAllowColleaguesThemes, FAllowOrders=@FAllowOrders, FAllowReminderOnOrderUnit=@FAllowReminderOnOrderUnit, RProfile=@RProfile, RReadLevel=@RReadLevel, DProfile=@DProfile, DReadLevel=@DReadLevel, DAllowMail=@DAllowMail, DAllowDownload=@DAllowDownload, DAllowLog=@DAllowLog, DAllowType1=@DAllowType1, DAllowType2=@DAllowType2, DAllowType3=@DAllowType3, DAllowType4=@DAllowType4, DAllowType5=@DAllowType5, AEditLevelWorkshop=@AEditLevelWorkshop, AReadLevelStructure=@AReadLevelStructure WHERE ID=@ID", parameterList);
    }
    public void updateUserRightBack1(TBack1Right aUnitRight, string aProjectID)
    {
        SqlDB dataReader;
        TParameterList parameterList = new TParameterList();
        parameterList = new TParameterList();
        parameterList.addParameter("ID", "int", aUnitRight.ID.ToString());
        parameterList.addParameter("userID", "string", UserID);
        parameterList.addParameter("OProfile", "string", aUnitRight.OProfile);
        parameterList.addParameter("orgID", "int", aUnitRight.OrgID.ToString());
        parameterList.addParameter("OReadLevel", "int", aUnitRight.OReadLevel.ToString());
        parameterList.addParameter("AReadLevel", "int", aUnitRight.AReadLevel.ToString());
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("UPDATE mapuserback1 SET OProfile=@OProfile, OReadLevel=@OReadLevel, AReadLevel=@AReadLevel WHERE ID=@ID", parameterList);
    }
    public void updateUserRightOrg2(TOrg2Right aUnitRight, string aProjectID)
    {
        string tempDelegate = "0";
        if (aUnitRight.OAllowDelegate)
            tempDelegate = "1";
        string tempLock = "0";
        if (aUnitRight.OAllowLock)
            tempLock = "1";
        string tempEmployeeRead = "0";
        if (aUnitRight.OAllowEmployeeRead)
            tempEmployeeRead = "1";
        string tempEmployeeEdit = "0";
        if (aUnitRight.OAllowEmployeeEdit)
            tempEmployeeEdit = "1";
        string tempEmployeeImport = "0";
        if (aUnitRight.OAllowEmployeeImport)
            tempEmployeeImport = "1";
        string tempEmployeeExport = "0";
        if (aUnitRight.OAllowEmployeeExport)
            tempEmployeeExport = "1";
        string tempUnitAdd = "0";
        if (aUnitRight.OAllowUnitAdd)
            tempUnitAdd = "1";
        string tempUnitMove = "0";
        if (aUnitRight.OAllowUnitMove)
            tempUnitMove = "1";
        string tempUnitDelete = "0";
        if (aUnitRight.OAllowUnitDelete)
            tempUnitDelete = "1";
        string tempUnitProperty = "0";
        if (aUnitRight.OAllowUnitProperty)
            tempUnitProperty = "1";
        string tempReportRecipient = "0";
        if (aUnitRight.OAllowReportRecipient)
            tempReportRecipient = "1";
        string tempStructureImport = "0";
        if (aUnitRight.OAllowStructureImport)
            tempStructureImport = "1";
        string tempStructureExport = "0";
        if (aUnitRight.OAllowStructureExport)
            tempStructureExport = "1";
        string tempAllowCommunications = "0";
        if (aUnitRight.FAllowCommunications)
            tempAllowCommunications = "1";
        string tempAllowMeasures = "0";
        if (aUnitRight.FAllowMeasures)
            tempAllowMeasures = "1";
        string tempDelete = "0";
        if (aUnitRight.FAllowDelete)
            tempDelete = "1";
        string tempExcelSummary = "0";
        if (aUnitRight.FAllowExcelExport)
            tempExcelSummary = "1";
        string tempAllowReminderOnMeasure = "0";
        if (aUnitRight.FAllowReminderOnMeasure)
            tempAllowReminderOnMeasure = "1";
        string tempAllowReminderOnUnit = "0";
        if (aUnitRight.FAllowReminderOnUnit)
            tempAllowReminderOnUnit = "1";
        string tempAllowUpUnitsMeasures = "0";
        if (aUnitRight.FAllowUpUnitsMeasures)
            tempAllowUpUnitsMeasures = "1";
        string tempAllowColleaguesMeasures = "0";
        if (aUnitRight.FAllowColleaguesMeasures)
            tempAllowColleaguesMeasures = "1";

        string tempAllowThemes = "0";
        if (aUnitRight.FAllowThemes)
            tempAllowThemes = "1";
        string tempAllowReminderOnThemes = "0";
        if (aUnitRight.FAllowReminderOnThemes)
            tempAllowReminderOnThemes = "1";
        string tempAllowReminderOnThemesUnit = "0";
        if (aUnitRight.FAllowReminderOnThemeUnit)
            tempAllowReminderOnThemesUnit = "1";
        string tempAllowTopUnitsThemes = "0";
        if (aUnitRight.FAllowTopUnitsThemes)
            tempAllowTopUnitsThemes = "1";
        string tempAllowColleguesThemes = "0";
        if (aUnitRight.FAllowColleaguesThemes)
            tempAllowColleguesThemes = "1";
        string tempAllowOrders = "0";
        if (aUnitRight.FAllowOrder)
            tempAllowOrders = "1";
        string tempAllowReminderOnOrdersUnit = "0";
        if (aUnitRight.FAllowReminderOnOrderUnit)
            tempAllowReminderOnOrdersUnit = "1";

        string tempMail = "0";
        if (aUnitRight.DAllowMail)
            tempMail = "1";
        string tempDownload = "0";
        if (aUnitRight.DAllowDownload)
            tempDownload = "1";
        string tempLog = "0";
        if (aUnitRight.DAllowLog)
            tempLog = "1";
        string tempFileType1 = "0";
        if (aUnitRight.DAllowType1)
            tempFileType1 = "1";
        string tempFileType2 = "0";
        if (aUnitRight.DAllowType2)
            tempFileType2 = "1";
        string tempFileType3 = "0";
        if (aUnitRight.DAllowType3)
            tempFileType3 = "1";
        string tempFileType4 = "0";
        if (aUnitRight.DAllowType4)
            tempFileType4 = "1";
        string tempFileType5 = "0";
        if (aUnitRight.DAllowType5)
            tempFileType5 = "1";

        SqlDB dataReader;
        TParameterList parameterList = new TParameterList();
        parameterList = new TParameterList();
        parameterList.addParameter("ID", "int", aUnitRight.ID.ToString());
        parameterList.addParameter("userID", "string", UserID);
        parameterList.addParameter("orgID", "int", aUnitRight.OrgID.ToString());
        parameterList.addParameter("OProfile", "string", aUnitRight.OProfile);
        parameterList.addParameter("OReadLevel", "int", aUnitRight.OReadLevel.ToString());
        parameterList.addParameter("OEditLevel", "int", aUnitRight.OEditLevel.ToString());
        parameterList.addParameter("OAllowDelegate", "int", tempDelegate);
        parameterList.addParameter("OAllowLock", "int", tempLock);
        parameterList.addParameter("OAllowEmployeeRead", "int", tempEmployeeRead);
        parameterList.addParameter("OAllowEmployeeEdit", "int", tempEmployeeEdit);
        parameterList.addParameter("OAllowEmployeeImport", "int", tempEmployeeImport);
        parameterList.addParameter("OAllowEmployeeExport", "int", tempEmployeeExport);
        parameterList.addParameter("OAllowUnitAdd", "int", tempUnitAdd);
        parameterList.addParameter("OAllowUnitMove", "int", tempUnitMove);
        parameterList.addParameter("OAllowUnitDelete", "int", tempUnitDelete);
        parameterList.addParameter("OAllowUnitProperty", "int", tempUnitProperty);
        parameterList.addParameter("OAllowReportRecipient", "int", tempReportRecipient);
        parameterList.addParameter("OAllowStructureImport", "int", tempStructureImport);
        parameterList.addParameter("OAllowStructureExport", "int", tempStructureExport);
        parameterList.addParameter("FProfile", "string", aUnitRight.FProfile);
        parameterList.addParameter("FReadLevel", "int", aUnitRight.FReadLevel.ToString());
        parameterList.addParameter("FEditLevel", "int", aUnitRight.FEditLevel.ToString());
        parameterList.addParameter("FSumLevel", "int", aUnitRight.FSumLevel.ToString());
        parameterList.addParameter("FAllowCommunications", "int", tempAllowCommunications.ToString());
        parameterList.addParameter("FAllowMeasures", "int", tempAllowMeasures.ToString());
        parameterList.addParameter("FAllowDelete", "int", tempDelete.ToString());
        parameterList.addParameter("FAllowExcelExport", "int", tempExcelSummary.ToString());
        parameterList.addParameter("FAllowReminderOnMeasure", "int", tempAllowReminderOnMeasure.ToString());
        parameterList.addParameter("FAllowReminderOnUnit", "int", tempAllowReminderOnUnit.ToString());
        parameterList.addParameter("FAllowUpUnitsMeasures", "int", tempAllowUpUnitsMeasures.ToString());
        parameterList.addParameter("FAllowColleaguesMeasures", "int", tempAllowColleaguesMeasures.ToString());
        parameterList.addParameter("FAllowThemes", "int", tempAllowThemes.ToString());
        parameterList.addParameter("FAllowReminderOnTheme", "int", tempAllowReminderOnThemes.ToString());
        parameterList.addParameter("FAllowReminderOnThemeUnit", "int", tempAllowReminderOnThemesUnit.ToString());
        parameterList.addParameter("FAllowTopUnitsThemes", "int", tempAllowTopUnitsThemes.ToString());
        parameterList.addParameter("FAllowColleaguesThemes", "int", tempAllowColleguesThemes.ToString());
        parameterList.addParameter("FAllowOrders", "int", tempAllowOrders.ToString());
        parameterList.addParameter("FAllowReminderOnOrderUnit", "int", tempAllowReminderOnOrdersUnit.ToString());
        parameterList.addParameter("RProfile", "string", aUnitRight.RProfile);
        parameterList.addParameter("RReadLevel", "int", aUnitRight.RReadLevel.ToString());
        parameterList.addParameter("DProfile", "string", aUnitRight.DProfile);
        parameterList.addParameter("DReadLevel", "int", aUnitRight.DReadLevel.ToString());
        parameterList.addParameter("DAllowMail", "int", tempMail);
        parameterList.addParameter("DAllowDownload", "int", tempDownload);
        parameterList.addParameter("DAllowLog", "int", tempLog);
        parameterList.addParameter("DAllowType1", "int", tempFileType1);
        parameterList.addParameter("DAllowType2", "int", tempFileType2);
        parameterList.addParameter("DAllowType3", "int", tempFileType3);
        parameterList.addParameter("DAllowType4", "int", tempFileType4);
        parameterList.addParameter("DAllowType5", "int", tempFileType5);
        parameterList.addParameter("AEditLevelWorkshop", "int", aUnitRight.AEditLevelWorkshop.ToString());
        parameterList.addParameter("AReadLevelStructure", "int", aUnitRight.AReadLevelStructure.ToString());
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("UPDATE mapuserorgid2 SET OProfile=@OProfile, OReadLevel=@OReadLevel, OEditLevel=@OEditLevel, OAllowDelegate=@OAllowDelegate, OAllowLock=@OAllowLock, OAllowEmployeeRead=@OAllowEmployeeRead, OAllowEmployeeEdit=@OAllowEmployeeEdit, OAllowEmployeeImport=@OAllowEmployeeImport, OAllowEmployeeExport=@OAllowEmployeeExport, OAllowUnitAdd=@OAllowUnitAdd, OAllowUnitMove=@OAllowUnitMove, OAllowUnitDelete=@OAllowUnitDelete, OAllowUnitProperty=@OAllowUnitProperty, OAllowReportRecipient=@OAllowReportRecipient, OAllowStructureImport=@OAllowStructureImport, OAllowStructureExport=@OAllowStructureExport, FProfile=@FProfile, FReadLevel=@FReadLevel, FEditLevel=@FEditLevel, FSumLevel=@FSumLevel, FAllowCommunications=@FAllowCommunications, FAllowMeasures=@FAllowMeasures, FAllowDelete=@FAllowDelete, FAllowExcelExport=@FAllowExcelExport, FAllowReminderOnMeasure=@FAllowReminderOnMeasure, FAllowReminderOnUnit=@FAllowReminderOnUnit, FAllowUpUnitsMeasures=@FAllowUpUnitsMeasures, FAllowColleaguesMeasure=@FAllowColleaguesMeasure, FAllowThemes=@FAllowThemes, FAllowReminderOnTheme=@FAllowReminderOnTheme, FAllowReminderOnThemeUnit=@FAllowReminderOnThemeUnit, FAllowTopUnitsThemes=@FAllowTopUnitsThemes, FAllowColleaguesThemes=@FAllowColleaguesThemes, FAllowOrders=@FAllowOrders, FAllowReminderOnOrderUnit=@FAllowReminderOnOrderUnit, RProfile=@RProfile, RReadLevel=@RReadLevel, DProfile=@DProfile, DReadLevel=@DReadLevel, DAllowMail=@DAllowMail, DAllowDownload=@DAllowDownload, DAllowLog=@DAllowLog, DAllowType1=@DAllowType1, DAllowType2=@DAllowType2, DAllowType3=@DAllowType3, DAllowType4=@DAllowType4, DAllowType5=@DAllowType5, AEditLevelWorkshop=@AEditLevelWorkshop, AReadLevelStructure=@AReadLevelStructure WHERE ID=@ID", parameterList);
    }
    public void updateUserRightBack2(TBack2Right aUnitRight, string aProjectID)
    {
        SqlDB dataReader;
        TParameterList parameterList = new TParameterList();
        parameterList = new TParameterList();
        parameterList.addParameter("ID", "int", aUnitRight.ID.ToString());
        parameterList.addParameter("userID", "string", UserID);
        parameterList.addParameter("orgID", "int", aUnitRight.OrgID.ToString());
        parameterList.addParameter("OProfile", "string", aUnitRight.OProfile);
        parameterList.addParameter("OReadLevel", "int", aUnitRight.OReadLevel.ToString());
        parameterList.addParameter("AReadLevel", "int", aUnitRight.AReadLevel.ToString());
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("UPDATE mapuserback2 SET OProfile=@OProfile, OReadLevel=@OReadLevel, AReadLevel=@AReadLevel WHERE ID=@ID", parameterList);
    }
    public void updateUserRightOrg3(TOrg3Right aUnitRight, string aProjectID)
    {
        string tempDelegate = "0";
        if (aUnitRight.OAllowDelegate)
            tempDelegate = "1";
        string tempLock = "0";
        if (aUnitRight.OAllowLock)
            tempLock = "1";
        string tempEmployeeRead = "0";
        if (aUnitRight.OAllowEmployeeRead)
            tempEmployeeRead = "1";
        string tempEmployeeEdit = "0";
        if (aUnitRight.OAllowEmployeeEdit)
            tempEmployeeEdit = "1";
        string tempEmployeeImport = "0";
        if (aUnitRight.OAllowEmployeeImport)
            tempEmployeeImport = "1";
        string tempEmployeeExport = "0";
        if (aUnitRight.OAllowEmployeeExport)
            tempEmployeeExport = "1";
        string tempUnitAdd = "0";
        if (aUnitRight.OAllowUnitAdd)
            tempUnitAdd = "1";
        string tempUnitMove = "0";
        if (aUnitRight.OAllowUnitMove)
            tempUnitMove = "1";
        string tempUnitDelete = "0";
        if (aUnitRight.OAllowUnitDelete)
            tempUnitDelete = "1";
        string tempUnitProperty = "0";
        if (aUnitRight.OAllowUnitProperty)
            tempUnitProperty = "1";
        string tempReportRecipient = "0";
        if (aUnitRight.OAllowReportRecipient)
            tempReportRecipient = "1";
        string tempStructureImport = "0";
        if (aUnitRight.OAllowStructureImport)
            tempStructureImport = "1";
        string tempStructureExport = "0";
        if (aUnitRight.OAllowStructureExport)
            tempStructureExport = "1";
        string tempAllowCommunications = "0";
        if (aUnitRight.FAllowCommunications)
            tempAllowCommunications = "1";
        string tempAllowMeasures = "0";
        if (aUnitRight.FAllowMeasures)
            tempAllowMeasures = "1";
        string tempDelete = "0";
        if (aUnitRight.FAllowDelete)
            tempDelete = "1";
        string tempExcelSummary = "0";
        if (aUnitRight.FAllowExcelExport)
            tempExcelSummary = "1";
        string tempAllowReminderOnMeasure = "0";
        if (aUnitRight.FAllowReminderOnMeasure)
            tempAllowReminderOnMeasure = "1";
        string tempAllowReminderOnUnit = "0";
        if (aUnitRight.FAllowReminderOnUnit)
            tempAllowReminderOnUnit = "1";
        string tempAllowUpUnitsMeasures = "0";
        if (aUnitRight.FAllowUpUnitsMeasures)
            tempAllowUpUnitsMeasures = "1";
        string tempAllowColleaguesMeasures = "0";
        if (aUnitRight.FAllowColleaguesMeasures)
            tempAllowColleaguesMeasures = "1";

        string tempAllowThemes = "0";
        if (aUnitRight.FAllowThemes)
            tempAllowThemes = "1";
        string tempAllowReminderOnThemes = "0";
        if (aUnitRight.FAllowReminderOnThemes)
            tempAllowReminderOnThemes = "1";
        string tempAllowReminderOnThemesUnit = "0";
        if (aUnitRight.FAllowReminderOnThemeUnit)
            tempAllowReminderOnThemesUnit = "1";
        string tempAllowTopUnitsThemes = "0";
        if (aUnitRight.FAllowTopUnitsThemes)
            tempAllowTopUnitsThemes = "1";
        string tempAllowColleguesThemes = "0";
        if (aUnitRight.FAllowColleaguesThemes)
            tempAllowColleguesThemes = "1";
        string tempAllowOrders = "0";
        if (aUnitRight.FAllowOrder)
            tempAllowOrders = "1";
        string tempAllowReminderOnOrdersUnit = "0";
        if (aUnitRight.FAllowReminderOnOrderUnit)
            tempAllowReminderOnOrdersUnit = "1";

        string tempMail = "0";
        if (aUnitRight.DAllowMail)
            tempMail = "1";
        string tempDownload = "0";
        if (aUnitRight.DAllowDownload)
            tempDownload = "1";
        string tempLog = "0";
        if (aUnitRight.DAllowLog)
            tempLog = "1";
        string tempFileType1 = "0";
        if (aUnitRight.DAllowType1)
            tempFileType1 = "1";
        string tempFileType2 = "0";
        if (aUnitRight.DAllowType2)
            tempFileType2 = "1";
        string tempFileType3 = "0";
        if (aUnitRight.DAllowType3)
            tempFileType3 = "1";
        string tempFileType4 = "0";
        if (aUnitRight.DAllowType4)
            tempFileType4 = "1";
        string tempFileType5 = "0";
        if (aUnitRight.DAllowType5)
            tempFileType5 = "1";

        SqlDB dataReader;
        TParameterList parameterList = new TParameterList();
        parameterList = new TParameterList();
        parameterList.addParameter("ID", "int", aUnitRight.ID.ToString());
        parameterList.addParameter("userID", "string", UserID);
        parameterList.addParameter("orgID", "int", aUnitRight.OrgID.ToString());
        parameterList.addParameter("OProfile", "string", aUnitRight.OProfile);
        parameterList.addParameter("OReadLevel", "int", aUnitRight.OReadLevel.ToString());
        parameterList.addParameter("OEditLevel", "int", aUnitRight.OEditLevel.ToString());
        parameterList.addParameter("OAllowDelegate", "int", tempDelegate);
        parameterList.addParameter("OAllowLock", "int", tempLock);
        parameterList.addParameter("OAllowEmployeeRead", "int", tempEmployeeRead);
        parameterList.addParameter("OAllowEmployeeEdit", "int", tempEmployeeEdit);
        parameterList.addParameter("OAllowEmployeeImport", "int", tempEmployeeImport);
        parameterList.addParameter("OAllowEmployeeExport", "int", tempEmployeeExport);
        parameterList.addParameter("OAllowUnitAdd", "int", tempUnitAdd);
        parameterList.addParameter("OAllowUnitMove", "int", tempUnitMove);
        parameterList.addParameter("OAllowUnitDelete", "int", tempUnitDelete);
        parameterList.addParameter("OAllowUnitProperty", "int", tempUnitProperty);
        parameterList.addParameter("OAllowReportRecipient", "int", tempReportRecipient);
        parameterList.addParameter("OAllowStructureImport", "int", tempStructureImport);
        parameterList.addParameter("OAllowStructureExport", "int", tempStructureExport);
        parameterList.addParameter("FProfile", "string", aUnitRight.FProfile);
        parameterList.addParameter("FReadLevel", "int", aUnitRight.FReadLevel.ToString());
        parameterList.addParameter("FEditLevel", "int", aUnitRight.FEditLevel.ToString());
        parameterList.addParameter("FSumLevel", "int", aUnitRight.FSumLevel.ToString());
        parameterList.addParameter("FAllowCommunications", "int", tempAllowCommunications.ToString());
        parameterList.addParameter("FAllowMeasures", "int", tempAllowMeasures.ToString());
        parameterList.addParameter("FAllowDelete", "int", tempDelete.ToString());
        parameterList.addParameter("FAllowExcelExport", "int", tempExcelSummary.ToString());
        parameterList.addParameter("FAllowReminderOnMeasure", "int", tempAllowReminderOnMeasure.ToString());
        parameterList.addParameter("FAllowReminderOnUnit", "int", tempAllowReminderOnUnit.ToString());
        parameterList.addParameter("FAllowUpUnitsMeasures", "int", tempAllowUpUnitsMeasures.ToString());
        parameterList.addParameter("FAllowColleaguesMeasures", "int", tempAllowColleaguesMeasures.ToString());
        parameterList.addParameter("FAllowThemes", "int", tempAllowThemes.ToString());
        parameterList.addParameter("FAllowReminderOnTheme", "int", tempAllowReminderOnThemes.ToString());
        parameterList.addParameter("FAllowReminderOnThemeUnit", "int", tempAllowReminderOnThemesUnit.ToString());
        parameterList.addParameter("FAllowTopUnitsThemes", "int", tempAllowTopUnitsThemes.ToString());
        parameterList.addParameter("FAllowColleaguesThemes", "int", tempAllowColleguesThemes.ToString());
        parameterList.addParameter("FAllowOrders", "int", tempAllowOrders.ToString());
        parameterList.addParameter("FAllowReminderOnOrderUnit", "int", tempAllowReminderOnOrdersUnit.ToString());
        parameterList.addParameter("RProfile", "string", aUnitRight.RProfile);
        parameterList.addParameter("RReadLevel", "int", aUnitRight.RReadLevel.ToString());
        parameterList.addParameter("DProfile", "string", aUnitRight.DProfile);
        parameterList.addParameter("DReadLevel", "int", aUnitRight.DReadLevel.ToString());
        parameterList.addParameter("DAllowMail", "int", tempMail);
        parameterList.addParameter("DAllowDownload", "int", tempDownload);
        parameterList.addParameter("DAllowLog", "int", tempLog);
        parameterList.addParameter("DAllowType1", "int", tempFileType1);
        parameterList.addParameter("DAllowType2", "int", tempFileType2);
        parameterList.addParameter("DAllowType3", "int", tempFileType3);
        parameterList.addParameter("DAllowType4", "int", tempFileType4);
        parameterList.addParameter("DAllowType5", "int", tempFileType5);
        parameterList.addParameter("AEditLevelWorkshop", "int", aUnitRight.AEditLevelWorkshop.ToString());
        parameterList.addParameter("AReadLevelStructure", "int", aUnitRight.AReadLevelStructure.ToString());
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("UPDATE mapuserorgid3 SET OProfile=@OProfile, OReadLevel=@OReadLevel, OEditLevel=@OEditLevel, OAllowDelegate=@OAllowDelegate, OAllowLock=@OAllowLock, OAllowEmployeeRead=@OAllowEmployeeRead, OAllowEmployeeEdit=@OAllowEmployeeEdit, OAllowEmployeeImport=@OAllowEmployeeImport, OAllowEmployeeExport=@OAllowEmployeeExport, OAllowUnitAdd=@OAllowUnitAdd, OAllowUnitMove=@OAllowUnitMove, OAllowUnitDelete=@OAllowUnitDelete, OAllowUnitProperty=@OAllowUnitProperty, OAllowReportRecipient=@OAllowReportRecipient, OAllowStructureImport=@OAllowStructureImport, OAllowStructureExport=@OAllowStructureExport, FProfile=@FProfile, FReadLevel=@FReadLevel, FEditLevel=@FEditLevel, FSumLevel=@FSumLevel, FAllowCommunications=@FAllowCommunications, FAllowMeasures=@FAllowMeasures, FAllowDelete=@FAllowDelete, FAllowExcelExport=@FAllowExcelExport, FAllowReminderOnMeasure=@FAllowReminderOnMeasure, FAllowReminderOnUnit=@FAllowReminderOnUnit, FAllowUpUnitsMeasures=@FAllowUpUnitsMeasures, FAllowColleaguesMeasure=@FAllowColleaguesMeasure, FAllowThemes=@FAllowThemes, FAllowReminderOnTheme=@FAllowReminderOnTheme, FAllowReminderOnThemeUnit=@FAllowReminderOnThemeUnit, FAllowTopUnitsThemes=@FAllowTopUnitsThemes, FAllowColleaguesThemes=@FAllowColleaguesThemes, FAllowOrders=@FAllowOrders, FAllowReminderOnOrderUnit=@FAllowReminderOnOrderUnit, RProfile=@RProfile, RReadLevel=@RReadLevel, DProfile=@DProfile, DReadLevel=@DReadLevel, DAllowMail=@DAllowMail, DAllowDownload=@DAllowDownload, DAllowLog=@DAllowLog, DAllowType1=@DAllowType1, DAllowType2=@DAllowType2, DAllowType3=@DAllowType3, DAllowType4=@DAllowType4, DAllowType5=@DAllowType5, AEditLevelWorkshop=@AEditLevelWorkshop, AReadLevelStructure=@AReadLevelStructure WHERE ID=@ID", parameterList);
    }
    public void updateUserRightBack3(TBack3Right aUnitRight, string aProjectID)
    {
        SqlDB dataReader;
        TParameterList parameterList = new TParameterList();
        parameterList = new TParameterList();
        parameterList.addParameter("ID", "int", aUnitRight.ID.ToString());
        parameterList.addParameter("userID", "string", UserID);
        parameterList.addParameter("orgID", "int", aUnitRight.OrgID.ToString());
        parameterList.addParameter("OProfile", "string", aUnitRight.OProfile);
        parameterList.addParameter("OReadLevel", "int", aUnitRight.OReadLevel.ToString());
        parameterList.addParameter("AReadLevel", "int", aUnitRight.AReadLevel.ToString());
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("UPDATE mapuserback3 SET OProfile=@OProfile, OReadLevel=@OReadLevel, AReadLevel=@AReadLevel WHERE ID=@ID", parameterList);
    }
    public void updateContainerRight(TContainerRight aContainerRight, string aProjectID)
    {
        string tempUpload = "0";
        if (aContainerRight.TAllowUpload)
            tempUpload = "1";
        string tempDownload = "0";
        if (aContainerRight.TAllowDownload)
            tempDownload = "1";
        string tempDeleteOwn = "0";
        if (aContainerRight.TAllowDeleteOwnFiles)
            tempDeleteOwn = "1";
        string tempDeleteAll = "0";
        if (aContainerRight.TAllowDeleteAllFiles)
            tempDeleteAll = "1";
        string tempAccessOwnFilesWithoutPassword = "0";
        if (aContainerRight.TAllowAccessOwnFilesWithoutPassword)
            tempAccessOwnFilesWithoutPassword = "1";
        string tempAccessAllFilesWithoutPassword = "0";
        if (aContainerRight.TAllowAccessAllFilesWithoutPassword)
            tempAccessAllFilesWithoutPassword = "1";
        string tempCreateFolder = "0";
        if (aContainerRight.TAllowCreateFolder)
            tempCreateFolder = "1";
        string tempDeleteOwnFolder = "0";
        if (aContainerRight.TAllowDeleteOwnFolder)
            tempDeleteOwnFolder = "1";
        string tempDeleteAllFolder = "0";
        if (aContainerRight.TAllowDeleteAllFolder)
            tempDeleteAllFolder = "1";
        string tempRestPassword = "0";
        if (aContainerRight.TAllowResetPassword)
            tempRestPassword = "1";
        string tempTakeOwnership = "0";
        if (aContainerRight.TAllowTakeOwnership)
            tempTakeOwnership = "1";
        SqlDB dataReader;
        TParameterList parameterList = new TParameterList();
        parameterList = new TParameterList();
        parameterList.addParameter("userID", "int", aContainerRight.ID.ToString());
        parameterList.addParameter("TAllowUpload", "int", tempUpload);
        parameterList.addParameter("TAllowDownload", "int", tempDownload);
        parameterList.addParameter("TAllowDeleteOwn", "int", tempDeleteOwn);
        parameterList.addParameter("TAllowDeleteAll", "int", tempDeleteAll);
        parameterList.addParameter("TAllowAccessOwnFilesWithputPassword", "int", tempAccessOwnFilesWithoutPassword);
        parameterList.addParameter("TAllowAccessAllFilesWithoutPassword", "int", tempAccessAllFilesWithoutPassword);
        parameterList.addParameter("TAllowCreateFolder", "int", tempCreateFolder);
        parameterList.addParameter("TAllowDeleteOwnFolder", "int", tempDeleteOwnFolder);
        parameterList.addParameter("TAllowDeleteAllFolder", "int", tempDeleteAllFolder);
        parameterList.addParameter("TAllowResetPassword", "int", tempRestPassword);
        parameterList.addParameter("TAllowTakeOwnership", "int", tempTakeOwnership);
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("UPDATE mapusercontainer SET allowUpload=@TAllowUpload, allowDownload=@TAllowDownload, allowDeleteOwnFiles=@TAllowDeleteOwn, allowDeleteAllFiles=@TAllowDeleteAll, allowAccessOwnFilesWithoutPassword=@TAllowAccessOwnFilesWithputPassword, allowAccessAllFilesWithoutPassword=@TAllowAccessAllFilesWithoutPassword, allowCreateFolder=@TAllowCreateFolder, allowDeleteOwnFolder=@TAllowDeleteOwnFolder, allowDeleteAllFolder=@TAllowDeleteAllFolder, allowResetPassword=@TAllowResetPassword, allowTakeOwnership=@TAllowTakeOwnership WHERE ID=@ID", parameterList);
    }
    /// <summary>
    /// Löschen eines Benutzers aus der DB
    /// </summary>
    /// <param name="aUserID">ID des Benutzers</param>
    /// <param name="aProjectID">ID des Projektes</param>
    public static void delete(string aUserID, string aProjectID)
    {
        SqlDB dataReader;
        // aus Datenbank löschen
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("userID", "string", aUserID);
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("DELETE FROM loginuser WHERE userID=@userID", parameterList);
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("DELETE FROM mapuserorgid WHERE userID=@userID", parameterList);
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("DELETE FROM mapuserback WHERE userID=@userID", parameterList);
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("DELETE FROM mapuserorgid1 WHERE userID=@userID", parameterList);
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("DELETE FROM mapuserback1 WHERE userID=@userID", parameterList);
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("DELETE FROM mapuserorgid2 WHERE userID=@userID", parameterList);
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("DELETE FROM mapuserback2 WHERE userID=@userID", parameterList);
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("DELETE FROM mapuserorgid3 WHERE userID=@userID", parameterList);
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("DELETE FROM mapuserback3 WHERE userID=@userID", parameterList);
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("DELETE FROM mapusercontainer WHERE userID=@userID", parameterList);
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("DELETE FROM mapuserrights WHERE userID=@userID", parameterList);
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("DELETE FROM mapusertaskrole WHERE userID=@userID", parameterList);
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("DELETE FROM mapusergroup WHERE userID=@userID", parameterList);
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("DELETE FROM mapuserhidestartpage WHERE userID=@userID", parameterList);
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("DELETE FROM mapusertranslatelanguage WHERE userID=@userID", parameterList);
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("DELETE FROM mapusertool WHERE userID=@userID", parameterList);
    }
    public static void deleteRight(int aID, string orgTable, string aProjectID)
    {
        SqlDB dataReader;
        // aus Datenbank löschen
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQL("DELETE FROM " + orgTable + " WHERE ID='" + aID + "'");
    }
    public void sendMail(string aServer, string aUser, string aPassword, bool aForceChange, bool aSystemGeneratedPassword, int aValidDays, string aPath, string aProjectID, string aMailTemplate, TSessionObj aSessionObj)
    {
        //TSessionObj SessionObj = (TSessionObj)System.Web.HttpContext.Current.Session["SessionObj"];

        // Maildatei öffnen
        string FileName = aPath + aProjectID + "\\Admin\\InviteMails\\" + aMailTemplate;
        if (File.Exists(FileName))
        {
            StringCollection mailText = new StringCollection();
            using (StreamReader sr = new StreamReader(FileName, Encoding.UTF8))
            {
                while (!sr.EndOfStream)
                {
                    mailText.Add(sr.ReadLine());
                }
            }
            // ermitteln, ob Kennwort im Mail vorhanden -> neues Kennwort generieren
            bool generateNewPassword = false;
            for (int lineIndex = 2; lineIndex < mailText.Count; lineIndex++)
            {
                if (mailText[lineIndex].IndexOf("#Password#") >= 0)
                    generateNewPassword = true;
            }
            string tempPassword = "";
            if (generateNewPassword)
            {
                // neues Kennwort erzeugen und speichern
                tempPassword = aSessionObj.Project.userPassword.generatePassword(UserID, aProjectID);
                //tempPassword = Login.generatePassword(this.UserID, aSessionObj.Project.passwordLength, aSessionObj.Project.passwordComplexity, aSessionObj.Project.passwortHistory, aSessionObj.Project.ProjectID);
                Password = TCrypt.StringVerschluesseln(tempPassword);
                updatePassword(true, true, aSessionObj.Project.userPassword.passwordSystemValid, aProjectID);
            }
            // Alle Einheiten zur Benutzerkennung ermitteln
            string unitList = "";
            string unitFirst = "";
            string unitFirstShort = "";
            SqlDB dataReader;
            TParameterList parameterList = new TParameterList();
            parameterList.addParameter("userID", "string", UserID);
            dataReader = new SqlDB("select orgID from mapuserorgid where userID=@userID", parameterList, aProjectID);
            while (dataReader.read())
            {
                string actDisplayName = "";
                string actDisplayNameShort = "";
                int actOrgID = dataReader.getInt32(0);
                TAdminStructure.getOrgName(actOrgID, ref actDisplayName, ref actDisplayNameShort, aProjectID);
                if (unitFirst == "")
                    unitFirst = actDisplayName + " (" + actDisplayNameShort + ") (" + actOrgID.ToString() + ")";
                if (unitFirstShort == "")
                    unitFirstShort = actDisplayNameShort;
                unitList = unitList + actDisplayName + " (" + actDisplayNameShort + ") (" + actOrgID.ToString() + ")<br>";
            }
            dataReader.close();

            System.Net.Mail.MailAddress From = new System.Net.Mail.MailAddress(mailText[0], mailText[1]);
            System.Net.Mail.MailAddress To = new System.Net.Mail.MailAddress(Email);
            System.Net.Mail.MailMessage Message = new System.Net.Mail.MailMessage(From, To);
            Message.SubjectEncoding = Encoding.UTF8;
            Message.BodyEncoding = Encoding.UTF8;
            Message.IsBodyHtml = true;
            Message.ReplyToList.Add(mailText[2]);
            // Platzhalter ersetzen
            for (int lineIndex = 3; lineIndex < mailText.Count; lineIndex++)
            {
                mailText[lineIndex] = mailText[lineIndex].Replace("#FirstName#", Firstname);
                mailText[lineIndex] = mailText[lineIndex].Replace("#Name#", Name);
                mailText[lineIndex] = mailText[lineIndex].Replace("#Title#", Title);
                mailText[lineIndex] = mailText[lineIndex].Replace("#UserID#", UserID);
                mailText[lineIndex] = mailText[lineIndex].Replace("#Password#", tempPassword);

                mailText[lineIndex] = mailText[lineIndex].Replace("#DisplayName#", DisplayName);
                mailText[lineIndex] = mailText[lineIndex].Replace("#Street#", Street);
                mailText[lineIndex] = mailText[lineIndex].Replace("#ZipCode#", Zipcode);
                mailText[lineIndex] = mailText[lineIndex].Replace("#City#", City);
                mailText[lineIndex] = mailText[lineIndex].Replace("#State#", State);
                mailText[lineIndex] = mailText[lineIndex].Replace("#Comment#", Comment);
                mailText[lineIndex] = mailText[lineIndex].Replace("#GroupID#", GroupID);

                mailText[lineIndex] = mailText[lineIndex].Replace("#UnitFirst#", unitFirst);
                mailText[lineIndex] = mailText[lineIndex].Replace("#UnitFirstShort#", unitFirstShort);
                mailText[lineIndex] = mailText[lineIndex].Replace("#UnitList#", unitList);

                string tempLink = aSessionObj.Domain + "/" + aSessionObj.StartApp + "?project=" + aProjectID;
                mailText[lineIndex] = mailText[lineIndex].Replace("#SISLink#", tempLink);

                if (lineIndex == 3)
                    Message.Subject = mailText[3];
                else
                    Message.Body = Message.Body + mailText[lineIndex];// + (char)10 + (char)13;
            }

            System.Net.NetworkCredential mailClientCredentials = new System.Net.NetworkCredential(aUser, aPassword);
            System.Net.Mail.SmtpClient Client = new System.Net.Mail.SmtpClient(aServer);
            Client.Credentials = mailClientCredentials;
            Client.Send(Message);

            // Mailversand protokollieren
            parameterList = new TParameterList();
            parameterList.addParameter("userID", "string", UserID);
            parameterList.addParameter("email", "string", Email);
            parameterList.addParameter("action", "int", "1");
            parameterList.addParameter("datetime", "datetime", DateTime.Now.ToString());
            parameterList.addParameter("comment", "string", "Einladung durch Admin: " + aMailTemplate);
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO maillog (userID, email, action, datetime, comment) VALUES (@userID, @email, @action, @datetime, @comment)", parameterList);
        }
    }
    /// <summary>
    /// Aktualiserung der Kennwort-Eigenschaften des Benutzers
    /// </summary>
    /// <param name="aProjectID">ID des Projektes</param>
    public void updatePassword(bool aForceChange, bool aSystemGeneratedPassword, int aValidDays, string aProjectID)
    {
        // Daten übernehmen
        ForceChange = aForceChange;
        SystemGeneratedPassword = aSystemGeneratedPassword;
        PasswordValidDate = DateTime.Now.AddDays(aValidDays);

        string tempForceChange = "0";
        if (ForceChange)
            tempForceChange = "1";
        string tempSystemGeneratedPassword = "0";
        if (SystemGeneratedPassword)
            tempSystemGeneratedPassword = "1";

        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("password", "string", Password);
        parameterList.addParameter("lastLoginErrorCount", "int", "0");
        parameterList.addParameter("forceChange", "int", tempForceChange);
        parameterList.addParameter("systemgeneratedpassword", "int", tempSystemGeneratedPassword);
        parameterList.addParameter("passwordValidDate", "datetime", PasswordValidDate.ToString());
        parameterList.addParameter("userID", "string", UserID);
        SqlDB dataReader;
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("UPDATE loginuser SET password=@password, lastLoginErrorCount=@lastLoginErrorCount, forceChange=@forceChange, systemgeneratedpassword=@systemgeneratedpassword, passwordValidDate=@passwordValidDate WHERE userID=@userID", parameterList);
        // Passwort-Historie ergänzen
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQLwithParameter("INSERT INTO loginpasswordhistory (userID, password) VALUES (@userID, @password)", parameterList);
    }
    public bool hasContainerRights(int aContainerID)
    {
        bool result = false;
        foreach (TContainerRight tempContainerRight in containerRights)
        {
            if (tempContainerRight.ContainerID == aContainerID)
                result = true;
        }
        return result;
    }

    public static TOrgRight getOrgRight(string aUserID, int aOrgID, string aProjectID)
    {
        TOrgRight tempUnitRight = new TOrgRight();
        tempUnitRight.OrgID = aOrgID;

        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("userID", "string", aUserID);
        parameterList.addParameter("orgID", "string", aOrgID.ToString());
        SqlDB dataReader = new SqlDB("select ID, orgID, OProfile, OReadLevel, OEditLevel, OAllowDelegate, OAllowLock, OAllowEmployeeRead, OAllowEmployeeEdit, OAllowEmployeeImport, OAllowEmployeeExport, OAllowUnitAdd, OAllowUnitMove, OAllowUnitDelete, OAllowUnitProperty, OAllowReportRecipient, OAllowStructureImport, OAllowStructureExport, OAllowBouncesRead, OAllowBouncesEdit, OAllowBouncesDelete, OAllowBouncesExport, OAllowBouncesImport, OAllowReInvitation, OAllowPostNomination, OAllowPostNominationImport, FProfile, FReadLevel, FEditLevel, FSumLevel, FAllowCommunications, FAllowMeasures, FAllowDelete, FAllowExcelExport, FAllowReminderOnMeasure, FAllowReminderOnUnit, FAllowUpUnitsMeasures, FAllowColleaguesMeasures, FAllowThemes, FAllowReminderOnTheme, FAllowReminderOnThemeUnit, FAllowTopUnitsThemes, FAllowColleaguesThemes, FAllowOrders, FAllowReminderOnOrderUnit, RProfile, RReadLevel, DProfile, DReadLevel, DAllowMail, DAllowDownload, DAllowLog, DAllowType1, DAllowType2, DAllowType3, DAllowType4, DAllowType5, AEditLevelWorkshop, AReadLevelStructure from mapuserorgid where userID=@userID AND orgID=@orgID", parameterList, aProjectID);
        if (dataReader.read())
        {
            tempUnitRight.ID = dataReader.getInt32(0);
            tempUnitRight.OrgID = dataReader.getInt32(1);
            tempUnitRight.OProfile = dataReader.getString(2);
            tempUnitRight.OReadLevel = dataReader.getInt32(3);
            tempUnitRight.OEditLevel = dataReader.getInt32(4);
            tempUnitRight.OAllowDelegate = dataReader.getBool(5);
            tempUnitRight.OAllowLock = dataReader.getBool(6);
            tempUnitRight.OAllowEmployeeRead = dataReader.getBool(7);
            tempUnitRight.OAllowEmployeeEdit = dataReader.getBool(8);
            tempUnitRight.OAllowEmployeeImport = dataReader.getBool(9);
            tempUnitRight.OAllowEmployeeExport = dataReader.getBool(10);
            tempUnitRight.OAllowUnitAdd = dataReader.getBool(11);
            tempUnitRight.OAllowUnitMove = dataReader.getBool(12);
            tempUnitRight.OAllowUnitDelete = dataReader.getBool(13);
            tempUnitRight.OAllowUnitProperty = dataReader.getBool(14);
            tempUnitRight.OAllowReportRecipient = dataReader.getBool(15);
            tempUnitRight.OAllowStructureImport = dataReader.getBool(16);
            tempUnitRight.OAllowStructureExport = dataReader.getBool(17);
            tempUnitRight.OAllowBouncesRead = dataReader.getBool(18);
            tempUnitRight.OAllowBouncesEdit = dataReader.getBool(19);
            tempUnitRight.OAllowBouncesDelete = dataReader.getBool(20);
            tempUnitRight.OAllowBouncesExport = dataReader.getBool(21);
            tempUnitRight.OAllowBouncesImport = dataReader.getBool(22);
            tempUnitRight.OAllowReInvitation = dataReader.getBool(23);
            tempUnitRight.OAllowPostNomination = dataReader.getBool(24);
            tempUnitRight.OAllowPostNominationImport = dataReader.getBool(25);
            tempUnitRight.FProfile = dataReader.getString(26);
            tempUnitRight.FReadLevel = dataReader.getInt32(27);
            tempUnitRight.FEditLevel = dataReader.getInt32(28);
            tempUnitRight.FSumLevel = dataReader.getInt32(29);
            tempUnitRight.FAllowCommunications = dataReader.getBool(30);
            tempUnitRight.FAllowMeasures = dataReader.getBool(31);
            tempUnitRight.FAllowDelete = dataReader.getBool(32);
            tempUnitRight.FAllowExcelExport = dataReader.getBool(33);
            tempUnitRight.FAllowReminderOnMeasure = dataReader.getBool(34);
            tempUnitRight.FAllowReminderOnUnit = dataReader.getBool(35);
            tempUnitRight.FAllowUpUnitsMeasures = dataReader.getBool(36);
            tempUnitRight.FAllowColleaguesMeasures = dataReader.getBool(37);
            tempUnitRight.FAllowThemes = dataReader.getBool(38);
            tempUnitRight.FAllowReminderOnThemes = dataReader.getBool(39);
            tempUnitRight.FAllowReminderOnThemeUnit = dataReader.getBool(40);
            tempUnitRight.FAllowTopUnitsThemes = dataReader.getBool(41);
            tempUnitRight.FAllowColleaguesThemes = dataReader.getBool(42);
            tempUnitRight.FAllowOrder = dataReader.getBool(43);
            tempUnitRight.FAllowReminderOnOrderUnit = dataReader.getBool(44);

            tempUnitRight.RProfile = dataReader.getString(45);
            tempUnitRight.RReadLevel = dataReader.getInt32(46);
            tempUnitRight.DProfile = dataReader.getString(47);
            tempUnitRight.DReadLevel = dataReader.getInt32(48);
            tempUnitRight.DAllowMail = dataReader.getBool(49);
            tempUnitRight.DAllowDownload = dataReader.getBool(50);
            tempUnitRight.DAllowLog = dataReader.getBool(51);
            tempUnitRight.DAllowType1 = dataReader.getBool(52);
            tempUnitRight.DAllowType2 = dataReader.getBool(53);
            tempUnitRight.DAllowType3 = dataReader.getBool(54);
            tempUnitRight.DAllowType4 = dataReader.getBool(55);
            tempUnitRight.DAllowType5 = dataReader.getBool(56);
            tempUnitRight.AEditLevelWorkshop = dataReader.getInt32(57);
            tempUnitRight.AReadLevelStructure = dataReader.getInt32(58);
        }
        dataReader.close();

        return tempUnitRight;
    }
    public static TOrg1Right getOrg1Right(string aUserID, int aOrgID, string aProjectID)
    {
        TOrg1Right tempUnitRight = new TOrg1Right();
        tempUnitRight.OrgID = aOrgID;

        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("userID", "string", aUserID);
        parameterList.addParameter("orgID", "string", aOrgID.ToString());
        SqlDB dataReader = new SqlDB("select ID, orgID, OProfile, OReadLevel, OEditLevel, OAllowDelegate, OAllowLock, OAllowEmployeeRead, OAllowEmployeeEdit, OAllowEmployeeImport, OAllowEmployeeExport, OAllowUnitAdd, OAllowUnitMove, OAllowUnitDelete, OAllowUnitProperty, OAllowReportRecipient, OAllowStructureImport, OAllowStructureExport, FProfile, FReadLevel, FEditLevel, FSumLevel, FAllowCommunications, FAllowMeasures, FAllowDelete, FAllowExcelExport, FAllowReminderOnMeasure, FAllowReminderOnUnit, FAllowUpUnitsMeasures, FAllowColleaguesMeasures, FAllowThemes, FAllowReminderOnTheme, FAllowReminderOnThemeUnit, FAllowTopUnitsThemes, FAllowColleaguesThemes, FAllowOrders, FAllowReminderOnOrderUnit, RProfile, RReadLevel, DProfile, DReadLevel, DAllowMail, DAllowDownload, DAllowLog, DAllowType1, DAllowType2, DAllowType3, DAllowType4, DAllowType5, AEditLevelWorkshop, AReadLevelStructure from mapuserorgid1 where userID=@userID AND orgID=@orgID", parameterList, aProjectID);
        if (dataReader.read())
        {
            tempUnitRight.ID = dataReader.getInt32(0);
            tempUnitRight.OrgID = dataReader.getInt32(1);
            tempUnitRight.OProfile = dataReader.getString(2);
            tempUnitRight.OReadLevel = dataReader.getInt32(3);
            tempUnitRight.OEditLevel = dataReader.getInt32(4);
            tempUnitRight.OAllowDelegate = dataReader.getBool(5);
            tempUnitRight.OAllowLock = dataReader.getBool(6);
            tempUnitRight.OAllowEmployeeRead = dataReader.getBool(7);
            tempUnitRight.OAllowEmployeeEdit = dataReader.getBool(8);
            tempUnitRight.OAllowEmployeeImport = dataReader.getBool(9);
            tempUnitRight.OAllowEmployeeExport = dataReader.getBool(10);
            tempUnitRight.OAllowUnitAdd = dataReader.getBool(11);
            tempUnitRight.OAllowUnitMove = dataReader.getBool(12);
            tempUnitRight.OAllowUnitDelete = dataReader.getBool(13);
            tempUnitRight.OAllowUnitProperty = dataReader.getBool(14);
            tempUnitRight.OAllowReportRecipient = dataReader.getBool(15);
            tempUnitRight.OAllowStructureImport = dataReader.getBool(16);
            tempUnitRight.OAllowStructureExport = dataReader.getBool(17);
            tempUnitRight.FProfile = dataReader.getString(18);
            tempUnitRight.FReadLevel = dataReader.getInt32(19);
            tempUnitRight.FEditLevel = dataReader.getInt32(20);
            tempUnitRight.FSumLevel = dataReader.getInt32(21);
            tempUnitRight.FAllowCommunications = dataReader.getBool(22);
            tempUnitRight.FAllowMeasures = dataReader.getBool(23);
            tempUnitRight.FAllowDelete = dataReader.getBool(24);
            tempUnitRight.FAllowExcelExport = dataReader.getBool(25);
            tempUnitRight.FAllowReminderOnMeasure = dataReader.getBool(26);
            tempUnitRight.FAllowReminderOnUnit = dataReader.getBool(27);
            tempUnitRight.FAllowUpUnitsMeasures = dataReader.getBool(28);
            tempUnitRight.FAllowColleaguesMeasures = dataReader.getBool(29);
            tempUnitRight.FAllowThemes = dataReader.getBool(30);
            tempUnitRight.FAllowReminderOnThemes = dataReader.getBool(31);
            tempUnitRight.FAllowReminderOnThemeUnit = dataReader.getBool(32);
            tempUnitRight.FAllowTopUnitsThemes = dataReader.getBool(33);
            tempUnitRight.FAllowColleaguesThemes = dataReader.getBool(34);
            tempUnitRight.FAllowOrder = dataReader.getBool(35);
            tempUnitRight.FAllowReminderOnOrderUnit = dataReader.getBool(36);

            tempUnitRight.RProfile = dataReader.getString(37);
            tempUnitRight.RReadLevel = dataReader.getInt32(38);
            tempUnitRight.DProfile = dataReader.getString(39);
            tempUnitRight.DReadLevel = dataReader.getInt32(40);
            tempUnitRight.DAllowMail = dataReader.getBool(41);
            tempUnitRight.DAllowDownload = dataReader.getBool(42);
            tempUnitRight.DAllowLog = dataReader.getBool(43);
            tempUnitRight.DAllowType1 = dataReader.getBool(44);
            tempUnitRight.DAllowType2 = dataReader.getBool(45);
            tempUnitRight.DAllowType3 = dataReader.getBool(46);
            tempUnitRight.DAllowType4 = dataReader.getBool(47);
            tempUnitRight.DAllowType5 = dataReader.getBool(48);
            tempUnitRight.AEditLevelWorkshop = dataReader.getInt32(49);
            tempUnitRight.AReadLevelStructure = dataReader.getInt32(50);
        }
        dataReader.close();

        return tempUnitRight;
    }
    public static TOrg2Right getOrg2Right(string aUserID, int aOrgID, string aProjectID)
    {
        TOrg2Right tempUnitRight = new TOrg2Right();
        tempUnitRight.OrgID = aOrgID;

        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("userID", "string", aUserID);
        parameterList.addParameter("orgID", "string", aOrgID.ToString());
        SqlDB dataReader = new SqlDB("select ID, orgID, OProfile, OReadLevel, OEditLevel, OAllowDelegate, OAllowLock, OAllowEmployeeRead, OAllowEmployeeEdit, OAllowEmployeeImport, OAllowEmployeeExport, OAllowUnitAdd, OAllowUnitMove, OAllowUnitDelete, OAllowUnitProperty, OAllowReportRecipient, OAllowStructureImport, OAllowStructureExport, FProfile, FReadLevel, FEditLevel, FSumLevel, FAllowCommunications, FAllowMeasures, FAllowDelete, FAllowExcelExport, FAllowReminderOnMeasure, FAllowReminderOnUnit, FAllowUpUnitsMeasures, FAllowColleaguesMeasures, FAllowThemes, FAllowReminderOnTheme, FAllowReminderOnThemeUnit, FAllowTopUnitsThemes, FAllowColleaguesThemes, FAllowOrders, FAllowReminderOnOrderUnit, RProfile, RReadLevel, DProfile, DReadLevel, DAllowMail, DAllowDownload, DAllowLog, DAllowType1, DAllowType2, DAllowType3, DAllowType4, DAllowType5, AEditLevelWorkshop, AReadLevelStructure from mapuserorgid2 where userID=@userID AND orgID=@orgID", parameterList, aProjectID);
        if (dataReader.read())
        {
            tempUnitRight.ID = dataReader.getInt32(0);
            tempUnitRight.OrgID = dataReader.getInt32(1);
            tempUnitRight.OProfile = dataReader.getString(2);
            tempUnitRight.OReadLevel = dataReader.getInt32(3);
            tempUnitRight.OEditLevel = dataReader.getInt32(4);
            tempUnitRight.OAllowDelegate = dataReader.getBool(5);
            tempUnitRight.OAllowLock = dataReader.getBool(6);
            tempUnitRight.OAllowEmployeeRead = dataReader.getBool(7);
            tempUnitRight.OAllowEmployeeEdit = dataReader.getBool(8);
            tempUnitRight.OAllowEmployeeImport = dataReader.getBool(9);
            tempUnitRight.OAllowEmployeeExport = dataReader.getBool(10);
            tempUnitRight.OAllowUnitAdd = dataReader.getBool(11);
            tempUnitRight.OAllowUnitMove = dataReader.getBool(12);
            tempUnitRight.OAllowUnitDelete = dataReader.getBool(13);
            tempUnitRight.OAllowUnitProperty = dataReader.getBool(14);
            tempUnitRight.OAllowReportRecipient = dataReader.getBool(15);
            tempUnitRight.OAllowStructureImport = dataReader.getBool(16);
            tempUnitRight.OAllowStructureExport = dataReader.getBool(17);
            tempUnitRight.FProfile = dataReader.getString(18);
            tempUnitRight.FReadLevel = dataReader.getInt32(19);
            tempUnitRight.FEditLevel = dataReader.getInt32(20);
            tempUnitRight.FSumLevel = dataReader.getInt32(21);
            tempUnitRight.FAllowCommunications = dataReader.getBool(22);
            tempUnitRight.FAllowMeasures = dataReader.getBool(23);
            tempUnitRight.FAllowDelete = dataReader.getBool(24);
            tempUnitRight.FAllowExcelExport = dataReader.getBool(25);
            tempUnitRight.FAllowReminderOnMeasure = dataReader.getBool(26);
            tempUnitRight.FAllowReminderOnUnit = dataReader.getBool(27);
            tempUnitRight.FAllowUpUnitsMeasures = dataReader.getBool(28);
            tempUnitRight.FAllowColleaguesMeasures = dataReader.getBool(29);
            tempUnitRight.FAllowThemes = dataReader.getBool(30);
            tempUnitRight.FAllowReminderOnThemes = dataReader.getBool(31);
            tempUnitRight.FAllowReminderOnThemeUnit = dataReader.getBool(32);
            tempUnitRight.FAllowTopUnitsThemes = dataReader.getBool(33);
            tempUnitRight.FAllowColleaguesThemes = dataReader.getBool(34);
            tempUnitRight.FAllowOrder = dataReader.getBool(35);
            tempUnitRight.FAllowReminderOnOrderUnit = dataReader.getBool(36);

            tempUnitRight.RProfile = dataReader.getString(37);
            tempUnitRight.RReadLevel = dataReader.getInt32(38);
            tempUnitRight.DProfile = dataReader.getString(39);
            tempUnitRight.DReadLevel = dataReader.getInt32(40);
            tempUnitRight.DAllowMail = dataReader.getBool(41);
            tempUnitRight.DAllowDownload = dataReader.getBool(42);
            tempUnitRight.DAllowLog = dataReader.getBool(43);
            tempUnitRight.DAllowType1 = dataReader.getBool(44);
            tempUnitRight.DAllowType2 = dataReader.getBool(45);
            tempUnitRight.DAllowType3 = dataReader.getBool(46);
            tempUnitRight.DAllowType4 = dataReader.getBool(47);
            tempUnitRight.DAllowType5 = dataReader.getBool(48);
            tempUnitRight.AEditLevelWorkshop = dataReader.getInt32(49);
            tempUnitRight.AReadLevelStructure = dataReader.getInt32(50);
        }
        dataReader.close();

        return tempUnitRight;
    }
    public static TOrg3Right getOrg3Right(string aUserID, int aOrgID, string aProjectID)
    {
        TOrg3Right tempUnitRight = new TOrg3Right();
        tempUnitRight.OrgID = aOrgID;

        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("userID", "string", aUserID);
        parameterList.addParameter("orgID", "string", aOrgID.ToString());
        SqlDB dataReader = new SqlDB("select ID, orgID, OProfile, OReadLevel, OEditLevel, OAllowDelegate, OAllowLock, OAllowEmployeeRead, OAllowEmployeeEdit, OAllowEmployeeImport, OAllowEmployeeExport, OAllowUnitAdd, OAllowUnitMove, OAllowUnitDelete, OAllowUnitProperty, OAllowReportRecipient, OAllowStructureImport, OAllowStructureExport, FProfile, FReadLevel, FEditLevel, FSumLevel, FAllowCommunications, FAllowMeasures, FAllowDelete, FAllowExcelExport, FAllowReminderOnMeasure, FAllowReminderOnUnit, FAllowUpUnitsMeasures, FAllowColleaguesMeasures, FAllowThemes, FAllowReminderOnTheme, FAllowReminderOnThemeUnit, FAllowTopUnitsThemes, FAllowColleaguesThemes, FAllowOrders, FAllowReminderOnOrderUnit, RProfile, RReadLevel, DProfile, DReadLevel, DAllowMail, DAllowDownload, DAllowLog, DAllowType1, DAllowType2, DAllowType3, DAllowType4, DAllowType5, AEditLevelWorkshop, AReadLevelStructure from mapuserorgid3 where userID=@userID AND orgID=@orgID", parameterList, aProjectID);
        if (dataReader.read())
        {
            tempUnitRight.ID = dataReader.getInt32(0);
            tempUnitRight.OrgID = dataReader.getInt32(1);
            tempUnitRight.OProfile = dataReader.getString(2);
            tempUnitRight.OReadLevel = dataReader.getInt32(3);
            tempUnitRight.OEditLevel = dataReader.getInt32(4);
            tempUnitRight.OAllowDelegate = dataReader.getBool(5);
            tempUnitRight.OAllowLock = dataReader.getBool(6);
            tempUnitRight.OAllowEmployeeRead = dataReader.getBool(7);
            tempUnitRight.OAllowEmployeeEdit = dataReader.getBool(8);
            tempUnitRight.OAllowEmployeeImport = dataReader.getBool(9);
            tempUnitRight.OAllowEmployeeExport = dataReader.getBool(10);
            tempUnitRight.OAllowUnitAdd = dataReader.getBool(11);
            tempUnitRight.OAllowUnitMove = dataReader.getBool(12);
            tempUnitRight.OAllowUnitDelete = dataReader.getBool(13);
            tempUnitRight.OAllowUnitProperty = dataReader.getBool(14);
            tempUnitRight.OAllowReportRecipient = dataReader.getBool(15);
            tempUnitRight.OAllowStructureImport = dataReader.getBool(16);
            tempUnitRight.OAllowStructureExport = dataReader.getBool(17);
            tempUnitRight.FProfile = dataReader.getString(18);
            tempUnitRight.FReadLevel = dataReader.getInt32(19);
            tempUnitRight.FEditLevel = dataReader.getInt32(20);
            tempUnitRight.FSumLevel = dataReader.getInt32(21);
            tempUnitRight.FAllowCommunications = dataReader.getBool(22);
            tempUnitRight.FAllowMeasures = dataReader.getBool(23);
            tempUnitRight.FAllowDelete = dataReader.getBool(24);
            tempUnitRight.FAllowExcelExport = dataReader.getBool(25);
            tempUnitRight.FAllowReminderOnMeasure = dataReader.getBool(26);
            tempUnitRight.FAllowReminderOnUnit = dataReader.getBool(27);
            tempUnitRight.FAllowUpUnitsMeasures = dataReader.getBool(28);
            tempUnitRight.FAllowColleaguesMeasures = dataReader.getBool(29);
            tempUnitRight.FAllowThemes = dataReader.getBool(30);
            tempUnitRight.FAllowReminderOnThemes = dataReader.getBool(31);
            tempUnitRight.FAllowReminderOnThemeUnit = dataReader.getBool(32);
            tempUnitRight.FAllowTopUnitsThemes = dataReader.getBool(33);
            tempUnitRight.FAllowColleaguesThemes = dataReader.getBool(34);
            tempUnitRight.FAllowOrder = dataReader.getBool(35);
            tempUnitRight.FAllowReminderOnOrderUnit = dataReader.getBool(36);

            tempUnitRight.RProfile = dataReader.getString(37);
            tempUnitRight.RReadLevel = dataReader.getInt32(38);
            tempUnitRight.DProfile = dataReader.getString(39);
            tempUnitRight.DReadLevel = dataReader.getInt32(40);
            tempUnitRight.DAllowMail = dataReader.getBool(41);
            tempUnitRight.DAllowDownload = dataReader.getBool(42);
            tempUnitRight.DAllowLog = dataReader.getBool(43);
            tempUnitRight.DAllowType1 = dataReader.getBool(44);
            tempUnitRight.DAllowType2 = dataReader.getBool(45);
            tempUnitRight.DAllowType3 = dataReader.getBool(46);
            tempUnitRight.DAllowType4 = dataReader.getBool(47);
            tempUnitRight.DAllowType5 = dataReader.getBool(48);
            tempUnitRight.AEditLevelWorkshop = dataReader.getInt32(49);
            tempUnitRight.AReadLevelStructure = dataReader.getInt32(50);
        }
        dataReader.close();

        return tempUnitRight;
    }
    public static TBackRight getBackRight(string aUserID, int aOrgID, string aProjectID)
    {
        TBackRight tempUnitRight = new TBackRight();
        tempUnitRight.OrgID = aOrgID;

        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("userID", "string", aUserID);
        parameterList.addParameter("orgID", "string", aOrgID.ToString());
        SqlDB dataReader = new SqlDB("select ID, orgID, OProfile, OReadLevel, AReadLevel from mapuserback where userID=@userID AND orgID=@orgID", parameterList, aProjectID);
        if (dataReader.read())
        {
            tempUnitRight.ID = dataReader.getInt32(0);
            tempUnitRight.OrgID = dataReader.getInt32(1);
            tempUnitRight.OProfile = dataReader.getString(2);
            tempUnitRight.OReadLevel = dataReader.getInt32(3);
            tempUnitRight.AReadLevel = dataReader.getInt32(4);
        }
        dataReader.close();

        return tempUnitRight;
    }
    public static TBack1Right getBack1Right(string aUserID, int aOrgID, string aProjectID)
    {
        TBack1Right tempUnitRight = new TBack1Right();
        tempUnitRight.OrgID = aOrgID;

        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("userID", "string", aUserID);
        parameterList.addParameter("orgID", "string", aOrgID.ToString());
        SqlDB dataReader = new SqlDB("select ID, orgID, OProfile, OReadLevel, AReadLevel from mapuserback1 where userID=@userID AND orgID=@orgID", parameterList, aProjectID);
        if (dataReader.read())
        {
            tempUnitRight.ID = dataReader.getInt32(0);
            tempUnitRight.OrgID = dataReader.getInt32(1);
            tempUnitRight.OProfile = dataReader.getString(2);
            tempUnitRight.OReadLevel = dataReader.getInt32(3);
            tempUnitRight.AReadLevel = dataReader.getInt32(4);
        }
        dataReader.close();

        return tempUnitRight;
    }
    public static TBack2Right getBack2Right(string aUserID, int aOrgID, string aProjectID)
    {
        TBack2Right tempUnitRight = new TBack2Right();
        tempUnitRight.OrgID = aOrgID;

        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("userID", "string", aUserID);
        parameterList.addParameter("orgID", "string", aOrgID.ToString());
        SqlDB dataReader = new SqlDB("select ID, orgID, OProfile, OReadLevel, AReadLevel from mapuserback2 where userID=@userID AND orgID=@orgID", parameterList, aProjectID);
        if (dataReader.read())
        {
            tempUnitRight.ID = dataReader.getInt32(0);
            tempUnitRight.OrgID = dataReader.getInt32(1);
            tempUnitRight.OProfile = dataReader.getString(2);
            tempUnitRight.OReadLevel = dataReader.getInt32(3);
            tempUnitRight.AReadLevel = dataReader.getInt32(4);
        }
        dataReader.close();

        return tempUnitRight;
    }
    public static TBack3Right getBack3Right(string aUserID, int aOrgID, string aProjectID)
    {
        TBack3Right tempUnitRight = new TBack3Right();
        tempUnitRight.OrgID = aOrgID;

        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("userID", "string", aUserID);
        parameterList.addParameter("orgID", "string", aOrgID.ToString());
        SqlDB dataReader = new SqlDB("select ID, orgID, OProfile, OReadLevel, AReadLevel from mapuserback3 where userID=@userID AND orgID=@orgID", parameterList, aProjectID);
        if (dataReader.read())
        {
            tempUnitRight.ID = dataReader.getInt32(0);
            tempUnitRight.OrgID = dataReader.getInt32(1);
            tempUnitRight.OProfile = dataReader.getString(2);
            tempUnitRight.OReadLevel = dataReader.getInt32(3);
            tempUnitRight.AReadLevel = dataReader.getInt32(4);
        }
        dataReader.close();

        return tempUnitRight;
    }
    public static TContainerRight getContainerRight(string aUserID, int aContainerID, string aProjectID)
    {
        TContainerRight tempContainerRight = new TContainerRight();
        tempContainerRight.ContainerID = aContainerID;

        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("userID", "string", aUserID);
        parameterList.addParameter("containerID", "string", aContainerID.ToString());
        SqlDB dataReader = new SqlDB("select ID, containerID, TContainerProfile, allowUpload, allowDownload, allowDeleteOwnFiles, allowDeleteAllFiles, allowAccessOwnFilesWithoutPassword, allowAccessAllFilesWithoutPassword, allowCreateFolder, allowDeleteOwnFolder, allowDeleteAllFolder, allowResetPassword, allowTakeOwnership from mapusercontainer where userID=@userID AND containerID=@containerID", parameterList, aProjectID);
        if (dataReader.read())
        {
            tempContainerRight.ID = dataReader.getInt32(0);
            tempContainerRight.ContainerID = dataReader.getInt32(1);
            tempContainerRight.TContainerProfile = dataReader.getString(2);
            tempContainerRight.TAllowUpload = dataReader.getBool(3);
            tempContainerRight.TAllowDownload = dataReader.getBool(4);
            tempContainerRight.TAllowDeleteOwnFiles = dataReader.getBool(5);
            tempContainerRight.TAllowDeleteAllFiles = dataReader.getBool(6);
            tempContainerRight.TAllowAccessOwnFilesWithoutPassword = dataReader.getBool(7);
            tempContainerRight.TAllowAccessAllFilesWithoutPassword = dataReader.getBool(8);
            tempContainerRight.TAllowCreateFolder = dataReader.getBool(9);
            tempContainerRight.TAllowDeleteOwnFolder = dataReader.getBool(10);
            tempContainerRight.TAllowDeleteAllFolder = dataReader.getBool(11);
            tempContainerRight.TAllowResetPassword = dataReader.getBool(12);
            tempContainerRight.TAllowTakeOwnership = dataReader.getBool(13);
        }
        dataReader.close();

        return tempContainerRight;
    }
    public static TUserRight getUserRight(string aUserID, string aProjectID)
    {
        TUserRight tempUserRight = new TUserRight();

        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("userID", "string", aUserID);
        SqlDB dataReader = new SqlDB("select ID, TTaskProfile, AllowTaskRead, AllowTaskCreateEditDelete, AllowTaskFileUpload, AllowTaskExport, AllowTaskImport, TContactProfile, AllowContactRead, AllowContactCreateEditDelete, AllowContactExport, AllowContactImport, TTranslateProfile, AllowTranslateRead, AllowTranslateEdit, AllowTranslateRelease, AllowTranslateNew, AllowTranslateDelete, AllowAccountTransfer, AllowAccountManagement from mapuserrights where userID=@userID", parameterList, aProjectID);
        if (dataReader.read())
        {
            tempUserRight.ID = dataReader.getInt32(0);
            tempUserRight.TTaskProfile = dataReader.getString(1);
            tempUserRight.TAllowTaskRead = dataReader.getBool(2);
            tempUserRight.TAllowTaskCreateEditDelete = dataReader.getBool(3);
            tempUserRight.TAllowTaskFileUpload = dataReader.getBool(4);
            tempUserRight.TAllowTaskExport = dataReader.getBool(5);
            tempUserRight.TAllowTaskImport = dataReader.getBool(6);
            tempUserRight.TContactProfile = dataReader.getString(7);
            tempUserRight.TAllowContactRead = dataReader.getBool(8);
            tempUserRight.TAllowContactCreateEditDelete = dataReader.getBool(9);
            tempUserRight.TAllowContactExport = dataReader.getBool(10);
            tempUserRight.TAllowContactImport = dataReader.getBool(11);
            tempUserRight.TTranslateProfile = dataReader.getString(12);
            tempUserRight.TAllowTranslateRead = dataReader.getBool(13);
            tempUserRight.TAllowTranslateEdit = dataReader.getBool(14);
            tempUserRight.TAllowTranslateRelease = dataReader.getBool(15);
            tempUserRight.TAllowTranslateNew = dataReader.getBool(16);
            tempUserRight.TAllowTranslateDelete = dataReader.getBool(17);
            tempUserRight.AAllowAccountTransfer = dataReader.getBool(18);
            tempUserRight.AAllowAccountManagement = dataReader.getBool(19);
        }
        dataReader.close();

        return tempUserRight;
    }
}

