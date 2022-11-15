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
/// Zusammenfassungsbeschreibung für TStructure
/// </summary>
public class TProfiles
{
    public class TProfileSetOrgManager
    {
        public string profileID = "";
        public int OReadLevel = 0;
        public int OEditLevel = 0;
        public bool OAllowDelegate = false;
        public bool OAllowLock = false;
        public bool OAllowEmployeeRead = false;
        public bool OAllowEmployeeEdit = false;
        public bool OAllowEmployeeImport = false;
        public bool OAllowEmployeeExport = false;
        public bool OAllowUnitAdd = false;
        public bool OAllowUnitMove = false;
        public bool OAllowUnitDelete = false;
        public bool OAllowUnitProperty = false;
        public bool OAllowReportRecipient = false;
        public bool OAllowStructureImport = false;
        public bool OAllowStructureExport = false;
        public bool OAllowBouncesRead = false;
        public bool OAllowBouncesEdit = false;
        public bool OAllowBouncesDelete = false;
        public bool OAllowBouncesExport = false;
        public bool OAllowBouncesImport = false;
        public bool OAllowReInvitation = false;
        public bool OAllowPostNomination = false;
        public bool OAllowPostNominationImport = false;
    }
    public class TProfileSetOrgManager1
    {
        public string profileID = "";
        public int OReadLevel = 0;
        public int OEditLevel = 0;
        public bool OAllowDelegate = false;
        public bool OAllowLock = false;
        public bool OAllowEmployeeRead = false;
        public bool OAllowEmployeeEdit = false;
        public bool OAllowEmployeeImport = false;
        public bool OAllowEmployeeExport = false;
        public bool OAllowUnitAdd = false;
        public bool OAllowUnitMove = false;
        public bool OAllowUnitDelete = false;
        public bool OAllowUnitProperty = false;
        public bool OAllowReportRecipient = false;
        public bool OAllowStructureImport = false;
        public bool OAllowStructureExport = false;
    }
    public class TProfileSetOrgManager2
    {
        public string profileID = "";
        public int OReadLevel = 0;
        public int OEditLevel = 0;
        public bool OAllowDelegate = false;
        public bool OAllowLock = false;
        public bool OAllowEmployeeRead = false;
        public bool OAllowEmployeeEdit = false;
        public bool OAllowEmployeeImport = false;
        public bool OAllowEmployeeExport = false;
        public bool OAllowUnitAdd = false;
        public bool OAllowUnitMove = false;
        public bool OAllowUnitDelete = false;
        public bool OAllowUnitProperty = false;
        public bool OAllowReportRecipient = false;
        public bool OAllowStructureImport = false;
        public bool OAllowStructureExport = false;
    }
    public class TProfileSetOrgManager3
    {
        public string profileID = "";
        public int OReadLevel = 0;
        public int OEditLevel = 0;
        public bool OAllowDelegate = false;
        public bool OAllowLock = false;
        public bool OAllowEmployeeRead = false;
        public bool OAllowEmployeeEdit = false;
        public bool OAllowEmployeeImport = false;
        public bool OAllowEmployeeExport = false;
        public bool OAllowUnitAdd = false;
        public bool OAllowUnitMove = false;
        public bool OAllowUnitDelete = false;
        public bool OAllowUnitProperty = false;
        public bool OAllowReportRecipient = false;
        public bool OAllowStructureImport = false;
        public bool OAllowStructureExport = false;
    }
    public class TProfileSetOrgManagerBack
    {
        public string profileID = "";
        public int OReadLevel = 0;
    }
    public class TProfileSetOrgManager1Back
    {
        public string profileID = "";
        public int OReadLevel = 0;
    }
    public class TProfileSetOrgManager2Back
    {
        public string profileID = "";
        public int OReadLevel = 0;
    }
    public class TProfileSetOrgManager3Back
    {
        public string profileID = "";
        public int OReadLevel = 0;
    }
    public class TProfileSetResponse
    {
        public string profileID = "";
        public int RReadLevel = 0;
    }
    public class TProfileSetResponse1
    {
        public string profileID = "";
        public int RReadLevel = 0;
    }
    public class TProfileSetResponse2
    {
        public string profileID = "";
        public int RReadLevel = 0;
    }
    public class TProfileSetResponse3
    {
        public string profileID = "";
        public int RReadLevel = 0;
    }
    public class TProfileSetDownload
    {
        public string profileID = "";
        public int DReadLevel = 0;
        public bool DAllowMail = false;
        public bool DAllowDownload = false;
        public bool DAllowLog = false;
        public bool DAllowType1 = false;
        public bool DAllowType2 = false;
        public bool DAllowType3 = false;
        public bool DAllowType4 = false;
        public bool DAllowType5 = false;
    }
    public class TProfileSetDownload1
    {
        public string profileID = "";
        public int DReadLevel = 0;
        public bool DAllowMail = false;
        public bool DAllowDownload = false;
        public bool DAllowLog = false;
        public bool DAllowType1 = false;
        public bool DAllowType2 = false;
        public bool DAllowType3 = false;
        public bool DAllowType4 = false;
        public bool DAllowType5 = false;
    }
    public class TProfileSetDownload2
    {
        public string profileID = "";
        public int DReadLevel = 0;
        public bool DAllowMail = false;
        public bool DAllowDownload = false;
        public bool DAllowLog = false;
        public bool DAllowType1 = false;
        public bool DAllowType2 = false;
        public bool DAllowType3 = false;
        public bool DAllowType4 = false;
        public bool DAllowType5 = false;
    }
    public class TProfileSetDownload3
    {
        public string profileID = "";
        public int DReadLevel = 0;
        public bool DAllowMail = false;
        public bool DAllowDownload = false;
        public bool DAllowLog = false;
        public bool DAllowType1 = false;
        public bool DAllowType2 = false;
        public bool DAllowType3 = false;
        public bool DAllowType4 = false;
        public bool DAllowType5 = false;
    }
    public class TProfileSetFollowUp
    {
        public string profileID = "";
        public int FReadLevel = 0;
        public int FEditLevel = 0;
        public int FSumLevel = 0;
        public bool FAllowCommunications = false;
        public bool FAllowMeasures = false;
        public bool FAllowDelete = false;
        public bool FAllowExcelExport = false;
        public bool FAllowReminderOnMeasure = false;
        public bool FAllowReminderOnUnit = false;
        public bool FAllowUpUnitsMeasures = false;
        public bool FAllowColleaguesMeasures = false;
        public bool FAllowTopUnitsThemes = false;
        public bool FAllowColleaguesThemes = false;
        public bool FAllowThemes = false;
        public bool FAllowReminderOnTheme = false;
        public bool FAllowReminderOnThemeUnit = false;
        public bool FAllowOrders = false;
        public bool FAllowReminderOnOrderUnit = false;
    }
    public class TProfileSetFollowUp1
    {
        public string profileID = "";
        public int FReadLevel = 0;
        public int FEditLevel = 0;
        public int FSumLevel = 0;
        public bool FAllowCommunications = false;
        public bool FAllowMeasures = false;
        public bool FAllowDelete = false;
        public bool FAllowExcelExport = false;
        public bool FAllowReminderOnMeasure = false;
        public bool FAllowReminderOnUnit = false;
        public bool FAllowUpUnitsMeasures = false;
        public bool FAllowColleaguesMeasures = false;
        public bool FAllowTopUnitsThemes = false;
        public bool FAllowColleaguesThemes = false;
        public bool FAllowThemes = false;
        public bool FAllowReminderOnTheme = false;
        public bool FAllowReminderOnThemeUnit = false;
        public bool FAllowOrders = false;
        public bool FAllowReminderOnOrderUnit = false;
    }
    public class TProfileSetFollowUp2
    {
        public string profileID = "";
        public int FReadLevel = 0;
        public int FEditLevel = 0;
        public int FSumLevel = 0;
        public bool FAllowCommunications = false;
        public bool FAllowMeasures = false;
        public bool FAllowDelete = false;
        public bool FAllowExcelExport = false;
        public bool FAllowReminderOnMeasure = false;
        public bool FAllowReminderOnUnit = false;
        public bool FAllowUpUnitsMeasures = false;
        public bool FAllowColleaguesMeasures = false;
        public bool FAllowTopUnitsThemes = false;
        public bool FAllowColleaguesThemes = false;
        public bool FAllowThemes = false;
        public bool FAllowReminderOnTheme = false;
        public bool FAllowReminderOnThemeUnit = false;
        public bool FAllowOrders = false;
        public bool FAllowReminderOnOrderUnit = false;
    }
    public class TProfileSetFollowUp3
    {
        public string profileID = "";
        public int FReadLevel = 0;
        public int FEditLevel = 0;
        public int FSumLevel = 0;
        public bool FAllowCommunications = false;
        public bool FAllowMeasures = false;
        public bool FAllowDelete = false;
        public bool FAllowExcelExport = false;
        public bool FAllowReminderOnMeasure = false;
        public bool FAllowReminderOnUnit = false;
        public bool FAllowUpUnitsMeasures = false;
        public bool FAllowColleaguesMeasures = false;
        public bool FAllowTopUnitsThemes = false;
        public bool FAllowColleaguesThemes = false;
        public bool FAllowThemes = false;
        public bool FAllowReminderOnTheme = false;
        public bool FAllowReminderOnThemeUnit = false;
        public bool FAllowOrders = false;
        public bool FAllowReminderOnOrderUnit = false;
    }
    public class TProfileSetAccountManager
    {
        public string profileID = "";
        public int AEditLevel = 0;
        public bool AAllowAccountTransfer = false;
        public bool AAllowAccountManagement = false;
        public bool AAllowTeamspace = false;
        public bool AAllowTeamspaceOnly = false;
        public int AEditLevelWorkshop = 0;
    }
    public class TProfileSetContainer
    {
        public string profileID = "";
        public string TContainerProfile = "";
        public bool allowUpload = false;
        public bool allowDownload = false;
        public bool allowDeleteOwnFiles = false;
        public bool allowDeleteAllFiles = false;
        public bool allowAccessOwnFilesWithoutPassword = false;
        public bool allowAccessAllFilesWithoutPassword = false;
        public bool allowCreateFolder = false;
        public bool allowDeleteOwnFolder = false;
        public bool allowDeleteAllFolder = false;
        public bool allowResetPassword = false;
        public bool allowTakeOwnership = false;
    }
    public class TProfileSetProjectplan
    {
        public string profileID = "";
        public bool AllowTaskRead = false;
        public bool AllowTaskCreateEditDelete = false;
        public bool AllowTaskFileUpload = false;
        public bool AllowTaskExport = false;
        public bool AllowTaskImport = false;
    }
    public class TProfileSetContacts
    {
        public string profileID = "";
        public bool AllowContactRead = false;
        public bool AllowContactCreateEditDelete = false;
        public bool AllowContactExport = false;
        public bool AllowContactImport = false;
    }
    public class TProfileSetTranslate
    {
        public string profileID = "";
        public bool AllowTranslateRead = false;
        public bool AllowTranslateEdit = false;
        public bool AllowTranslateRelease = false;
        public bool AllowTranslateNew = false;
        public bool AllowTranslateDelete = false;
    }

    public TProfiles(string aTool, string aProfileName, string aProjectID)
    {
    }
    public static TProfileSetOrgManager getProfileOrgManager(string aProfileName, string aProjectID)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        TProfileSetOrgManager profileOrgManager = new TProfileSetOrgManager();

        SqlDB dataReader = new SqlDB(aProjectID);

        dataReader = new SqlDB("SELECT OReadLevel, OEditLevel, OAllowDelegate, OAllowLock, OAllowEmployeeRead, OAllowEmployeeEdit, OAllowEmployeeImport, OAllowEmployeeExport, OAllowUnitAdd, OAllowUnitMove, OAllowUnitDelete, OAllowUnitProperty, OAllowReportRecipient, OAllowStructureImport, OAllowStructureExport, OAllowBouncesRead, OAllowBouncesEdit, OAllowBouncesDelete, OAllowBouncesExport, OAllowBouncesImport, OAllowReInvitation, OAllowPostNomination, OAllowPostNominationImport FROM orgmanager_profiles WHERE profileID='" + aProfileName + "'", SessionObj.Project.ProjectID);
        if (dataReader.read())
        {
            profileOrgManager.profileID = aProfileName;
            profileOrgManager.OReadLevel = dataReader.getInt32(0);
            profileOrgManager.OEditLevel = dataReader.getInt32(1);
            profileOrgManager.OAllowDelegate = dataReader.getBool(2);
            profileOrgManager.OAllowLock = dataReader.getBool(3);
            profileOrgManager.OAllowEmployeeRead = dataReader.getBool(4);
            profileOrgManager.OAllowEmployeeEdit = dataReader.getBool(5);
            profileOrgManager.OAllowEmployeeImport = dataReader.getBool(6);
            profileOrgManager.OAllowEmployeeExport = dataReader.getBool(7);
            profileOrgManager.OAllowUnitAdd = dataReader.getBool(8);
            profileOrgManager.OAllowUnitMove = dataReader.getBool(9);
            profileOrgManager.OAllowUnitDelete = dataReader.getBool(10);
            profileOrgManager.OAllowUnitProperty = dataReader.getBool(11);
            profileOrgManager.OAllowReportRecipient = dataReader.getBool(12);
            profileOrgManager.OAllowStructureImport = dataReader.getBool(13);
            profileOrgManager.OAllowStructureExport = dataReader.getBool(14);
            profileOrgManager.OAllowBouncesRead = dataReader.getBool(15);
            profileOrgManager.OAllowBouncesEdit = dataReader.getBool(16);
            profileOrgManager.OAllowBouncesDelete = dataReader.getBool(17);
            profileOrgManager.OAllowBouncesExport = dataReader.getBool(18);
            profileOrgManager.OAllowBouncesImport = dataReader.getBool(19);
            profileOrgManager.OAllowReInvitation = dataReader.getBool(20);
            profileOrgManager.OAllowPostNomination = dataReader.getBool(21);
            profileOrgManager.OAllowPostNominationImport = dataReader.getBool(22);
        }
        return profileOrgManager;
    }
    public static TProfileSetOrgManager1 getProfileOrgManager1(string aProfileName, string aProjectID)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        TProfileSetOrgManager1 profileOrgManager1 = new TProfileSetOrgManager1();

        SqlDB dataReader = new SqlDB(aProjectID);

        dataReader = new SqlDB("SELECT OReadLevel, OEditLevel, OAllowDelegate, OAllowLock, OAllowEmployeeRead, OAllowEmployeeEdit, OAllowEmployeeImport, OAllowEmployeeExport, OAllowUnitAdd, OAllowUnitMove, OAllowUnitDelete, OAllowUnitProperty, OAllowReportRecipient, OAllowStructureImport, OAllowStructureExport FROM orgmanager_profiles1 WHERE profileID='" + aProfileName + "'", SessionObj.Project.ProjectID);
        if (dataReader.read())
        {
            profileOrgManager1.profileID = aProfileName;
            profileOrgManager1.OReadLevel = dataReader.getInt32(0);
            profileOrgManager1.OEditLevel = dataReader.getInt32(1);
            profileOrgManager1.OAllowDelegate = dataReader.getBool(2);
            profileOrgManager1.OAllowLock = dataReader.getBool(3);
            profileOrgManager1.OAllowEmployeeRead = dataReader.getBool(4);
            profileOrgManager1.OAllowEmployeeEdit = dataReader.getBool(5);
            profileOrgManager1.OAllowEmployeeImport = dataReader.getBool(6);
            profileOrgManager1.OAllowEmployeeExport = dataReader.getBool(7);
            profileOrgManager1.OAllowUnitAdd = dataReader.getBool(8);
            profileOrgManager1.OAllowUnitMove = dataReader.getBool(9);
            profileOrgManager1.OAllowUnitDelete = dataReader.getBool(10);
            profileOrgManager1.OAllowUnitProperty = dataReader.getBool(11);
            profileOrgManager1.OAllowReportRecipient = dataReader.getBool(12);
            profileOrgManager1.OAllowStructureImport = dataReader.getBool(13);
            profileOrgManager1.OAllowStructureExport = dataReader.getBool(14);
        }
        return profileOrgManager1;
    }
    public static TProfileSetOrgManager2 getProfileOrgManager2(string aProfileName, string aProjectID)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        TProfileSetOrgManager2 profileOrgManager2 = new TProfileSetOrgManager2();

        SqlDB dataReader = new SqlDB(aProjectID);

        dataReader = new SqlDB("SELECT OReadLevel, OEditLevel, OAllowDelegate, OAllowLock, OAllowEmployeeRead, OAllowEmployeeEdit, OAllowEmployeeImport, OAllowEmployeeExport, OAllowUnitAdd, OAllowUnitMove, OAllowUnitDelete, OAllowUnitProperty, OAllowReportRecipient, OAllowStructureImport, OAllowStructureExport FROM orgmanager_profiles2 WHERE profileID='" + aProfileName + "'", SessionObj.Project.ProjectID);
        if (dataReader.read())
        {
            profileOrgManager2.profileID = aProfileName;
            profileOrgManager2.OReadLevel = dataReader.getInt32(0);
            profileOrgManager2.OEditLevel = dataReader.getInt32(1);
            profileOrgManager2.OAllowDelegate = dataReader.getBool(2);
            profileOrgManager2.OAllowLock = dataReader.getBool(3);
            profileOrgManager2.OAllowEmployeeRead = dataReader.getBool(4);
            profileOrgManager2.OAllowEmployeeEdit = dataReader.getBool(5);
            profileOrgManager2.OAllowEmployeeImport = dataReader.getBool(6);
            profileOrgManager2.OAllowEmployeeExport = dataReader.getBool(7);
            profileOrgManager2.OAllowUnitAdd = dataReader.getBool(8);
            profileOrgManager2.OAllowUnitMove = dataReader.getBool(9);
            profileOrgManager2.OAllowUnitDelete = dataReader.getBool(10);
            profileOrgManager2.OAllowUnitProperty = dataReader.getBool(11);
            profileOrgManager2.OAllowReportRecipient = dataReader.getBool(12);
            profileOrgManager2.OAllowStructureImport = dataReader.getBool(13);
            profileOrgManager2.OAllowStructureExport = dataReader.getBool(14);
        }
        return profileOrgManager2;
    }
    public static TProfileSetOrgManager3 getProfileOrgManager3(string aProfileName, string aProjectID)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        TProfileSetOrgManager3 profileOrgManager3 = new TProfileSetOrgManager3();

        SqlDB dataReader = new SqlDB(aProjectID);

        dataReader = new SqlDB("SELECT OReadLevel, OEditLevel, OAllowDelegate, OAllowLock, OAllowEmployeeRead, OAllowEmployeeEdit, OAllowEmployeeImport, OAllowEmployeeExport, OAllowUnitAdd, OAllowUnitMove, OAllowUnitDelete, OAllowUnitProperty, OAllowReportRecipient, OAllowStructureImport, OAllowStructureExport FROM orgmanager_profiles3 WHERE profileID='" + aProfileName + "'", SessionObj.Project.ProjectID);
        if (dataReader.read())
        {
            profileOrgManager3.profileID = aProfileName;
            profileOrgManager3.OReadLevel = dataReader.getInt32(0);
            profileOrgManager3.OEditLevel = dataReader.getInt32(1);
            profileOrgManager3.OAllowDelegate = dataReader.getBool(2);
            profileOrgManager3.OAllowLock = dataReader.getBool(3);
            profileOrgManager3.OAllowEmployeeRead = dataReader.getBool(4);
            profileOrgManager3.OAllowEmployeeEdit = dataReader.getBool(5);
            profileOrgManager3.OAllowEmployeeImport = dataReader.getBool(6);
            profileOrgManager3.OAllowEmployeeExport = dataReader.getBool(7);
            profileOrgManager3.OAllowUnitAdd = dataReader.getBool(8);
            profileOrgManager3.OAllowUnitMove = dataReader.getBool(9);
            profileOrgManager3.OAllowUnitDelete = dataReader.getBool(10);
            profileOrgManager3.OAllowUnitProperty = dataReader.getBool(11);
            profileOrgManager3.OAllowReportRecipient = dataReader.getBool(12);
            profileOrgManager3.OAllowStructureImport = dataReader.getBool(13);
            profileOrgManager3.OAllowStructureExport = dataReader.getBool(14);
        }
        return profileOrgManager3;
    }
    public static TProfileSetOrgManagerBack getProfileOrgManagerBack(string aProfileName, string aProjectID)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        TProfileSetOrgManagerBack profileOrgManagerBack = new TProfileSetOrgManagerBack();

        SqlDB dataReader = new SqlDB(aProjectID);

        dataReader = new SqlDB("SELECT OReadLevel FROM orgmanager_profilesback WHERE profileID='" + aProfileName + "'", SessionObj.Project.ProjectID);
        if (dataReader.read())
        {
            profileOrgManagerBack.profileID = aProfileName;
            profileOrgManagerBack.OReadLevel = dataReader.getInt32(0);
        }
        return profileOrgManagerBack;
    }
    public static TProfileSetOrgManager1Back getProfileOrgManager1Back(string aProfileName, string aProjectID)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        TProfileSetOrgManager1Back profileOrgManager1Back = new TProfileSetOrgManager1Back();

        SqlDB dataReader = new SqlDB(aProjectID);

        dataReader = new SqlDB("SELECT OReadLevel FROM orgmanager_profiles1back WHERE profileID='" + aProfileName + "'", SessionObj.Project.ProjectID);
        if (dataReader.read())
        {
            profileOrgManager1Back.profileID = aProfileName;
            profileOrgManager1Back.OReadLevel = dataReader.getInt32(0);
        }
        return profileOrgManager1Back;
    }
    public static TProfileSetOrgManager2Back getProfileOrgManager2Back(string aProfileName, string aProjectID)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        TProfileSetOrgManager2Back profileOrgManager2Back = new TProfileSetOrgManager2Back();

        SqlDB dataReader = new SqlDB(aProjectID);

        dataReader = new SqlDB("SELECT OReadLevel FROM orgmanager_profiles2back WHERE profileID='" + aProfileName + "'", SessionObj.Project.ProjectID);
        if (dataReader.read())
        {
            profileOrgManager2Back.profileID = aProfileName;
            profileOrgManager2Back.OReadLevel = dataReader.getInt32(0);
        }
        return profileOrgManager2Back;
    }
    public static TProfileSetOrgManager3Back getProfileOrgManager3Back(string aProfileName, string aProjectID)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        TProfileSetOrgManager3Back profileOrgManager3Back = new TProfileSetOrgManager3Back();

        SqlDB dataReader = new SqlDB(aProjectID);

        dataReader = new SqlDB("SELECT OReadLevel FROM orgmanager_profiles3back WHERE profileID='" + aProfileName + "'", SessionObj.Project.ProjectID);
        if (dataReader.read())
        {
            profileOrgManager3Back.profileID = aProfileName;
            profileOrgManager3Back.OReadLevel = dataReader.getInt32(0);
        }
        return profileOrgManager3Back;
    }
    public static TProfileSetResponse getProfileResponse(string aProfileName, string aProjectID)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        TProfileSetResponse profileResponse = new TProfileSetResponse();

        SqlDB dataReader = new SqlDB(aProjectID);

        dataReader = new SqlDB("SELECT RReadLevel FROM response_profiles WHERE profileID='" + aProfileName + "'", SessionObj.Project.ProjectID);
        if (dataReader.read())
        {
            profileResponse.profileID = aProfileName;
            profileResponse.RReadLevel = dataReader.getInt32(0);
        }
        return profileResponse;
    }
    public static TProfileSetResponse1 getProfileResponse1(string aProfileName, string aProjectID)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        TProfileSetResponse1 profileResponse = new TProfileSetResponse1();

        SqlDB dataReader = new SqlDB(aProjectID);

        dataReader = new SqlDB("SELECT RReadLevel FROM response_profiles1 WHERE profileID='" + aProfileName + "'", SessionObj.Project.ProjectID);
        if (dataReader.read())
        {
            profileResponse.profileID = aProfileName;
            profileResponse.RReadLevel = dataReader.getInt32(0);
        }
        return profileResponse;
    }
    public static TProfileSetResponse2 getProfileResponse2(string aProfileName, string aProjectID)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        TProfileSetResponse2 profileResponse = new TProfileSetResponse2();

        SqlDB dataReader = new SqlDB(aProjectID);

        dataReader = new SqlDB("SELECT RReadLevel FROM response_profiles2 WHERE profileID='" + aProfileName + "'", SessionObj.Project.ProjectID);
        if (dataReader.read())
        {
            profileResponse.profileID = aProfileName;
            profileResponse.RReadLevel = dataReader.getInt32(0);
        }
        return profileResponse;
    }
    public static TProfileSetResponse3 getProfileResponse3(string aProfileName, string aProjectID)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        TProfileSetResponse3 profileResponse = new TProfileSetResponse3();

        SqlDB dataReader = new SqlDB(aProjectID);

        dataReader = new SqlDB("SELECT RReadLevel FROM response_profiles3 WHERE profileID='" + aProfileName + "'", SessionObj.Project.ProjectID);
        if (dataReader.read())
        {
            profileResponse.profileID = aProfileName;
            profileResponse.RReadLevel = dataReader.getInt32(0);
        }
        return profileResponse;
    }
    public static TProfileSetDownload getProfileDownload(string aProfileName, string aProjectID)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        TProfileSetDownload profileDownload = new TProfileSetDownload();

        SqlDB dataReader = new SqlDB(aProjectID);

        dataReader = new SqlDB("SELECT DReadLevel, DAllowMail, DAllowDownload, DAllowLog, DAllowType1, DAllowType2, DAllowType3, DAllowType4, DAllowType5 FROM download_profiles WHERE profileID='" + aProfileName + "'", SessionObj.Project.ProjectID);
        if (dataReader.read())
        {
            profileDownload.profileID = aProfileName;
            profileDownload.DReadLevel = dataReader.getInt32(0);
            profileDownload.DAllowMail = dataReader.getBool(1);
            profileDownload.DAllowDownload = dataReader.getBool(2);
            profileDownload.DAllowLog = dataReader.getBool(3);
            profileDownload.DAllowType1 = dataReader.getBool(4);
            profileDownload.DAllowType2 = dataReader.getBool(5);
            profileDownload.DAllowType3 = dataReader.getBool(6);
            profileDownload.DAllowType4 = dataReader.getBool(7);
            profileDownload.DAllowType5 = dataReader.getBool(8);
        }
        return profileDownload;
    }
    public static TProfileSetDownload1 getProfileDownload1(string aProfileName, string aProjectID)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        TProfileSetDownload1 profileDownload = new TProfileSetDownload1();

        SqlDB dataReader = new SqlDB(aProjectID);

        dataReader = new SqlDB("SELECT DReadLevel, DAllowMail, DAllowDownload, DAllowLog, DAllowType1, DAllowType2, DAllowType3, DAllowType4, DAllowType5 FROM download_profiles1 WHERE profileID='" + aProfileName + "'", SessionObj.Project.ProjectID);
        if (dataReader.read())
        {
            profileDownload.profileID = aProfileName;
            profileDownload.DReadLevel = dataReader.getInt32(0);
            profileDownload.DAllowMail = dataReader.getBool(1);
            profileDownload.DAllowDownload = dataReader.getBool(2);
            profileDownload.DAllowLog = dataReader.getBool(3);
            profileDownload.DAllowType1 = dataReader.getBool(4);
            profileDownload.DAllowType2 = dataReader.getBool(5);
            profileDownload.DAllowType3 = dataReader.getBool(6);
            profileDownload.DAllowType4 = dataReader.getBool(7);
            profileDownload.DAllowType5 = dataReader.getBool(8);
        }
        return profileDownload;
    }
    public static TProfileSetDownload2 getProfileDownload2(string aProfileName, string aProjectID)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        TProfileSetDownload2 profileDownload = new TProfileSetDownload2();

        SqlDB dataReader = new SqlDB(aProjectID);

        dataReader = new SqlDB("SELECT DReadLevel, DAllowMail, DAllowDownload, DAllowLog, DAllowType1, DAllowType2, DAllowType3, DAllowType4, DAllowType5 FROM download_profiles2 WHERE profileID='" + aProfileName + "'", SessionObj.Project.ProjectID);
        if (dataReader.read())
        {
            profileDownload.profileID = aProfileName;
            profileDownload.DReadLevel = dataReader.getInt32(0);
            profileDownload.DAllowMail = dataReader.getBool(1);
            profileDownload.DAllowDownload = dataReader.getBool(2);
            profileDownload.DAllowLog = dataReader.getBool(3);
            profileDownload.DAllowType1 = dataReader.getBool(4);
            profileDownload.DAllowType2 = dataReader.getBool(5);
            profileDownload.DAllowType3 = dataReader.getBool(6);
            profileDownload.DAllowType4 = dataReader.getBool(7);
            profileDownload.DAllowType5 = dataReader.getBool(8);
        }
        return profileDownload;
    }
    public static TProfileSetDownload3 getProfileDownload3(string aProfileName, string aProjectID)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        TProfileSetDownload3 profileDownload = new TProfileSetDownload3();

        SqlDB dataReader = new SqlDB(aProjectID);

        dataReader = new SqlDB("SELECT DReadLevel, DAllowMail, DAllowDownload, DAllowLog, DAllowType1, DAllowType2, DAllowType3, DAllowType4, DAllowType5 FROM download_profiles3 WHERE profileID='" + aProfileName + "'", SessionObj.Project.ProjectID);
        if (dataReader.read())
        {
            profileDownload.profileID = aProfileName;
            profileDownload.DReadLevel = dataReader.getInt32(0);
            profileDownload.DAllowMail = dataReader.getBool(1);
            profileDownload.DAllowDownload = dataReader.getBool(2);
            profileDownload.DAllowLog = dataReader.getBool(3);
            profileDownload.DAllowType1 = dataReader.getBool(4);
            profileDownload.DAllowType2 = dataReader.getBool(5);
            profileDownload.DAllowType3 = dataReader.getBool(6);
            profileDownload.DAllowType4 = dataReader.getBool(7);
            profileDownload.DAllowType5 = dataReader.getBool(8);
        }
        return profileDownload;
    }
    public static TProfileSetFollowUp getProfileFollowUp(string aProfileName, string aProjectID)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        TProfileSetFollowUp profileFollowUp = new TProfileSetFollowUp();

        SqlDB dataReader = new SqlDB(aProjectID);

        dataReader = new SqlDB("SELECT FReadLevel, FEditLevel, FSumLevel, FAllowCommunications, FAllowMeasures, FAllowDelete, FAllowExcelExport, FAllowReminderOnMeasure, FAllowReminderOnUnit, FAllowUpUnitsMeasures, FAllowColleaguesMeasures, FAllowTopUnitsThemes, FAllowColleaguesThemes, FAllowThemes, FAllowReminderOnTheme, FAllowReminderOnThemeUnit, FAllowOrders, FAllowReminderOnOrderUnit FROM followup_profiles WHERE profileID='" + aProfileName + "'", SessionObj.Project.ProjectID);
        if (dataReader.read())
        {
            profileFollowUp.profileID = aProfileName;
            profileFollowUp.FReadLevel = dataReader.getInt32(0);
            profileFollowUp.FEditLevel = dataReader.getInt32(1);
            profileFollowUp.FSumLevel = dataReader.getInt32(2);
            profileFollowUp.FAllowCommunications = dataReader.getBool(3);
            profileFollowUp.FAllowMeasures = dataReader.getBool(4);
            profileFollowUp.FAllowDelete = dataReader.getBool(5);
            profileFollowUp.FAllowExcelExport = dataReader.getBool(6);
            profileFollowUp.FAllowReminderOnMeasure = dataReader.getBool(7);
            profileFollowUp.FAllowReminderOnUnit = dataReader.getBool(8);
            profileFollowUp.FAllowUpUnitsMeasures = dataReader.getBool(9);
            profileFollowUp.FAllowColleaguesMeasures = dataReader.getBool(10);
            profileFollowUp.FAllowTopUnitsThemes = dataReader.getBool(11);
            profileFollowUp.FAllowColleaguesThemes = dataReader.getBool(12);
            profileFollowUp.FAllowThemes = dataReader.getBool(13);
            profileFollowUp.FAllowReminderOnTheme = dataReader.getBool(14);
            profileFollowUp.FAllowReminderOnThemeUnit = dataReader.getBool(15);
            profileFollowUp.FAllowOrders = dataReader.getBool(16);
            profileFollowUp.FAllowReminderOnOrderUnit = dataReader.getBool(17);
        }
        return profileFollowUp;
    }
    public static TProfileSetFollowUp1 getProfileFollowUp1(string aProfileName, string aProjectID)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        TProfileSetFollowUp1 profileFollowUp = new TProfileSetFollowUp1();

        SqlDB dataReader = new SqlDB(aProjectID);

        dataReader = new SqlDB("SELECT FReadLevel, FEditLevel, FSumLevel, FAllowCommunications, FAllowMeasures, FAllowDelete, FAllowExcelExport, FAllowReminderOnMeasure, FAllowReminderOnUnit, FAllowUpUnitsMeasures, FAllowColleaguesMeasures, FAllowTopUnitsThemes, FAllowColleaguesThemes, FAllowThemes, FAllowReminderOnTheme, FAllowReminderOnThemeUnit, FAllowOrders, FAllowReminderOnOrderUnit FROM followup_profiles1 WHERE profileID='" + aProfileName + "'", SessionObj.Project.ProjectID);
        if (dataReader.read())
        {
            profileFollowUp.profileID = aProfileName;
            profileFollowUp.FReadLevel = dataReader.getInt32(0);
            profileFollowUp.FEditLevel = dataReader.getInt32(1);
            profileFollowUp.FSumLevel = dataReader.getInt32(2);
            profileFollowUp.FAllowCommunications = dataReader.getBool(3);
            profileFollowUp.FAllowMeasures = dataReader.getBool(4);
            profileFollowUp.FAllowDelete = dataReader.getBool(5);
            profileFollowUp.FAllowExcelExport = dataReader.getBool(6);
            profileFollowUp.FAllowReminderOnMeasure = dataReader.getBool(7);
            profileFollowUp.FAllowReminderOnUnit = dataReader.getBool(8);
            profileFollowUp.FAllowUpUnitsMeasures = dataReader.getBool(9);
            profileFollowUp.FAllowColleaguesMeasures = dataReader.getBool(10);
            profileFollowUp.FAllowTopUnitsThemes = dataReader.getBool(11);
            profileFollowUp.FAllowColleaguesThemes = dataReader.getBool(12);
            profileFollowUp.FAllowThemes = dataReader.getBool(13);
            profileFollowUp.FAllowReminderOnTheme = dataReader.getBool(14);
            profileFollowUp.FAllowReminderOnThemeUnit = dataReader.getBool(15);
            profileFollowUp.FAllowOrders = dataReader.getBool(16);
            profileFollowUp.FAllowReminderOnOrderUnit = dataReader.getBool(17);
        }
        return profileFollowUp;
    }
    public static TProfileSetFollowUp2 getProfileFollowUp2(string aProfileName, string aProjectID)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        TProfileSetFollowUp2 profileFollowUp = new TProfileSetFollowUp2();

        SqlDB dataReader = new SqlDB(aProjectID);

        dataReader = new SqlDB("SELECT FReadLevel, FEditLevel, FSumLevel, FAllowCommunications, FAllowMeasures, FAllowDelete, FAllowExcelExport, FAllowReminderOnMeasure, FAllowReminderOnUnit, FAllowUpUnitsMeasures, FAllowColleaguesMeasures, FAllowTopUnitsThemes, FAllowColleaguesThemes, FAllowThemes, FAllowReminderOnTheme, FAllowReminderOnThemeUnit, FAllowOrders, FAllowReminderOnOrderUnit FROM followup_profiles2 WHERE profileID='" + aProfileName + "'", SessionObj.Project.ProjectID);
        if (dataReader.read())
        {
            profileFollowUp.profileID = aProfileName;
            profileFollowUp.FReadLevel = dataReader.getInt32(0);
            profileFollowUp.FEditLevel = dataReader.getInt32(1);
            profileFollowUp.FSumLevel = dataReader.getInt32(2);
            profileFollowUp.FAllowCommunications = dataReader.getBool(3);
            profileFollowUp.FAllowMeasures = dataReader.getBool(4);
            profileFollowUp.FAllowDelete = dataReader.getBool(5);
            profileFollowUp.FAllowExcelExport = dataReader.getBool(6);
            profileFollowUp.FAllowReminderOnMeasure = dataReader.getBool(7);
            profileFollowUp.FAllowReminderOnUnit = dataReader.getBool(8);
            profileFollowUp.FAllowUpUnitsMeasures = dataReader.getBool(9);
            profileFollowUp.FAllowColleaguesMeasures = dataReader.getBool(10);
            profileFollowUp.FAllowTopUnitsThemes = dataReader.getBool(11);
            profileFollowUp.FAllowColleaguesThemes = dataReader.getBool(12);
            profileFollowUp.FAllowThemes = dataReader.getBool(13);
            profileFollowUp.FAllowReminderOnTheme = dataReader.getBool(14);
            profileFollowUp.FAllowReminderOnThemeUnit = dataReader.getBool(15);
            profileFollowUp.FAllowOrders = dataReader.getBool(16);
            profileFollowUp.FAllowReminderOnOrderUnit = dataReader.getBool(17);
        }
        return profileFollowUp;
    }
    public static TProfileSetFollowUp3 getProfileFollowUp3(string aProfileName, string aProjectID)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        TProfileSetFollowUp3 profileFollowUp = new TProfileSetFollowUp3();

        SqlDB dataReader = new SqlDB(aProjectID);

        dataReader = new SqlDB("SELECT FReadLevel, FEditLevel, FSumLevel, FAllowCommunications, FAllowMeasures, FAllowDelete, FAllowExcelExport, FAllowReminderOnMeasure, FAllowReminderOnUnit, FAllowUpUnitsMeasures, FAllowColleaguesMeasures, FAllowTopUnitsThemes, FAllowColleaguesThemes, FAllowThemes, FAllowReminderOnTheme, FAllowReminderOnThemeUnit, FAllowOrders, FAllowReminderOnOrderUnit FROM followup_profiles3 WHERE profileID='" + aProfileName + "'", SessionObj.Project.ProjectID);
        if (dataReader.read())
        {
            profileFollowUp.profileID = aProfileName;
            profileFollowUp.FReadLevel = dataReader.getInt32(0);
            profileFollowUp.FEditLevel = dataReader.getInt32(1);
            profileFollowUp.FSumLevel = dataReader.getInt32(2);
            profileFollowUp.FAllowCommunications = dataReader.getBool(3);
            profileFollowUp.FAllowMeasures = dataReader.getBool(4);
            profileFollowUp.FAllowDelete = dataReader.getBool(5);
            profileFollowUp.FAllowExcelExport = dataReader.getBool(6);
            profileFollowUp.FAllowReminderOnMeasure = dataReader.getBool(7);
            profileFollowUp.FAllowReminderOnUnit = dataReader.getBool(8);
            profileFollowUp.FAllowUpUnitsMeasures = dataReader.getBool(9);
            profileFollowUp.FAllowColleaguesMeasures = dataReader.getBool(10);
            profileFollowUp.FAllowTopUnitsThemes = dataReader.getBool(11);
            profileFollowUp.FAllowColleaguesThemes = dataReader.getBool(12);
            profileFollowUp.FAllowThemes = dataReader.getBool(13);
            profileFollowUp.FAllowReminderOnTheme = dataReader.getBool(14);
            profileFollowUp.FAllowReminderOnThemeUnit = dataReader.getBool(15);
            profileFollowUp.FAllowOrders = dataReader.getBool(16);
            profileFollowUp.FAllowReminderOnOrderUnit = dataReader.getBool(17);
        }
        return profileFollowUp;
    }
    public static TProfileSetAccountManager getProfileAccountManager(string aProfileName, string aProjectID)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        TProfileSetAccountManager profileAccountManager = new TProfileSetAccountManager();

        SqlDB dataReader = new SqlDB(aProjectID);

        dataReader = new SqlDB("SELECT AEditLevel, AAllowAccountTransfer, AAllowAccountManagement, AAllowTeamspace, AAllowTeamspaceOnly, AEditLevelWorkshop FROM followup_profiles WHERE profileID='" + aProfileName + "'", SessionObj.Project.ProjectID);
        if (dataReader.read())
        {
            profileAccountManager.profileID = aProfileName;
            profileAccountManager.AEditLevel = dataReader.getInt32(0);
            profileAccountManager.AAllowAccountTransfer = dataReader.getBool(1);
            profileAccountManager.AAllowAccountManagement = dataReader.getBool(2);
            profileAccountManager.AAllowTeamspace = dataReader.getBool(3);
            profileAccountManager.AAllowTeamspaceOnly = dataReader.getBool(4);
            profileAccountManager.AEditLevelWorkshop = dataReader.getInt32(5);
        }
        return profileAccountManager;
    }
    public static TProfileSetContainer getProfileContainer(string aProfileName, string aProjectID)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        TProfileSetContainer profileContainer = new TProfileSetContainer();

        SqlDB dataReader = new SqlDB(aProjectID);

        dataReader = new SqlDB("SELECT allowUpload, allowDownload, allowDeleteOwnFiles, allowDeleteAllFiles, allowAccessOwnFilesWithoutPassword, allowAccessAllFilesWithoutPassword, allowCreateFolder, allowDeleteOwnFolder, allowDeleteAllFolder, allowResetPassword, allowTakeOwnership FROM teamspace_profilescontainer WHERE profileID='" + aProfileName + "'", SessionObj.Project.ProjectID);
        if (dataReader.read())
        {
            profileContainer.profileID = aProfileName;
            profileContainer.allowUpload = dataReader.getBool(0);
            profileContainer.allowDownload = dataReader.getBool(1);
            profileContainer.allowDeleteOwnFiles = dataReader.getBool(2);
            profileContainer.allowDeleteAllFiles = dataReader.getBool(3);
            profileContainer.allowAccessOwnFilesWithoutPassword = dataReader.getBool(4);
            profileContainer.allowAccessAllFilesWithoutPassword = dataReader.getBool(5);
            profileContainer.allowCreateFolder = dataReader.getBool(6);
            profileContainer.allowDeleteOwnFolder = dataReader.getBool(7);
            profileContainer.allowDeleteAllFolder = dataReader.getBool(8);
            profileContainer.allowResetPassword = dataReader.getBool(9);
            profileContainer.allowTakeOwnership = dataReader.getBool(10);
        }
        return profileContainer;
    }
    public static TProfileSetProjectplan getProfileProjectplan(string aProfileName, string aProjectID)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        TProfileSetProjectplan profileProjectplan = new TProfileSetProjectplan();

        SqlDB dataReader = new SqlDB(aProjectID);

        dataReader = new SqlDB("SELECT AllowTaskRead, AllowTaskCreateEditDelete, AllowTaskFileUpload, AllowTaskExport, AllowTaskImport FROM teamspace_profilestasks WHERE profileID='" + aProfileName + "'", SessionObj.Project.ProjectID);
        if (dataReader.read())
        {
            profileProjectplan.profileID = aProfileName;
            profileProjectplan.AllowTaskRead = dataReader.getBool(0);
            profileProjectplan.AllowTaskCreateEditDelete = dataReader.getBool(1);
            profileProjectplan.AllowTaskFileUpload = dataReader.getBool(2);
            profileProjectplan.AllowTaskExport = dataReader.getBool(3);
            profileProjectplan.AllowTaskImport = dataReader.getBool(4);
        }
        return profileProjectplan;
    }
    public static TProfileSetContacts getProfileContacts(string aProfileName, string aProjectID)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        TProfileSetContacts profileContacts = new TProfileSetContacts();

        SqlDB dataReader = new SqlDB(aProjectID);

        dataReader = new SqlDB("SELECT AllowContactRead, AllowContactCreateEditDelete, AllowContactExport, AllowContactImport FROM teamspace_profilescontacts WHERE profileID='" + aProfileName + "'", SessionObj.Project.ProjectID);
        if (dataReader.read())
        {
            profileContacts.profileID = aProfileName;
            profileContacts.AllowContactRead = dataReader.getBool(0);
            profileContacts.AllowContactCreateEditDelete = dataReader.getBool(1);
            profileContacts.AllowContactExport = dataReader.getBool(2);
            profileContacts.AllowContactImport = dataReader.getBool(3);
        }
        return profileContacts;
    }
    public static TProfileSetTranslate getProfileTranslate(string aProfileName, string aProjectID)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        TProfileSetTranslate profileTranslate = new TProfileSetTranslate();

        SqlDB dataReader = new SqlDB(aProjectID);

        dataReader = new SqlDB("SELECT AllowTranslateRead, AllowTranslateEdit, AllowTranslateRelease, AllowTranslateNew, AllowTranslateDelete FROM teamspace_profilestranslate WHERE profileID='" + aProfileName + "'", SessionObj.Project.ProjectID);
        if (dataReader.read())
        {
            profileTranslate.profileID = aProfileName;
            profileTranslate.AllowTranslateRead = dataReader.getBool(0);
            profileTranslate.AllowTranslateEdit = dataReader.getBool(1);
            profileTranslate.AllowTranslateRelease = dataReader.getBool(2);
            profileTranslate.AllowTranslateNew = dataReader.getBool(3);
            profileTranslate.AllowTranslateDelete = dataReader.getBool(4);
        }
        return profileTranslate;
    }

    public static bool saveProfileOrgManager(TProfileSetOrgManager aProfile, string aProjectID)
    {
        SqlDB dataReader;

        // ProfileID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("profileID", "string", aProfile.profileID);
        dataReader = new SqlDB("SELECT profileID FROM orgmanager_profiles WHERE profileID=@profileID", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            string tempDelegate = "0";
            if (aProfile.OAllowDelegate)
                tempDelegate = "1";
            string tempLock = "0";
            if (aProfile.OAllowLock)
                tempLock = "1";
            string tempEmployeeRead = "0";
            if (aProfile.OAllowEmployeeRead)
                tempEmployeeRead = "1";
            string tempEmployeeEdit = "0";
            if (aProfile.OAllowEmployeeEdit)
                tempEmployeeEdit = "1";
            string tempEmployeeImport = "0";
            if (aProfile.OAllowEmployeeImport)
                tempEmployeeImport = "1";
            string tempEmployeeExport = "0";
            if (aProfile.OAllowEmployeeExport)
                tempEmployeeExport = "1";
            string tempUnitAdd = "0";
            if (aProfile.OAllowUnitAdd)
                tempUnitAdd = "1";
            string tempUnitMove = "0";
            if (aProfile.OAllowUnitMove)
                tempUnitMove = "1";
            string tempUnitDelete = "0";
            if (aProfile.OAllowUnitDelete)
                tempUnitDelete = "1";
            string tempUnitProperty = "0";
            if (aProfile.OAllowUnitProperty)
                tempUnitProperty = "1";
            string tempReportRecipient = "0";
            if (aProfile.OAllowReportRecipient)
                tempReportRecipient = "1";
            string tempStructureImport = "0";
            if (aProfile.OAllowStructureImport)
                tempStructureImport = "1";
            string tempStructureExport = "0";
            if (aProfile.OAllowStructureExport)
                tempStructureExport = "1";
            string tempBouncesRead = "0";
            if (aProfile.OAllowBouncesRead)
                tempBouncesRead = "1";
            string tempBouncesEdit = "0";
            if (aProfile.OAllowBouncesEdit)
                tempBouncesEdit = "1";
            string tempBouncesDelete = "0";
            if (aProfile.OAllowBouncesDelete)
                tempBouncesDelete = "1";
            string tempBouncesExport = "0";
            if (aProfile.OAllowBouncesExport)
                tempBouncesExport = "1";
            string tempBouncesImport = "0";
            if (aProfile.OAllowBouncesImport)
                tempBouncesImport = "1";
            string tempAllowReInvitation = "0";
            if (aProfile.OAllowReInvitation)
                tempAllowReInvitation = "1";
            string tempAllowPostNomination = "0";
            if (aProfile.OAllowPostNomination)
                tempAllowPostNomination = "1";
            string tempAllowPostNominationImport = "0";
            if (aProfile.OAllowPostNominationImport)
                tempAllowPostNominationImport = "1";
            parameterList = new TParameterList();
            parameterList.addParameter("profileID", "string", aProfile.profileID);
            parameterList.addParameter("OReadLevel", "int", aProfile.OReadLevel.ToString());
            parameterList.addParameter("OEditLevel", "int", aProfile.OEditLevel.ToString());
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
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO orgmanager_profiles (profileID, OReadLevel, OEditLevel, OAllowDelegate, OAllowLock, OAllowEmployeeRead, OAllowEmployeeEdit, OAllowEmployeeImport, OAllowEmployeeExport, OAllowUnitAdd, OAllowUnitMove, OAllowUnitDelete, OAllowUnitProperty, OAllowReportRecipient, OAllowStructureImport, OAllowStructureExport, OAllowBouncesRead, OAllowBouncesEdit, OAllowBouncesDelete, OAllowBouncesExport, OAllowBouncesImport, OAllowReInvitation, OAllowPostNomination, OAllowPostNominationImport) VALUES (@profileID, @OReadLevel, @OEditLevel, @OAllowDelegate, @OAllowLock, @OAllowEmployeeRead, @OAllowEmployeeEdit, @OAllowEmployeeImport, @OAllowEmployeeExport, @OAllowUnitAdd, @OAllowUnitMove, @OAllowUnitDelete, @OAllowUnitProperty, @OAllowReportRecipient, @OAllowStructureImport, @OAllowStructureExport, @OAllowBouncesRead, @OAllowBouncesEdit, @OAllowBouncesDelete, @OAllowBouncesExport, @OAllowBouncesImport, @OAllowReInvitation, @OAllowPostNomination, @OAllowPostNominationImport)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveProfileOrgManagerBack(TProfileSetOrgManagerBack aProfile, string aProjectID)
    {
        SqlDB dataReader;

        // ProfileID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("profileID", "string", aProfile.profileID);
        dataReader = new SqlDB("SELECT profileID FROM orgmanager_profilesback WHERE profileID=@profileID", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            parameterList = new TParameterList();
            parameterList.addParameter("profileID", "string", aProfile.profileID);
            parameterList.addParameter("OReadLevel", "int", aProfile.OReadLevel.ToString());
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO orgmanager_profilesback (profileID, OReadLevel) VALUES (@profileID, @OReadLevel)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveProfileOrgManager1(TProfileSetOrgManager1 aProfile, string aProjectID)
    {
        SqlDB dataReader;

        // ProfileID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("profileID", "string", aProfile.profileID);
        dataReader = new SqlDB("SELECT profileID FROM orgmanager_profiles1 WHERE profileID=@profileID", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            string tempDelegate = "0";
            if (aProfile.OAllowDelegate)
                tempDelegate = "1";
            string tempLock = "0";
            if (aProfile.OAllowLock)
                tempLock = "1";
            string tempEmployeeRead = "0";
            if (aProfile.OAllowEmployeeRead)
                tempEmployeeRead = "1";
            string tempEmployeeEdit = "0";
            if (aProfile.OAllowEmployeeEdit)
                tempEmployeeEdit = "1";
            string tempEmployeeImport = "0";
            if (aProfile.OAllowEmployeeImport)
                tempEmployeeImport = "1";
            string tempEmployeeExport = "0";
            if (aProfile.OAllowEmployeeExport)
                tempEmployeeExport = "1";
            string tempUnitAdd = "0";
            if (aProfile.OAllowUnitAdd)
                tempUnitAdd = "1";
            string tempUnitMove = "0";
            if (aProfile.OAllowUnitMove)
                tempUnitMove = "1";
            string tempUnitDelete = "0";
            if (aProfile.OAllowUnitDelete)
                tempUnitDelete = "1";
            string tempUnitProperty = "0";
            if (aProfile.OAllowUnitProperty)
                tempUnitProperty = "1";
            string tempReportRecipient = "0";
            if (aProfile.OAllowReportRecipient)
                tempReportRecipient = "1";
            string tempStructureImport = "0";
            if (aProfile.OAllowStructureImport)
                tempStructureImport = "1";
            string tempStructureExport = "0";
            if (aProfile.OAllowStructureExport)
                tempStructureExport = "1";
            parameterList = new TParameterList();
            parameterList.addParameter("profileID", "string", aProfile.profileID);
            parameterList.addParameter("OReadLevel", "int", aProfile.OReadLevel.ToString());
            parameterList.addParameter("OEditLevel", "int", aProfile.OEditLevel.ToString());
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
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO orgmanager_profiles1 (profileID, OReadLevel, OEditLevel, OAllowDelegate, OAllowLock, OAllowEmployeeRead, OAllowEmployeeEdit, OAllowEmployeeImport, OAllowEmployeeExport, OAllowUnitAdd, OAllowUnitMove, OAllowUnitDelete, OAllowUnitProperty, OAllowReportRecipient, OAllowStructureImport, OAllowStructureExport) VALUES (@profileID, @OReadLevel, @OEditLevel, @OAllowDelegate, @OAllowLock, @OAllowEmployeeRead, @OAllowEmployeeEdit, @OAllowEmployeeImport, @OAllowEmployeeExport, @OAllowUnitAdd, @OAllowUnitMove, @OAllowUnitDelete, @OAllowUnitProperty, @OAllowReportRecipient, @OAllowStructureImport, @OAllowStructureExport)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveProfileOrgManager1Back(TProfileSetOrgManager1Back aProfile, string aProjectID)
    {
        SqlDB dataReader;

        // ProfileID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("profileID", "string", aProfile.profileID);
        dataReader = new SqlDB("SELECT profileID FROM orgmanager_profiles1back WHERE profileID=@profileID", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            parameterList = new TParameterList();
            parameterList.addParameter("profileID", "string", aProfile.profileID);
            parameterList.addParameter("OReadLevel", "int", aProfile.OReadLevel.ToString());
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO orgmanager_profiles1back (profileID, OReadLevel) VALUES (@profileID, @OReadLevel)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveProfileOrgManager2(TProfileSetOrgManager2 aProfile, string aProjectID)
    {
        SqlDB dataReader;

        // ProfileID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("profileID", "string", aProfile.profileID);
        dataReader = new SqlDB("SELECT profileID FROM orgmanager_profiles2 WHERE profileID=@profileID", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            string tempDelegate = "0";
            if (aProfile.OAllowDelegate)
                tempDelegate = "1";
            string tempLock = "0";
            if (aProfile.OAllowLock)
                tempLock = "1";
            string tempEmployeeRead = "0";
            if (aProfile.OAllowEmployeeRead)
                tempEmployeeRead = "1";
            string tempEmployeeEdit = "0";
            if (aProfile.OAllowEmployeeEdit)
                tempEmployeeEdit = "1";
            string tempEmployeeImport = "0";
            if (aProfile.OAllowEmployeeImport)
                tempEmployeeImport = "1";
            string tempEmployeeExport = "0";
            if (aProfile.OAllowEmployeeExport)
                tempEmployeeExport = "1";
            string tempUnitAdd = "0";
            if (aProfile.OAllowUnitAdd)
                tempUnitAdd = "1";
            string tempUnitMove = "0";
            if (aProfile.OAllowUnitMove)
                tempUnitMove = "1";
            string tempUnitDelete = "0";
            if (aProfile.OAllowUnitDelete)
                tempUnitDelete = "1";
            string tempUnitProperty = "0";
            if (aProfile.OAllowUnitProperty)
                tempUnitProperty = "1";
            string tempReportRecipient = "0";
            if (aProfile.OAllowReportRecipient)
                tempReportRecipient = "1";
            string tempStructureImport = "0";
            if (aProfile.OAllowStructureImport)
                tempStructureImport = "1";
            string tempStructureExport = "0";
            if (aProfile.OAllowStructureExport)
                tempStructureExport = "1";
            parameterList = new TParameterList();
            parameterList.addParameter("profileID", "string", aProfile.profileID);
            parameterList.addParameter("OReadLevel", "int", aProfile.OReadLevel.ToString());
            parameterList.addParameter("OEditLevel", "int", aProfile.OEditLevel.ToString());
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
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO orgmanager_profiles2 (profileID, OReadLevel, OEditLevel, OAllowDelegate, OAllowLock, OAllowEmployeeRead, OAllowEmployeeEdit, OAllowEmployeeImport, OAllowEmployeeExport, OAllowUnitAdd, OAllowUnitMove, OAllowUnitDelete, OAllowUnitProperty, OAllowReportRecipient, OAllowStructureImport, OAllowStructureExport) VALUES (@profileID, @OReadLevel, @OEditLevel, @OAllowDelegate, @OAllowLock, @OAllowEmployeeRead, @OAllowEmployeeEdit, @OAllowEmployeeImport, @OAllowEmployeeExport, @OAllowUnitAdd, @OAllowUnitMove, @OAllowUnitDelete, @OAllowUnitProperty, @OAllowReportRecipient, @OAllowStructureImport, @OAllowStructureExport)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveProfileOrgManager2Back(TProfileSetOrgManager2Back aProfile, string aProjectID)
    {
        SqlDB dataReader;

        // ProfileID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("profileID", "string", aProfile.profileID);
        dataReader = new SqlDB("SELECT profileID FROM orgmanager_profiles2back WHERE profileID=@profileID", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            parameterList = new TParameterList();
            parameterList.addParameter("profileID", "string", aProfile.profileID);
            parameterList.addParameter("OReadLevel", "int", aProfile.OReadLevel.ToString());
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO orgmanager_profiles2back (profileID, OReadLevel) VALUES (@profileID, @OReadLevel)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveProfileOrgManager3(TProfileSetOrgManager3 aProfile, string aProjectID)
    {
        SqlDB dataReader;

        // ProfileID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("profileID", "string", aProfile.profileID);
        dataReader = new SqlDB("SELECT profileID FROM orgmanager_profiles3 WHERE profileID=@profileID", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            string tempDelegate = "0";
            if (aProfile.OAllowDelegate)
                tempDelegate = "1";
            string tempLock = "0";
            if (aProfile.OAllowLock)
                tempLock = "1";
            string tempEmployeeRead = "0";
            if (aProfile.OAllowEmployeeRead)
                tempEmployeeRead = "1";
            string tempEmployeeEdit = "0";
            if (aProfile.OAllowEmployeeEdit)
                tempEmployeeEdit = "1";
            string tempEmployeeImport = "0";
            if (aProfile.OAllowEmployeeImport)
                tempEmployeeImport = "1";
            string tempEmployeeExport = "0";
            if (aProfile.OAllowEmployeeExport)
                tempEmployeeExport = "1";
            string tempUnitAdd = "0";
            if (aProfile.OAllowUnitAdd)
                tempUnitAdd = "1";
            string tempUnitMove = "0";
            if (aProfile.OAllowUnitMove)
                tempUnitMove = "1";
            string tempUnitDelete = "0";
            if (aProfile.OAllowUnitDelete)
                tempUnitDelete = "1";
            string tempUnitProperty = "0";
            if (aProfile.OAllowUnitProperty)
                tempUnitProperty = "1";
            string tempReportRecipient = "0";
            if (aProfile.OAllowReportRecipient)
                tempReportRecipient = "1";
            string tempStructureImport = "0";
            if (aProfile.OAllowStructureImport)
                tempStructureImport = "1";
            string tempStructureExport = "0";
            if (aProfile.OAllowStructureExport)
                tempStructureExport = "1";
            parameterList = new TParameterList();
            parameterList.addParameter("profileID", "string", aProfile.profileID);
            parameterList.addParameter("OReadLevel", "int", aProfile.OReadLevel.ToString());
            parameterList.addParameter("OEditLevel", "int", aProfile.OEditLevel.ToString());
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
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO orgmanager_profiles3 (profileID, OReadLevel, OEditLevel, OAllowDelegate, OAllowLock, OAllowEmployeeRead, OAllowEmployeeEdit, OAllowEmployeeImport, OAllowEmployeeExport, OAllowUnitAdd, OAllowUnitMove, OAllowUnitDelete, OAllowUnitProperty, OAllowReportRecipient, OAllowStructureImport, OAllowStructureExport) VALUES (@profileID, @OReadLevel, @OEditLevel, @OAllowDelegate, @OAllowLock, @OAllowEmployeeRead, @OAllowEmployeeEdit, @OAllowEmployeeImport, @OAllowEmployeeExport, @OAllowUnitAdd, @OAllowUnitMove, @OAllowUnitDelete, @OAllowUnitProperty, @OAllowReportRecipient, @OAllowStructureImport, @OAllowStructureExport)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveProfileOrgManager3Back(TProfileSetOrgManager3Back aProfile, string aProjectID)
    {
        SqlDB dataReader;

        // ProfileID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("profileID", "string", aProfile.profileID);
        dataReader = new SqlDB("SELECT profileID FROM orgmanager_profiles3back WHERE profileID=@profileID", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            parameterList = new TParameterList();
            parameterList.addParameter("profileID", "string", aProfile.profileID);
            parameterList.addParameter("OReadLevel", "int", aProfile.OReadLevel.ToString());
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO orgmanager_profiles3back (profileID, OReadLevel) VALUES (@profileID, @OReadLevel)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveProfileResponse(TProfileSetResponse aProfile, string aProjectID)
    {
        SqlDB dataReader;

        // ProfileID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("profileID", "string", aProfile.profileID);
        dataReader = new SqlDB("SELECT profileID FROM response_profiles WHERE profileID=@profileID", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            parameterList = new TParameterList();
            parameterList.addParameter("profileID", "string", aProfile.profileID);
            parameterList.addParameter("RReadLevel", "int", aProfile.RReadLevel.ToString());
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO response_profiles (profileID, RReadLevel) VALUES (@profileID, @RReadLevel)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveProfileResponse1(TProfileSetResponse1 aProfile, string aProjectID)
    {
        SqlDB dataReader;

        // ProfileID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("profileID", "string", aProfile.profileID);
        dataReader = new SqlDB("SELECT profileID FROM response_profiles1 WHERE profileID=@profileID", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            parameterList = new TParameterList();
            parameterList.addParameter("profileID", "string", aProfile.profileID);
            parameterList.addParameter("RReadLevel", "int", aProfile.RReadLevel.ToString());
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO response_profiles1 (profileID, RReadLevel) VALUES (@profileID, @RReadLevel)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveProfileResponse2(TProfileSetResponse2 aProfile, string aProjectID)
    {
        SqlDB dataReader;

        // ProfileID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("profileID", "string", aProfile.profileID);
        dataReader = new SqlDB("SELECT profileID FROM response_profiles2 WHERE profileID=@profileID", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            parameterList = new TParameterList();
            parameterList.addParameter("profileID", "string", aProfile.profileID);
            parameterList.addParameter("RReadLevel", "int", aProfile.RReadLevel.ToString());
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO response_profiles2 (profileID, RReadLevel) VALUES (@profileID, @RReadLevel)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveProfileResponse3(TProfileSetResponse3 aProfile, string aProjectID)
    {
        SqlDB dataReader;

        // ProfileID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("profileID", "string", aProfile.profileID);
        dataReader = new SqlDB("SELECT profileID FROM response_profiles3 WHERE profileID=@profileID", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            parameterList = new TParameterList();
            parameterList.addParameter("profileID", "string", aProfile.profileID);
            parameterList.addParameter("RReadLevel", "int", aProfile.RReadLevel.ToString());
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO response_profiles3 (profileID, RReadLevel) VALUES (@profileID, @RReadLevel)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveProfileDownload(TProfileSetDownload aProfile, string aProjectID)
    {
        SqlDB dataReader;

        // ProfileID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("profileID", "string", aProfile.profileID);
        dataReader = new SqlDB("SELECT profileID FROM download_profiles WHERE profileID=@profileID", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            string tempMail = "0";
            if (aProfile.DAllowMail)
                tempMail = "1";
            string tempDownload = "0";
            if (aProfile.DAllowDownload)
                tempDownload = "1";
            string tempLog = "0";
            if (aProfile.DAllowLog)
                tempLog = "1";
            string tempType1 = "0";
            if (aProfile.DAllowType1)
                tempType1 = "1";
            string tempType2 = "0";
            if (aProfile.DAllowType2)
                tempType2 = "1";
            string tempType3 = "0";
            if (aProfile.DAllowType3)
                tempType3 = "1";
            string tempType4 = "0";
            if (aProfile.DAllowType4)
                tempType4 = "1";
            string tempType5 = "0";
            if (aProfile.DAllowType5)
                tempType5 = "1";
            parameterList = new TParameterList();
            parameterList.addParameter("profileID", "string", aProfile.profileID);
            parameterList.addParameter("DReadLevel", "int", aProfile.DReadLevel.ToString());
            parameterList.addParameter("DAllowMail", "int", tempMail);
            parameterList.addParameter("DAllowDownload", "int", tempDownload);
            parameterList.addParameter("DAllowLog", "int", tempLog);
            parameterList.addParameter("DAllowType1", "int", tempType1);
            parameterList.addParameter("DAllowType2", "int", tempType2);
            parameterList.addParameter("DAllowType3", "int", tempType3);
            parameterList.addParameter("DAllowType4", "int", tempType4);
            parameterList.addParameter("DAllowType5", "int", tempType5);
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO download_profiles (profileID, DReadLevel, DAllowMail, DAllowDownload, DAllowLog, DAllowType1, DAllowType2, DAllowType3, DAllowType4, DAllowType5) VALUES (@profileID, @DReadLevel, @DAllowMail, @DAllowDownload, @DAllowLog, @DAllowType1, @DAllowType2, @DAllowType3, @DAllowType4, @DAllowType5)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveProfileDownload1(TProfileSetDownload1 aProfile, string aProjectID)
    {
        SqlDB dataReader;

        // ProfileID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("profileID", "string", aProfile.profileID);
        dataReader = new SqlDB("SELECT profileID FROM download_profiles1 WHERE profileID=@profileID", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            string tempMail = "0";
            if (aProfile.DAllowMail)
                tempMail = "1";
            string tempDownload = "0";
            if (aProfile.DAllowDownload)
                tempDownload = "1";
            string tempLog = "0";
            if (aProfile.DAllowLog)
                tempLog = "1";
            string tempType1 = "0";
            if (aProfile.DAllowType1)
                tempType1 = "1";
            string tempType2 = "0";
            if (aProfile.DAllowType2)
                tempType2 = "1";
            string tempType3 = "0";
            if (aProfile.DAllowType3)
                tempType3 = "1";
            string tempType4 = "0";
            if (aProfile.DAllowType4)
                tempType4 = "1";
            string tempType5 = "0";
            if (aProfile.DAllowType5)
                tempType5 = "1";
            parameterList = new TParameterList();
            parameterList.addParameter("profileID", "string", aProfile.profileID);
            parameterList.addParameter("DReadLevel", "int", aProfile.DReadLevel.ToString());
            parameterList.addParameter("DAllowMail", "int", tempMail);
            parameterList.addParameter("DAllowDownload", "int", tempDownload);
            parameterList.addParameter("DAllowLog", "int", tempLog);
            parameterList.addParameter("DAllowType1", "int", tempType1);
            parameterList.addParameter("DAllowType2", "int", tempType2);
            parameterList.addParameter("DAllowType3", "int", tempType3);
            parameterList.addParameter("DAllowType4", "int", tempType4);
            parameterList.addParameter("DAllowType5", "int", tempType5);
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO download_profiles1 (profileID, DReadLevel, DAllowMail, DAllowDownload, DAllowLog, DAllowType1, DAllowType2, DAllowType3, DAllowType4, DAllowType5) VALUES (@profileID, @DReadLevel, @DAllowMail, @DAllowDownload, @DAllowLog, @DAllowType1, @DAllowType2, @DAllowType3, @DAllowType4, @DAllowType5)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveProfileDownload2(TProfileSetDownload2 aProfile, string aProjectID)
    {
        SqlDB dataReader;

        // ProfileID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("profileID", "string", aProfile.profileID);
        dataReader = new SqlDB("SELECT profileID FROM download_profiles2 WHERE profileID=@profileID", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            string tempMail = "0";
            if (aProfile.DAllowMail)
                tempMail = "1";
            string tempDownload = "0";
            if (aProfile.DAllowDownload)
                tempDownload = "1";
            string tempLog = "0";
            if (aProfile.DAllowLog)
                tempLog = "1";
            string tempType1 = "0";
            if (aProfile.DAllowType1)
                tempType1 = "1";
            string tempType2 = "0";
            if (aProfile.DAllowType2)
                tempType2 = "1";
            string tempType3 = "0";
            if (aProfile.DAllowType3)
                tempType3 = "1";
            string tempType4 = "0";
            if (aProfile.DAllowType4)
                tempType4 = "1";
            string tempType5 = "0";
            if (aProfile.DAllowType5)
                tempType5 = "1";
            parameterList = new TParameterList();
            parameterList.addParameter("profileID", "string", aProfile.profileID);
            parameterList.addParameter("DReadLevel", "int", aProfile.DReadLevel.ToString());
            parameterList.addParameter("DAllowMail", "int", tempMail);
            parameterList.addParameter("DAllowDownload", "int", tempDownload);
            parameterList.addParameter("DAllowLog", "int", tempLog);
            parameterList.addParameter("DAllowType1", "int", tempType1);
            parameterList.addParameter("DAllowType2", "int", tempType2);
            parameterList.addParameter("DAllowType3", "int", tempType3);
            parameterList.addParameter("DAllowType4", "int", tempType4);
            parameterList.addParameter("DAllowType5", "int", tempType5);
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO download_profiles2 (profileID, DReadLevel, DAllowMail, DAllowDownload, DAllowLog, DAllowType1, DAllowType2, DAllowType3, DAllowType4, DAllowType5) VALUES (@profileID, @DReadLevel, @DAllowMail, @DAllowDownload, @DAllowLog, @DAllowType1, @DAllowType2, @DAllowType3, @DAllowType4, @DAllowType5)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveProfileDownload3(TProfileSetDownload3 aProfile, string aProjectID)
    {
        SqlDB dataReader;

        // ProfileID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("profileID", "string", aProfile.profileID);
        dataReader = new SqlDB("SELECT profileID FROM download_profiles3 WHERE profileID=@profileID", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            string tempMail = "0";
            if (aProfile.DAllowMail)
                tempMail = "1";
            string tempDownload = "0";
            if (aProfile.DAllowDownload)
                tempDownload = "1";
            string tempLog = "0";
            if (aProfile.DAllowLog)
                tempLog = "1";
            string tempType1 = "0";
            if (aProfile.DAllowType1)
                tempType1 = "1";
            string tempType2 = "0";
            if (aProfile.DAllowType2)
                tempType2 = "1";
            string tempType3 = "0";
            if (aProfile.DAllowType3)
                tempType3 = "1";
            string tempType4 = "0";
            if (aProfile.DAllowType4)
                tempType4 = "1";
            string tempType5 = "0";
            if (aProfile.DAllowType5)
                tempType5 = "1";
            parameterList = new TParameterList();
            parameterList.addParameter("profileID", "string", aProfile.profileID);
            parameterList.addParameter("DReadLevel", "int", aProfile.DReadLevel.ToString());
            parameterList.addParameter("DAllowMail", "int", tempMail);
            parameterList.addParameter("DAllowDownload", "int", tempDownload);
            parameterList.addParameter("DAllowLog", "int", tempLog);
            parameterList.addParameter("DAllowType1", "int", tempType1);
            parameterList.addParameter("DAllowType2", "int", tempType2);
            parameterList.addParameter("DAllowType3", "int", tempType3);
            parameterList.addParameter("DAllowType4", "int", tempType4);
            parameterList.addParameter("DAllowType5", "int", tempType5);
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO download_profiles3 (profileID, DReadLevel, DAllowMail, DAllowDownload, DAllowLog, DAllowType1, DAllowType2, DAllowType3, DAllowType4, DAllowType5) VALUES (@profileID, @DReadLevel, @DAllowMail, @DAllowDownload, @DAllowLog, @DAllowType1, @DAllowType2, @DAllowType3, @DAllowType4, @DAllowType5)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveProfileFollowUp(TProfileSetFollowUp aProfile, string aProjectID)
    {
        SqlDB dataReader;

        // ProfileID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("profileID", "string", aProfile.profileID);
        dataReader = new SqlDB("SELECT profileID FROM followup_profiles WHERE profileID=@profileID", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            string tempCommunications = "0";
            if (aProfile.FAllowCommunications)
                tempCommunications = "1";
            string tempMeasures = "0";
            if (aProfile.FAllowMeasures)
                tempMeasures = "1";
            string tempDelete = "0";
            if (aProfile.FAllowDelete)
                tempDelete = "1";
            string tempExcelExport = "0";
            if (aProfile.FAllowExcelExport)
                tempExcelExport = "1";
            string tempReminderOnMeasure = "0";
            if (aProfile.FAllowReminderOnMeasure)
                tempReminderOnMeasure = "1";
            string tempReminderOnUnit = "0";
            if (aProfile.FAllowReminderOnUnit)
                tempReminderOnUnit = "1";
            string tempUpUnitsMeasures = "0";
            if (aProfile.FAllowUpUnitsMeasures)
                tempUpUnitsMeasures = "1";
            string tempColleaguesMeasures = "0";
            if (aProfile.FAllowColleaguesMeasures)
                tempColleaguesMeasures = "1";
            string tempTopUnitsThemes = "0";
            if (aProfile.FAllowTopUnitsThemes)
                tempTopUnitsThemes = "1";
            string tempColleaguesThemes = "0";
            if (aProfile.FAllowColleaguesThemes)
                tempColleaguesThemes = "1";
            string tempThemes = "0";
            if (aProfile.FAllowThemes)
                tempThemes = "1";
            string tempReminderOnTheme = "0";
            if (aProfile.FAllowReminderOnTheme)
                tempReminderOnTheme = "1";
            string tempReminderOnThemeUnit = "0";
            if (aProfile.FAllowReminderOnThemeUnit)
                tempReminderOnThemeUnit = "1";
            string tempOrders = "0";
            if (aProfile.FAllowOrders)
                tempOrders = "1";
            string tempReminderOnOrderUnit = "0";
            if (aProfile.FAllowReminderOnOrderUnit)
                tempReminderOnOrderUnit = "1";
            parameterList = new TParameterList();
            parameterList.addParameter("profileID", "string", aProfile.profileID);
            parameterList.addParameter("FReadLevel", "int", aProfile.FReadLevel.ToString());
            parameterList.addParameter("FEditLevel", "int", aProfile.FEditLevel.ToString());
            parameterList.addParameter("FSumLevel", "int", aProfile.FSumLevel.ToString());
            parameterList.addParameter("FAllowCommunications", "int", tempCommunications);
            parameterList.addParameter("FAllowMeasures", "int", tempMeasures);
            parameterList.addParameter("FAllowDelete", "int", tempDelete);
            parameterList.addParameter("FAllowExcelExport", "int", tempExcelExport);
            parameterList.addParameter("FAllowReminderOnMeasure", "int", tempReminderOnMeasure);
            parameterList.addParameter("FAllowReminderOnUnit", "int", tempReminderOnUnit);
            parameterList.addParameter("FAllowUpUnitsMeasures", "int", tempUpUnitsMeasures);
            parameterList.addParameter("FAllowColleaguesMeasures", "int", tempColleaguesMeasures);
            parameterList.addParameter("FAllowTopUnitsThemes", "int", tempTopUnitsThemes);
            parameterList.addParameter("FAllowColleaguesThemes", "int", tempColleaguesThemes);
            parameterList.addParameter("FAllowThemes", "int", tempThemes);
            parameterList.addParameter("FAllowReminderOnTheme", "int", tempReminderOnTheme);
            parameterList.addParameter("FAllowReminderOnThemeUnit", "int", tempReminderOnThemeUnit);
            parameterList.addParameter("FAllowOrders", "int", tempOrders);
            parameterList.addParameter("FAllowReminderOnOrderUnit", "int", tempReminderOnOrderUnit);
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO followup_profiles (profileID, FReadLevel, FEditLevel, FSumLevel, FAllowCommunications, FAllowMeasures, FAllowDelete, FAllowExcelExport, FAllowReminderOnMeasure, FAllowReminderOnUnit, FAllowUpUnitsMeasures, FAllowColleaguesMeasures, FAllowTopUnitsThemes, FAllowColleaguesThemes, FAllowThemes, FAllowReminderOnTheme, FAllowReminderOnThemeUnit, FAllowOrders, FAllowReminderOnOrderUnit) VALUES (@profileID, @FReadLevel, @FEditLevel, @FSumLevel, @FAllowCommunications, @FAllowMeasures, @FAllowDelete, @FAllowExcelExport, @FAllowReminderOnMeasure, @FAllowReminderOnUnit, @FAllowUpUnitsMeasures, @FAllowColleaguesMeasures, @FAllowTopUnitsThemes, @FAllowColleaguesThemes, @FAllowThemes, @FAllowReminderOnTheme, @FAllowReminderOnThemeUnit, @FAllowOrders, @FAllowReminderOnOrderUnit)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveProfileFollowUp1(TProfileSetFollowUp1 aProfile, string aProjectID)
    {
        SqlDB dataReader;

        // ProfileID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("profileID", "string", aProfile.profileID);
        dataReader = new SqlDB("SELECT profileID FROM followup_profiles1 WHERE profileID=@profileID", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            string tempCommunications = "0";
            if (aProfile.FAllowCommunications)
                tempCommunications = "1";
            string tempMeasures = "0";
            if (aProfile.FAllowMeasures)
                tempMeasures = "1";
            string tempDelete = "0";
            if (aProfile.FAllowDelete)
                tempDelete = "1";
            string tempExcelExport = "0";
            if (aProfile.FAllowExcelExport)
                tempExcelExport = "1";
            string tempReminderOnMeasure = "0";
            if (aProfile.FAllowReminderOnMeasure)
                tempReminderOnMeasure = "1";
            string tempReminderOnUnit = "0";
            if (aProfile.FAllowReminderOnUnit)
                tempReminderOnUnit = "1";
            string tempUpUnitsMeasures = "0";
            if (aProfile.FAllowUpUnitsMeasures)
                tempUpUnitsMeasures = "1";
            string tempColleaguesMeasures = "0";
            if (aProfile.FAllowColleaguesMeasures)
                tempColleaguesMeasures = "1";
            string tempTopUnitsThemes = "0";
            if (aProfile.FAllowTopUnitsThemes)
                tempTopUnitsThemes = "1";
            string tempColleaguesThemes = "0";
            if (aProfile.FAllowColleaguesThemes)
                tempColleaguesThemes = "1";
            string tempThemes = "0";
            if (aProfile.FAllowThemes)
                tempThemes = "1";
            string tempReminderOnTheme = "0";
            if (aProfile.FAllowReminderOnTheme)
                tempReminderOnTheme = "1";
            string tempReminderOnThemeUnit = "0";
            if (aProfile.FAllowReminderOnThemeUnit)
                tempReminderOnThemeUnit = "1";
            string tempOrders = "0";
            if (aProfile.FAllowOrders)
                tempOrders = "1";
            string tempReminderOnOrderUnit = "0";
            if (aProfile.FAllowReminderOnOrderUnit)
                tempReminderOnOrderUnit = "1";
            parameterList = new TParameterList();
            parameterList.addParameter("profileID", "string", aProfile.profileID);
            parameterList.addParameter("FReadLevel", "int", aProfile.FReadLevel.ToString());
            parameterList.addParameter("FEditLevel", "int", aProfile.FEditLevel.ToString());
            parameterList.addParameter("FSumLevel", "int", aProfile.FSumLevel.ToString());
            parameterList.addParameter("FAllowCommunications", "int", tempCommunications);
            parameterList.addParameter("FAllowMeasures", "int", tempMeasures);
            parameterList.addParameter("FAllowDelete", "int", tempDelete);
            parameterList.addParameter("FAllowExcelExport", "int", tempExcelExport);
            parameterList.addParameter("FAllowReminderOnMeasure", "int", tempReminderOnMeasure);
            parameterList.addParameter("FAllowReminderOnUnit", "int", tempReminderOnUnit);
            parameterList.addParameter("FAllowUpUnitsMeasures", "int", tempUpUnitsMeasures);
            parameterList.addParameter("FAllowColleaguesMeasures", "int", tempColleaguesMeasures);
            parameterList.addParameter("FAllowTopUnitsThemes", "int", tempTopUnitsThemes);
            parameterList.addParameter("FAllowColleaguesThemes", "int", tempColleaguesThemes);
            parameterList.addParameter("FAllowThemes", "int", tempThemes);
            parameterList.addParameter("FAllowReminderOnTheme", "int", tempReminderOnTheme);
            parameterList.addParameter("FAllowReminderOnThemeUnit", "int", tempReminderOnThemeUnit);
            parameterList.addParameter("FAllowOrders", "int", tempOrders);
            parameterList.addParameter("FAllowReminderOnOrderUnit", "int", tempReminderOnOrderUnit);
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO followup_profiles1 (profileID, FReadLevel, FEditLevel, FSumLevel, FAllowCommunications, FAllowMeasures, FAllowDelete, FAllowExcelExport, FAllowReminderOnMeasure, FAllowReminderOnUnit, FAllowUpUnitsMeasures, FAllowColleaguesMeasures, FAllowTopUnitsThemes, FAllowColleaguesThemes, FAllowThemes, FAllowReminderOnTheme, FAllowReminderOnThemeUnit, FAllowOrders, FAllowReminderOnOrderUnit) VALUES (@profileID, @FReadLevel, @FEditLevel, @FSumLevel, @FAllowCommunications, @FAllowMeasures, @FAllowDelete, @FAllowExcelExport, @FAllowReminderOnMeasure, @FAllowReminderOnUnit, @FAllowUpUnitsMeasures, @FAllowColleaguesMeasures, @FAllowTopUnitsThemes, @FAllowColleaguesThemes, @FAllowThemes, @FAllowReminderOnTheme, @FAllowReminderOnThemeUnit, @FAllowOrders, @FAllowReminderOnOrderUnit)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveProfileFollowUp2(TProfileSetFollowUp2 aProfile, string aProjectID)
    {
        SqlDB dataReader;

        // ProfileID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("profileID", "string", aProfile.profileID);
        dataReader = new SqlDB("SELECT profileID FROM followup_profiles2 WHERE profileID=@profileID", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            string tempCommunications = "0";
            if (aProfile.FAllowCommunications)
                tempCommunications = "1";
            string tempMeasures = "0";
            if (aProfile.FAllowMeasures)
                tempMeasures = "1";
            string tempDelete = "0";
            if (aProfile.FAllowDelete)
                tempDelete = "1";
            string tempExcelExport = "0";
            if (aProfile.FAllowExcelExport)
                tempExcelExport = "1";
            string tempReminderOnMeasure = "0";
            if (aProfile.FAllowReminderOnMeasure)
                tempReminderOnMeasure = "1";
            string tempReminderOnUnit = "0";
            if (aProfile.FAllowReminderOnUnit)
                tempReminderOnUnit = "1";
            string tempUpUnitsMeasures = "0";
            if (aProfile.FAllowUpUnitsMeasures)
                tempUpUnitsMeasures = "1";
            string tempColleaguesMeasures = "0";
            if (aProfile.FAllowColleaguesMeasures)
                tempColleaguesMeasures = "1";
            string tempTopUnitsThemes = "0";
            if (aProfile.FAllowTopUnitsThemes)
                tempTopUnitsThemes = "1";
            string tempColleaguesThemes = "0";
            if (aProfile.FAllowColleaguesThemes)
                tempColleaguesThemes = "1";
            string tempThemes = "0";
            if (aProfile.FAllowThemes)
                tempThemes = "1";
            string tempReminderOnTheme = "0";
            if (aProfile.FAllowReminderOnTheme)
                tempReminderOnTheme = "1";
            string tempReminderOnThemeUnit = "0";
            if (aProfile.FAllowReminderOnThemeUnit)
                tempReminderOnThemeUnit = "1";
            string tempOrders = "0";
            if (aProfile.FAllowOrders)
                tempOrders = "1";
            string tempReminderOnOrderUnit = "0";
            if (aProfile.FAllowReminderOnOrderUnit)
                tempReminderOnOrderUnit = "1";
            parameterList = new TParameterList();
            parameterList.addParameter("profileID", "string", aProfile.profileID);
            parameterList.addParameter("FReadLevel", "int", aProfile.FReadLevel.ToString());
            parameterList.addParameter("FEditLevel", "int", aProfile.FEditLevel.ToString());
            parameterList.addParameter("FSumLevel", "int", aProfile.FSumLevel.ToString());
            parameterList.addParameter("FAllowCommunications", "int", tempCommunications);
            parameterList.addParameter("FAllowMeasures", "int", tempMeasures);
            parameterList.addParameter("FAllowDelete", "int", tempDelete);
            parameterList.addParameter("FAllowExcelExport", "int", tempExcelExport);
            parameterList.addParameter("FAllowReminderOnMeasure", "int", tempReminderOnMeasure);
            parameterList.addParameter("FAllowReminderOnUnit", "int", tempReminderOnUnit);
            parameterList.addParameter("FAllowUpUnitsMeasures", "int", tempUpUnitsMeasures);
            parameterList.addParameter("FAllowColleaguesMeasures", "int", tempColleaguesMeasures);
            parameterList.addParameter("FAllowTopUnitsThemes", "int", tempTopUnitsThemes);
            parameterList.addParameter("FAllowColleaguesThemes", "int", tempColleaguesThemes);
            parameterList.addParameter("FAllowThemes", "int", tempThemes);
            parameterList.addParameter("FAllowReminderOnTheme", "int", tempReminderOnTheme);
            parameterList.addParameter("FAllowReminderOnThemeUnit", "int", tempReminderOnThemeUnit);
            parameterList.addParameter("FAllowOrders", "int", tempOrders);
            parameterList.addParameter("FAllowReminderOnOrderUnit", "int", tempReminderOnOrderUnit);
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO followup_profiles2 (profileID, FReadLevel, FEditLevel, FSumLevel, FAllowCommunications, FAllowMeasures, FAllowDelete, FAllowExcelExport, FAllowReminderOnMeasure, FAllowReminderOnUnit, FAllowUpUnitsMeasures, FAllowColleaguesMeasures, FAllowTopUnitsThemes, FAllowColleaguesThemes, FAllowThemes, FAllowReminderOnTheme, FAllowReminderOnThemeUnit, FAllowOrders, FAllowReminderOnOrderUnit) VALUES (@profileID, @FReadLevel, @FEditLevel, @FSumLevel, @FAllowCommunications, @FAllowMeasures, @FAllowDelete, @FAllowExcelExport, @FAllowReminderOnMeasure, @FAllowReminderOnUnit, @FAllowUpUnitsMeasures, @FAllowColleaguesMeasures, @FAllowTopUnitsThemes, @FAllowColleaguesThemes, @FAllowThemes, @FAllowReminderOnTheme, @FAllowReminderOnThemeUnit, @FAllowOrders, @FAllowReminderOnOrderUnit)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveProfileFollowUp3(TProfileSetFollowUp3 aProfile, string aProjectID)
    {
        SqlDB dataReader;

        // ProfileID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("profileID", "string", aProfile.profileID);
        dataReader = new SqlDB("SELECT profileID FROM followup_profiles3 WHERE profileID=@profileID", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            string tempCommunications = "0";
            if (aProfile.FAllowCommunications)
                tempCommunications = "1";
            string tempMeasures = "0";
            if (aProfile.FAllowMeasures)
                tempMeasures = "1";
            string tempDelete = "0";
            if (aProfile.FAllowDelete)
                tempDelete = "1";
            string tempExcelExport = "0";
            if (aProfile.FAllowExcelExport)
                tempExcelExport = "1";
            string tempReminderOnMeasure = "0";
            if (aProfile.FAllowReminderOnMeasure)
                tempReminderOnMeasure = "1";
            string tempReminderOnUnit = "0";
            if (aProfile.FAllowReminderOnUnit)
                tempReminderOnUnit = "1";
            string tempUpUnitsMeasures = "0";
            if (aProfile.FAllowUpUnitsMeasures)
                tempUpUnitsMeasures = "1";
            string tempColleaguesMeasures = "0";
            if (aProfile.FAllowColleaguesMeasures)
                tempColleaguesMeasures = "1";
            string tempTopUnitsThemes = "0";
            if (aProfile.FAllowTopUnitsThemes)
                tempTopUnitsThemes = "1";
            string tempColleaguesThemes = "0";
            if (aProfile.FAllowColleaguesThemes)
                tempColleaguesThemes = "1";
            string tempThemes = "0";
            if (aProfile.FAllowThemes)
                tempThemes = "1";
            string tempReminderOnTheme = "0";
            if (aProfile.FAllowReminderOnTheme)
                tempReminderOnTheme = "1";
            string tempReminderOnThemeUnit = "0";
            if (aProfile.FAllowReminderOnThemeUnit)
                tempReminderOnThemeUnit = "1";
            string tempOrders = "0";
            if (aProfile.FAllowOrders)
                tempOrders = "1";
            string tempReminderOnOrderUnit = "0";
            if (aProfile.FAllowReminderOnOrderUnit)
                tempReminderOnOrderUnit = "1";
            parameterList = new TParameterList();
            parameterList.addParameter("profileID", "string", aProfile.profileID);
            parameterList.addParameter("FReadLevel", "int", aProfile.FReadLevel.ToString());
            parameterList.addParameter("FEditLevel", "int", aProfile.FEditLevel.ToString());
            parameterList.addParameter("FSumLevel", "int", aProfile.FSumLevel.ToString());
            parameterList.addParameter("FAllowCommunications", "int", tempCommunications);
            parameterList.addParameter("FAllowMeasures", "int", tempMeasures);
            parameterList.addParameter("FAllowDelete", "int", tempDelete);
            parameterList.addParameter("FAllowExcelExport", "int", tempExcelExport);
            parameterList.addParameter("FAllowReminderOnMeasure", "int", tempReminderOnMeasure);
            parameterList.addParameter("FAllowReminderOnUnit", "int", tempReminderOnUnit);
            parameterList.addParameter("FAllowUpUnitsMeasures", "int", tempUpUnitsMeasures);
            parameterList.addParameter("FAllowColleaguesMeasures", "int", tempColleaguesMeasures);
            parameterList.addParameter("FAllowTopUnitsThemes", "int", tempTopUnitsThemes);
            parameterList.addParameter("FAllowColleaguesThemes", "int", tempColleaguesThemes);
            parameterList.addParameter("FAllowThemes", "int", tempThemes);
            parameterList.addParameter("FAllowReminderOnTheme", "int", tempReminderOnTheme);
            parameterList.addParameter("FAllowReminderOnThemeUnit", "int", tempReminderOnThemeUnit);
            parameterList.addParameter("FAllowOrders", "int", tempOrders);
            parameterList.addParameter("FAllowReminderOnOrderUnit", "int", tempReminderOnOrderUnit);
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO followup_profiles3 (profileID, FReadLevel, FEditLevel, FSumLevel, FAllowCommunications, FAllowMeasures, FAllowDelete, FAllowExcelExport, FAllowReminderOnMeasure, FAllowReminderOnUnit, FAllowUpUnitsMeasures, FAllowColleaguesMeasures, FAllowTopUnitsThemes, FAllowColleaguesThemes, FAllowThemes, FAllowReminderOnTheme, FAllowReminderOnThemeUnit, FAllowOrders, FAllowReminderOnOrderUnit) VALUES (@profileID, @FReadLevel, @FEditLevel, @FSumLevel, @FAllowCommunications, @FAllowMeasures, @FAllowDelete, @FAllowExcelExport, @FAllowReminderOnMeasure, @FAllowReminderOnUnit, @FAllowUpUnitsMeasures, @FAllowColleaguesMeasures, @FAllowTopUnitsThemes, @FAllowColleaguesThemes, @FAllowThemes, @FAllowReminderOnTheme, @FAllowReminderOnThemeUnit, @FAllowOrders, @FAllowReminderOnOrderUnit)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveProfileTeamspaceTasks(TProfileSetProjectplan aProfile, string aProjectID)
    {
        SqlDB dataReader;

        // ProfileID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("profileID", "string", aProfile.profileID);
        dataReader = new SqlDB("SELECT profileID FROM teamspace_profilestasks WHERE profileID=@profileID", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            string tempTasksRead = "0";
            if (aProfile.AllowTaskRead)
                tempTasksRead = "1";
            string tempTasksCreateEditDelete = "0";
            if (aProfile.AllowTaskCreateEditDelete)
                tempTasksCreateEditDelete = "1";
            string tempTasksFileUpload = "0";
            if (aProfile.AllowTaskFileUpload)
                tempTasksFileUpload = "1";
            string tempTasksImport = "0";
            if (aProfile.AllowTaskImport)
                tempTasksImport = "1";
            string tempTaskExport = "0";
            if (aProfile.AllowTaskExport)
                tempTaskExport = "1";
            parameterList = new TParameterList();
            parameterList.addParameter("profileID", "string", aProfile.profileID);
            parameterList.addParameter("AllowTaskRead", "int", tempTasksRead);
            parameterList.addParameter("AllowTaskCreateEditDelete", "int", tempTasksCreateEditDelete);
            parameterList.addParameter("AllowTaskFileUpload", "int", tempTasksFileUpload);
            parameterList.addParameter("AllowTaskExport", "int", tempTaskExport);
            parameterList.addParameter("AllowTaskImport", "int", tempTasksImport);
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO teamspace_profilestasks (profileID, AllowTaskRead, AllowTaskCreateEditDelete, AllowTaskFileUpload, AllowTaskExport, AllowTaskImport) VALUES (@profileID, @AllowTaskRead, @AllowTaskCreateEditDelete, @AllowTaskFileUpload, @AllowTaskExport, @AllowTaskImport)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveProfileTeamspaceContacts(TProfileSetContacts aProfile, string aProjectID)
    {
        SqlDB dataReader;

        // ProfileID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("profileID", "string", aProfile.profileID);
        dataReader = new SqlDB("SELECT profileID FROM teamspace_profilescontacts WHERE profileID=@profileID", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            string tempContactRead = "0";
            if (aProfile.AllowContactRead)
                tempContactRead = "1";
            string tempContactCreateEditDelete = "0";
            if (aProfile.AllowContactCreateEditDelete)
                tempContactCreateEditDelete = "1";
            string tempContactImport = "0";
            if (aProfile.AllowContactImport)
                tempContactImport = "1";
            string tempContactExport = "0";
            if (aProfile.AllowContactExport)
                tempContactExport = "1";
            parameterList = new TParameterList();
            parameterList.addParameter("profileID", "string", aProfile.profileID);
            parameterList.addParameter("AllowContactRead", "int", tempContactRead);
            parameterList.addParameter("AllowContactCreateEditDelete", "int", tempContactCreateEditDelete);
            parameterList.addParameter("AllowContactExport", "int", tempContactImport);
            parameterList.addParameter("AllowContactImport", "int", tempContactExport);
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO teamspace_profilescontacts (profileID, AllowContactRead, AllowContactCreateEditDelete, AllowContactExport, AllowContactImport) VALUES (@profileID, @AllowContactRead, @AllowContactCreateEditDelete, @AllowContactExport, @AllowContactImport)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveProfileTeamspaceContainer(TProfileSetContainer aProfile, string aProjectID)
    {
        SqlDB dataReader;

        // ProfileID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("profileID", "string", aProfile.profileID);
        dataReader = new SqlDB("SELECT profileID FROM teamspace_profilescontainer WHERE profileID=@profileID", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            string tempUpload = "0";
            if (aProfile.allowUpload)
                tempUpload = "1";
            string tempDownload = "0";
            if (aProfile.allowDownload)
                tempDownload = "1";
            string tempDeleteOwnFiles = "0";
            if (aProfile.allowDeleteOwnFiles)
                tempDeleteOwnFiles = "1";
            string tempDeleteAllFiles = "0";
            if (aProfile.allowDeleteAllFiles)
                tempDeleteAllFiles = "1";
            string tempAccessOwnFilesWithoutPassword = "0";
            if (aProfile.allowAccessOwnFilesWithoutPassword)
                tempAccessOwnFilesWithoutPassword = "1";
            string tempAccessAllFilesWithoutPassword = "0";
            if (aProfile.allowAccessAllFilesWithoutPassword)
                tempAccessAllFilesWithoutPassword = "1";
            string tempCreateFolder = "0";
            if (aProfile.allowCreateFolder)
                tempCreateFolder = "1";
            string tempDeleteOwnFolder = "0";
            if (aProfile.allowDeleteOwnFolder)
                tempDeleteOwnFolder = "1";
            string tempDeleteAllFolder = "0";
            if (aProfile.allowDeleteAllFolder)
                tempDeleteAllFolder = "1";
            string tempResetPassword = "0";
            if (aProfile.allowResetPassword)
                tempResetPassword = "1";
            string tempTakeOwnership = "0";
            if (aProfile.allowTakeOwnership)
                tempTakeOwnership = "1";
            parameterList = new TParameterList();
            parameterList.addParameter("profileID", "string", aProfile.profileID);
            parameterList.addParameter("allowUpload", "int", tempUpload);
            parameterList.addParameter("allowDownload", "int", tempDownload);
            parameterList.addParameter("allowDeleteOwnFiles", "int", tempDeleteOwnFiles);
            parameterList.addParameter("allowDeleteAllFiles", "int", tempDeleteAllFiles);
            parameterList.addParameter("allowAccessOwnFilesWithoutPassword", "int", tempAccessOwnFilesWithoutPassword);
            parameterList.addParameter("allowAccessAllFilesWithoutPassword", "int", tempAccessAllFilesWithoutPassword);
            parameterList.addParameter("allowCreateFolder", "int", tempCreateFolder);
            parameterList.addParameter("allowDeleteOwnFolder", "int", tempDeleteOwnFolder);
            parameterList.addParameter("allowDeleteAllFolder", "int", tempDeleteAllFolder);
            parameterList.addParameter("allowResetPassword", "int", tempResetPassword);
            parameterList.addParameter("allowTakeOwnership", "int", tempTakeOwnership);
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO teamspace_profilescontainer (profileID, allowUpload, allowDownload, allowDeleteOwnFiles, allowDeleteAllFiles, allowAccessOwnFilesWithoutPassword, allowAccessAllFilesWithoutPassword, allowCreateFolder, allowDeleteOwnFolder, allowDeleteAllFolder, allowResetPassword, allowTakeOwnership) VALUES (@profileID, @allowUpload, @allowDownload, @allowDeleteOwnFiles, @allowDeleteAllFiles, @allowAccessOwnFilesWithoutPassword, @allowAccessAllFilesWithoutPassword, @allowCreateFolder, @allowDeleteOwnFolder, @allowDeleteAllFolder, @allowResetPassword, @allowTakeOwnership)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
    public static bool saveProfileTeamspaceTranslate(TProfileSetTranslate aProfile, string aProjectID)
    {
        SqlDB dataReader;

        // ProfileID auf Eindeutigkeit prüfen
        bool exists = false;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("profileID", "string", aProfile.profileID);
        dataReader = new SqlDB("SELECT profileID FROM teamspace_profilestranslate WHERE profileID=@profileID", parameterList, aProjectID);
        if (dataReader.read())
        {
            exists = true;
        }
        dataReader.close();

        if (!exists)
        {
            string tempTranslateRead = "0";
            if (aProfile.AllowTranslateRead)
                tempTranslateRead = "1";
            string tempTranslateEdit = "0";
            if (aProfile.AllowTranslateEdit)
                tempTranslateEdit = "1";
            string tempTranslateRelease = "0";
            if (aProfile.AllowTranslateRelease)
                tempTranslateRelease = "1";
            string tempTranslateNew = "0";
            if (aProfile.AllowTranslateNew)
                tempTranslateNew = "1";
            string tempTranslateDelete = "0";
            if (aProfile.AllowTranslateDelete)
                tempTranslateDelete = "1";
            parameterList = new TParameterList();
            parameterList.addParameter("profileID", "string", aProfile.profileID);
            parameterList.addParameter("AllowTranslateRead", "int", tempTranslateRead);
            parameterList.addParameter("AllowTranslateEdit", "int", tempTranslateEdit);
            parameterList.addParameter("AllowTranslateRelease", "int", tempTranslateRelease);
            parameterList.addParameter("AllowTranslateNew", "int", tempTranslateNew);
            parameterList.addParameter("AllowTranslateDelete", "int", tempTranslateDelete);
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("INSERT INTO teamspace_profilestranslate (profileID, AllowTranslateRead, AllowTranslateEdit, AllowTranslateRelease, AllowTranslateNew, AllowTranslateDelete) VALUES (@profileID, @AllowTranslateRead, @AllowTranslateEdit, @AllowTranslateRelease, @AllowTranslateNew, @AllowTranslateDelete)", parameterList);

            return true;
        }
        else
        {
            return false;
        }
    }
}