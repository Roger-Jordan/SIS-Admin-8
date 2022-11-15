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
/// Zusammenfassungsbeschreibung für TMultiUser
/// </summary>
public class TRights
{
    public class TKeyValue
    {
        public string key;
        public bool value;
    }
    public string UserID;				// eindeutige Kennung des Benutzers der Berechtigung
    public int OrgID;				// eindeutige Kennung der Org-Einheit der Berechtigung
    public string OProfile;
    public int OReadLevel;					// Aktueller Benutzer ist aktiviert
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
    public string FProfile;
    public int FReadLevel;
    public int FEditLevel;
    public int FSumLevel;
    public bool FAllowCommunications;
    public bool FAllowDelete;
    public bool FAllowExcelExport;
    public bool FAllowReminderOnMeasure;
    public bool FAllowReminderOnUnit;
    public bool FAllowUpUnitsMeasures;
    public bool FAllowColleaguesMeasures;
    public bool FAllowMeasure;
    public bool FAllowThemes;
    public bool FAllowReminderOnThemes;
    public bool FAllowReminderOnThemeUnit;
    public bool FAllowTopUnitsThemes;
    public bool FAllowColleaguesThemes;
    public bool FAllowOrders;
    public bool FAllowReminderOnOrderUnit;
    public string RProfile;
    public int RReadLevel;
    public string DProfile;
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
    public int containerID;
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
    public bool TAllowTaskRead;
    public bool TAllowTaskEdit;
    public bool TAllowTaskFileUpload;
    public bool TAllowTaskExport;
    public bool TAllowTaskImport;
    public bool TAllowContactRead;
    public bool TAllowContactEdit;
    public bool TAllowContactExport;
    public bool TAllowContactImport;
    public bool AAllowAccountTransfer;
    public bool AAllowAccountManagement;
    public string groupID;
    public bool TAllowTranslateRead;
    public bool TAllowTranslateEdit;
    public bool TAllowTranslateRealease;
    public bool TAllowTranslateNew;
    public bool TAllowTranslateDelete;
    public int roleID;
    public int translateLanguageID;

    public TRights(int aID, string aRightsTable, TUsergroups aUserGroups, string aProjectID)
    {
        groupID = "";

        // Berechigungen aus DB ermitteln
        if (aRightsTable == "mapuserorgid")
        {
            // Org
            SqlDB dataReader;
            dataReader = new SqlDB("select userID, orgID, OProfile, OReadLevel, OEditLevel, OAllowDelegate, OAllowLock, OAllowEmployeeRead, OAllowEmployeeEdit, OAllowEmployeeImport, OAllowEmployeeExport, OAllowUnitAdd, OAllowUnitMove, OAllowUnitDelete, OAllowUnitProperty, OAllowReportRecipient, OAllowStructureImport, OAllowStructureExport, OAllowBouncesRead, OAllowBouncesEdit, OAllowBouncesDelete, OAllowBouncesExport, OAllowBouncesImport, OAllowReInvitation, OAllowPostNomination, OAllowPostNominationImport, FProfile, FReadLevel, FEditLevel, FSumLevel, FAllowCommunications, FAllowMeasures, FAllowDelete, FAllowExcelExport, FAllowReminderOnMeasure, FAllowReminderOnUnit, FAllowUpUnitsMeasures, FAllowColleaguesMeasures, FAllowTopUnitsThemes, FAllowColleaguesThemes, FAllowThemes, FAllowReminderOnTheme, FAllowReminderOnThemeUnit, FAllowOrders, FAllowReminderOnOrderUnit, RProfile, RReadLevel, DProfile, DReadLevel, DAllowMail, DAllowDownload, DAllowLog, DAllowType1, DAllowType2, DAllowType3, DAllowType4, DAllowType5, AEditLevelWorkshop, AReadLevelStructure FROM mapuserorgid WHERE ID='" + aID.ToString() + "'", aProjectID);
            if (dataReader.read())
            {
                UserID = dataReader.getString(0);
                OrgID = dataReader.getInt32(1);
                OProfile = dataReader.getString(2);
                OReadLevel = dataReader.getInt32(3);
                OEditLevel = dataReader.getInt32(4);
                OAllowDelegate = dataReader.getBool(5);
                OAllowLock = dataReader.getBool(6);
                OAllowEmployeeRead = dataReader.getBool(7);
                OAllowEmployeeEdit = dataReader.getBool(8);
                OAllowEmployeeImport = dataReader.getBool(9);
                OAllowEmployeeExport = dataReader.getBool(10);
                OAllowUnitAdd = dataReader.getBool(11);
                OAllowUnitMove = dataReader.getBool(12);
                OAllowUnitDelete = dataReader.getBool(13);
                OAllowUnitProperty = dataReader.getBool(14);
                OAllowReportRecipient = dataReader.getBool(15);
                OAllowStructureImport = dataReader.getBool(16);
                OAllowStructureExport = dataReader.getBool(17);
                OAllowBouncesRead = dataReader.getBool(18);
                OAllowBouncesEdit = dataReader.getBool(19);
                OAllowBouncesDelete = dataReader.getBool(20);
                OAllowBouncesExport = dataReader.getBool(21);
                OAllowBouncesImport = dataReader.getBool(22);
                OAllowReInvitation = dataReader.getBool(23);
                OAllowPostNomination = dataReader.getBool(24);
                OAllowPostNominationImport = dataReader.getBool(25);
                FProfile = dataReader.getString(26);
                FReadLevel = dataReader.getInt32(27);
                FEditLevel = dataReader.getInt32(28);
                FSumLevel = dataReader.getInt32(29);
                FAllowCommunications = dataReader.getBool(30);
                FAllowMeasure = dataReader.getBool(31);
                FAllowDelete = dataReader.getBool(32);
                FAllowExcelExport = dataReader.getBool(33);
                FAllowReminderOnMeasure = dataReader.getBool(34);
                FAllowReminderOnUnit = dataReader.getBool(35);
                FAllowUpUnitsMeasures = dataReader.getBool(36);
                FAllowColleaguesMeasures = dataReader.getBool(37);
                FAllowTopUnitsThemes = dataReader.getBool(38);
                FAllowColleaguesThemes = dataReader.getBool(39);
                FAllowThemes = dataReader.getBool(40);
                FAllowReminderOnThemes = dataReader.getBool(41);
                FAllowReminderOnThemeUnit = dataReader.getBool(42);
                FAllowOrders = dataReader.getBool(43);
                FAllowReminderOnOrderUnit = dataReader.getBool(44);
                RProfile = dataReader.getString(45);
                RReadLevel = dataReader.getInt32(46);
                DProfile = dataReader.getString(47);
                DReadLevel = dataReader.getInt32(48);
                DAllowMail = dataReader.getBool(49);
                DAllowDownload = dataReader.getBool(50);
                DAllowLog = dataReader.getBool(51);
                DAllowType1 = dataReader.getBool(52);
                DAllowType2 = dataReader.getBool(53);
                DAllowType3 = dataReader.getBool(54);
                DAllowType4 = dataReader.getBool(55);
                DAllowType5 = dataReader.getBool(56);
                AEditLevelWorkshop = dataReader.getInt32(57);
                AReadLevelStructure = dataReader.getInt32(58);
            }
            dataReader.close();
        }
        // Org 1
        if (aRightsTable == "mapuserorgid1")
        {
            SqlDB dataReader;
            dataReader = new SqlDB("select userID, orgID, OProfile, OReadLevel, OEditLevel, OAllowDelegate, OAllowLock, OAllowEmployeeRead, OAllowEmployeeEdit, OAllowEmployeeImport, OAllowEmployeeExport, OAllowUnitAdd, OAllowUnitMove, OAllowUnitDelete, OAllowUnitProperty, OAllowReportRecipient, OAllowStructureImport, OAllowStructureExport, FProfile, FReadLevel, FEditLevel, FSumLevel, FAllowCommunications, FAllowMeasures, FAllowDelete, FAllowExcelExport, FAllowReminderOnMeasure, FAllowReminderOnUnit, FAllowUpUnitsMeasures, FAllowColleaguesMeasures, FAllowTopUnitsThemes, FAllowColleaguesThemes, FAllowThemes, FAllowReminderOnTheme, FAllowReminderOnThemeUnit, FAllowOrders, FAllowReminderOnOrderUnit, RProfile, RReadLevel, DProfile, DReadLevel, DAllowMail, DAllowDownload, DAllowLog, DAllowType1, DAllowType2, DAllowType3, DAllowType4, DAllowType5, AEditLevelWorkshop, AReadLevelStructure FROM mapuserorgid1 WHERE ID='" + aID.ToString() + "'", aProjectID);
            if (dataReader.read())
            {
                UserID = dataReader.getString(0);
                OrgID = dataReader.getInt32(1);
                OProfile = dataReader.getString(2);
                OReadLevel = dataReader.getInt32(3);
                OEditLevel = dataReader.getInt32(4);
                OAllowDelegate = dataReader.getBool(5);
                OAllowLock = dataReader.getBool(6);
                OAllowEmployeeRead = dataReader.getBool(7);
                OAllowEmployeeEdit = dataReader.getBool(8);
                OAllowEmployeeImport = dataReader.getBool(9);
                OAllowEmployeeExport = dataReader.getBool(10);
                OAllowUnitAdd = dataReader.getBool(11);
                OAllowUnitMove = dataReader.getBool(12);
                OAllowUnitDelete = dataReader.getBool(13);
                OAllowUnitProperty = dataReader.getBool(14);
                OAllowReportRecipient = dataReader.getBool(15);
                OAllowStructureImport = dataReader.getBool(16);
                OAllowStructureExport = dataReader.getBool(17);
                FProfile = dataReader.getString(18);
                FReadLevel = dataReader.getInt32(19);
                FEditLevel = dataReader.getInt32(20);
                FSumLevel = dataReader.getInt32(21);
                FAllowCommunications = dataReader.getBool(22);
                FAllowMeasure = dataReader.getBool(23);
                FAllowDelete = dataReader.getBool(24);
                FAllowExcelExport = dataReader.getBool(25);
                FAllowReminderOnMeasure = dataReader.getBool(26);
                FAllowReminderOnUnit = dataReader.getBool(27);
                FAllowUpUnitsMeasures = dataReader.getBool(28);
                FAllowColleaguesMeasures = dataReader.getBool(29);
                FAllowTopUnitsThemes = dataReader.getBool(30);
                FAllowColleaguesThemes = dataReader.getBool(31);
                FAllowThemes = dataReader.getBool(32);
                FAllowReminderOnThemes = dataReader.getBool(33);
                FAllowReminderOnThemeUnit = dataReader.getBool(34);
                FAllowOrders = dataReader.getBool(35);
                FAllowReminderOnOrderUnit = dataReader.getBool(36);
                RProfile = dataReader.getString(37);
                RReadLevel = dataReader.getInt32(38);
                DProfile = dataReader.getString(39);
                DReadLevel = dataReader.getInt32(40);
                DAllowMail = dataReader.getBool(41);
                DAllowDownload = dataReader.getBool(42);
                DAllowLog = dataReader.getBool(43);
                DAllowType1 = dataReader.getBool(44);
                DAllowType2 = dataReader.getBool(45);
                DAllowType3 = dataReader.getBool(46);
                DAllowType4 = dataReader.getBool(47);
                DAllowType5 = dataReader.getBool(48);
                AEditLevelWorkshop = dataReader.getInt32(49);
                AReadLevelStructure = dataReader.getInt32(50);
            }
            dataReader.close();
        }
        // Org 2
        if (aRightsTable == "mapuserorgid2")
        {
            SqlDB dataReader;
            dataReader = new SqlDB("select userID, orgID, OProfile, OReadLevel, OEditLevel, OAllowDelegate, OAllowLock, OAllowEmployeeRead, OAllowEmployeeEdit, OAllowEmployeeImport, OAllowEmployeeExport, OAllowUnitAdd, OAllowUnitMove, OAllowUnitDelete, OAllowUnitProperty, OAllowReportRecipient, OAllowStructureImport, OAllowStructureExport, FProfile, FReadLevel, FEditLevel, FSumLevel, FAllowCommunications, FAllowMeasures, FAllowDelete, FAllowExcelExport, FAllowReminderOnMeasure, FAllowReminderOnUnit, FAllowUpUnitsMeasures, FAllowColleaguesMeasures, FAllowTopUnitsThemes, FAllowColleaguesThemes, FAllowThemes, FAllowReminderOnTheme, FAllowReminderOnThemeUnit, FAllowOrders, FAllowReminderOnOrderUnit, RProfile, RReadLevel, DProfile, DReadLevel, DAllowMail, DAllowDownload, DAllowLog, DAllowType1, DAllowType2, DAllowType3, DAllowType4, DAllowType5, AEditLevelWorkshop, AReadLevelStructure FROM mapuserorgid2 WHERE ID='" + aID.ToString() + "'", aProjectID);
            if (dataReader.read())
            {
                UserID = dataReader.getString(0);
                OrgID = dataReader.getInt32(1);
                OProfile = dataReader.getString(2);
                OReadLevel = dataReader.getInt32(3);
                OEditLevel = dataReader.getInt32(4);
                OAllowDelegate = dataReader.getBool(5);
                OAllowLock = dataReader.getBool(6);
                OAllowEmployeeRead = dataReader.getBool(7);
                OAllowEmployeeEdit = dataReader.getBool(8);
                OAllowEmployeeImport = dataReader.getBool(9);
                OAllowEmployeeExport = dataReader.getBool(10);
                OAllowUnitAdd = dataReader.getBool(11);
                OAllowUnitMove = dataReader.getBool(12);
                OAllowUnitDelete = dataReader.getBool(13);
                OAllowUnitProperty = dataReader.getBool(14);
                OAllowReportRecipient = dataReader.getBool(15);
                OAllowStructureImport = dataReader.getBool(16);
                OAllowStructureExport = dataReader.getBool(17);
                FProfile = dataReader.getString(18);
                FReadLevel = dataReader.getInt32(19);
                FEditLevel = dataReader.getInt32(20);
                FSumLevel = dataReader.getInt32(21);
                FAllowCommunications = dataReader.getBool(22);
                FAllowMeasure = dataReader.getBool(23);
                FAllowDelete = dataReader.getBool(24);
                FAllowExcelExport = dataReader.getBool(25);
                FAllowReminderOnMeasure = dataReader.getBool(26);
                FAllowReminderOnUnit = dataReader.getBool(27);
                FAllowUpUnitsMeasures = dataReader.getBool(28);
                FAllowColleaguesMeasures = dataReader.getBool(29);
                FAllowTopUnitsThemes = dataReader.getBool(30);
                FAllowColleaguesThemes = dataReader.getBool(31);
                FAllowThemes = dataReader.getBool(32);
                FAllowReminderOnThemes = dataReader.getBool(33);
                FAllowReminderOnThemeUnit = dataReader.getBool(34);
                FAllowOrders = dataReader.getBool(35);
                FAllowReminderOnOrderUnit = dataReader.getBool(36);
                RProfile = dataReader.getString(37);
                RReadLevel = dataReader.getInt32(38);
                DProfile = dataReader.getString(39);
                DReadLevel = dataReader.getInt32(40);
                DAllowMail = dataReader.getBool(41);
                DAllowDownload = dataReader.getBool(42);
                DAllowLog = dataReader.getBool(43);
                DAllowType1 = dataReader.getBool(44);
                DAllowType2 = dataReader.getBool(45);
                DAllowType3 = dataReader.getBool(46);
                DAllowType4 = dataReader.getBool(47);
                DAllowType5 = dataReader.getBool(48);
                AEditLevelWorkshop = dataReader.getInt32(49);
                AReadLevelStructure = dataReader.getInt32(50);
            }
            dataReader.close();
        }
        // Org 3
        if (aRightsTable == "mapuserorgid3")
        {
            SqlDB dataReader;
            dataReader = new SqlDB("select userID, orgID, OProfile, OReadLevel, OEditLevel, OAllowDelegate, OAllowLock, OAllowEmployeeRead, OAllowEmployeeEdit, OAllowEmployeeImport, OAllowEmployeeExport, OAllowUnitAdd, OAllowUnitMove, OAllowUnitDelete, OAllowUnitProperty, OAllowReportRecipient, OAllowStructureImport, OAllowStructureExport, FProfile, FReadLevel, FEditLevel, FSumLevel, FAllowCommunications, FAllowMeasures, FAllowDelete, FAllowExcelExport, FAllowReminderOnMeasure, FAllowReminderOnUnit, FAllowUpUnitsMeasures, FAllowColleaguesMeasures, FAllowTopUnitsThemes, FAllowColleaguesThemes, FAllowThemes, FAllowReminderOnTheme, FAllowReminderOnThemeUnit, FAllowOrders, FAllowReminderOnOrderUnit, RProfile, RReadLevel, DProfile, DReadLevel, DAllowMail, DAllowDownload, DAllowLog, DAllowType1, DAllowType2, DAllowType3, DAllowType4, DAllowType5, AEditLevelWorkshop, AReadLevelStructure FROM mapuserorgid3 WHERE ID='" + aID.ToString() + "'", aProjectID);
            if (dataReader.read())
            {
                UserID = dataReader.getString(0);
                OrgID = dataReader.getInt32(1);
                OProfile = dataReader.getString(2);
                OReadLevel = dataReader.getInt32(3);
                OEditLevel = dataReader.getInt32(4);
                OAllowDelegate = dataReader.getBool(5);
                OAllowLock = dataReader.getBool(6);
                OAllowEmployeeRead = dataReader.getBool(7);
                OAllowEmployeeEdit = dataReader.getBool(8);
                OAllowEmployeeImport = dataReader.getBool(9);
                OAllowEmployeeExport = dataReader.getBool(10);
                OAllowUnitAdd = dataReader.getBool(11);
                OAllowUnitMove = dataReader.getBool(12);
                OAllowUnitDelete = dataReader.getBool(13);
                OAllowUnitProperty = dataReader.getBool(14);
                OAllowReportRecipient = dataReader.getBool(15);
                OAllowStructureImport = dataReader.getBool(16);
                OAllowStructureExport = dataReader.getBool(17);
                FProfile = dataReader.getString(18);
                FReadLevel = dataReader.getInt32(19);
                FEditLevel = dataReader.getInt32(20);
                FSumLevel = dataReader.getInt32(21);
                FAllowCommunications = dataReader.getBool(22);
                FAllowMeasure = dataReader.getBool(23);
                FAllowDelete = dataReader.getBool(24);
                FAllowExcelExport = dataReader.getBool(25);
                FAllowReminderOnMeasure = dataReader.getBool(26);
                FAllowReminderOnUnit = dataReader.getBool(27);
                FAllowUpUnitsMeasures = dataReader.getBool(28);
                FAllowColleaguesMeasures = dataReader.getBool(29);
                FAllowTopUnitsThemes = dataReader.getBool(30);
                FAllowColleaguesThemes = dataReader.getBool(31);
                FAllowThemes = dataReader.getBool(32);
                FAllowReminderOnThemes = dataReader.getBool(33);
                FAllowReminderOnThemeUnit = dataReader.getBool(34);
                FAllowOrders = dataReader.getBool(35);
                FAllowReminderOnOrderUnit = dataReader.getBool(36);
                RProfile = dataReader.getString(37);
                RReadLevel = dataReader.getInt32(38);
                DProfile = dataReader.getString(39);
                DReadLevel = dataReader.getInt32(40);
                DAllowMail = dataReader.getBool(41);
                DAllowDownload = dataReader.getBool(42);
                DAllowLog = dataReader.getBool(43);
                DAllowType1 = dataReader.getBool(44);
                DAllowType2 = dataReader.getBool(45);
                DAllowType3 = dataReader.getBool(46);
                DAllowType4 = dataReader.getBool(47);
                DAllowType5 = dataReader.getBool(48);
                AEditLevelWorkshop = dataReader.getInt32(49);
                AReadLevelStructure = dataReader.getInt32(50);
            }
            dataReader.close();
        }
        // OrgBack
        if (aRightsTable == "mapuserback")
        {
            SqlDB dataReader;
            dataReader = new SqlDB("select userID, orgID, OReadLevel FROM mapuserback WHERE ID='" + aID.ToString() + "'", aProjectID);
            if (dataReader.read())
            {
                UserID = dataReader.getString(0);
                OrgID = dataReader.getInt32(1);
                OReadLevel = dataReader.getInt32(2);
            }
            dataReader.close();
        }
        // OrgBack1
        if (aRightsTable == "mapuserback1")
        {
            SqlDB dataReader;
            dataReader = new SqlDB("select userID, orgID, OReadLevel FROM mapuserback1 WHERE ID='" + aID.ToString() + "'", aProjectID);
            if (dataReader.read())
            {
                UserID = dataReader.getString(0);
                OrgID = dataReader.getInt32(1);
                OReadLevel = dataReader.getInt32(2);
            }
            dataReader.close();
        }
        // OrgBack2
        if (aRightsTable == "mapuserback2")
        {
            SqlDB dataReader;
            dataReader = new SqlDB("select userID, orgID, OReadLevel FROM mapuserback2 WHERE ID='" + aID.ToString() + "'", aProjectID);
            if (dataReader.read())
            {
                UserID = dataReader.getString(0);
                OrgID = dataReader.getInt32(1);
                OReadLevel = dataReader.getInt32(2);
            }
            dataReader.close();
        }
        // OrgBack3
        if (aRightsTable == "mapuserback3")
        {
            SqlDB dataReader;
            dataReader = new SqlDB("select userID, orgID, OReadLevel FROM mapuserback3 WHERE ID='" + aID.ToString() + "'", aProjectID);
            if (dataReader.read())
            {
                UserID = dataReader.getString(0);
                OrgID = dataReader.getInt32(1);
                OReadLevel = dataReader.getInt32(2);
            }
            dataReader.close();
        }
        // Container
        if (aRightsTable == "mapusercontainer")
        {
            SqlDB dataReader;
            dataReader = new SqlDB("select userID, containerID, allowUpload, allowDownload, allowDeleteOwnFiles, allowDeleteAllFiles, allowAccessOwnFilesWithoutPassword, allowAccessAllFilesWithoutPassword, allowCreateFolder, allowDeleteOwnFolder, allowDeleteAllFolder, allowResetPassword, allowTakeOwnership FROM mapusercontainer WHERE ID='" + aID.ToString() + "'", aProjectID);
            if (dataReader.read())
            {
                UserID = dataReader.getString(0);
                containerID = dataReader.getInt32(1);
                TAllowUpload = dataReader.getBool(2);
                TAllowDownload = dataReader.getBool(3);
                TAllowDeleteOwnFiles = dataReader.getBool(4);
                TAllowDeleteAllFiles = dataReader.getBool(5);
                TAllowAccessOwnFilesWithoutPassword = dataReader.getBool(6);
                TAllowAccessAllFilesWithoutPassword = dataReader.getBool(7);
                TAllowCreateFolder = dataReader.getBool(8);
                TAllowDeleteOwnFolder = dataReader.getBool(9);
                TAllowDeleteAllFolder = dataReader.getBool(10);
                TAllowResetPassword = dataReader.getBool(11);
                TAllowTakeOwnership = dataReader.getBool(12);
            }
            dataReader.close();
        }
        // TaskContacts
        if (aRightsTable == "mapuserrights")
        {
            SqlDB dataReader;
            dataReader = new SqlDB("select userID, AllowTaskRead, AllowTaskCreateEditDelete, AllowTaskFileUpload, AllowTaskExport, AllowTaskImport, AllowContactRead, AllowContactCreateEditDelete, AllowContactExport, AllowContactImport, AllowTranslateRead, AllowTranslateEdit, AllowTranslateRelease, AllowTranslateNew, AllowTranslateDelete, AllowAccountTransfer, AllowAccountManagement FROM mapuserrights WHERE ID='" + aID.ToString() + "'", aProjectID);
            if (dataReader.read())
            {
                UserID = dataReader.getString(0);
                TAllowTaskRead = dataReader.getBool(1);
                TAllowTaskEdit = dataReader.getBool(2);
                TAllowTaskFileUpload = dataReader.getBool(3);
                TAllowTaskExport = dataReader.getBool(4);
                TAllowTaskImport = dataReader.getBool(5);
                TAllowContactRead = dataReader.getBool(6);
                TAllowContactEdit = dataReader.getBool(7);
                TAllowContactExport = dataReader.getBool(8);
                TAllowContactImport = dataReader.getBool(9);
                TAllowTranslateRead = dataReader.getBool(10);
                TAllowTranslateEdit = dataReader.getBool(11);
                TAllowTranslateRealease = dataReader.getBool(12);
                TAllowTranslateNew = dataReader.getBool(13);
                TAllowTranslateDelete = dataReader.getBool(14);
                AAllowAccountTransfer = dataReader.getBool(15);
                AAllowAccountManagement = dataReader.getBool(16); ;
            }
            dataReader.close();
        }
        // Taskroles
        if (aRightsTable == "mapusertaskrole")
        {
            SqlDB dataReader;
            dataReader = new SqlDB("select userID, roleID FROM mapusertaskrole WHERE ID='" + aID.ToString() + "'", aProjectID);
            if (dataReader.read())
            {
                UserID = dataReader.getString(0);
                roleID = dataReader.getInt32(1);
            }
            dataReader.close();
        }
        // Usergroups
        if (aRightsTable == "mapusergroup")
        {
            SqlDB dataReader;
            dataReader = new SqlDB("select userID, groupID FROM mapusergroup WHERE ID='" + aID.ToString() + "'", aProjectID);
            if (dataReader.read())
            {
                UserID = dataReader.getString(0);
                groupID = dataReader.getString(1);
            }
            dataReader.close();
        }
        // TranslateLanguage
        if (aRightsTable == "mapusertranslatelanguage")
        {
            SqlDB dataReader;
            dataReader = new SqlDB("select userID, translateLanguageID FROM mapusertranslateLanguage WHERE ID='" + aID.ToString() + "'", aProjectID);
            if (dataReader.read())
            {
                UserID = dataReader.getString(0);
                translateLanguageID = dataReader.getInt32(1);
            }
            dataReader.close();
        }
    }
}

