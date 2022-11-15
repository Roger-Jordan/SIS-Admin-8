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
public class TMultiRights
{
    public class TKeyValue
    {
        public string key;
        public int value;
    }
    public ArrayList IDList;            // Liste der eindeutigen IDs alle Berechtigungen
    public string UserID;				// eindeutige Kennung des Benutzers der Berechtigung
    public string OrgID;				// eindeutige Kennung der Org-Einheit der Berechtigung
    public string OProfile;
    public int OReadLevel;					
    public int OEditLevel;
    public int OAllowDelegate;
    public int OAllowLock;
    public int OAllowEmployeeRead;
    public int OAllowEmployeeEdit;
    public int OAllowEmployeeImport;
    public int OAllowEmployeeExport;
    public int OAllowUnitAdd;
    public int OAllowUnitMove;
    public int OAllowUnitDelete;
    public int OAllowUnitProperty;
    public int OAllowReportRecipient;
    public int OAllowStructureImport;
    public int OAllowStructureExport;
    public int OAllowBouncesRead;
    public int OAllowBouncesEdit;
    public int OAllowBouncesDelete;
    public int OAllowBouncesExport;
    public int OAllowBouncesImport;
    public int OAllowReInvitation;
    public int OAllowPostNomination;
    public int OAllowPostNominationImport;
    public string FProfile;
    public int FReadLevel;
    public int FEditLevel;
    public int FSumLevel;
    public int FAllowCommunications;
    public int FAllowMeasures;
    public int FAllowDelete;
    public int FAllowExcelExport;
    public int FAllowReminderOnMeasure;
    public int FAllowReminderOnUnit;
    public int FAllowUpUnitsMeasures;
    public int FAllowColleaguesMeasures;
    public int FAllowThemes;
    public int FAllowReminderOnThemes;
    public int FAllowReminderOnThemeUnit;
    public int FAllowTopUnitsThemes;
    public int FAllowColleaguesThemes;
    public int FAllowOrders;
    public int FAllowReminderOnOrderUnit;
    public string RProfile;
    public int RReadLevel;
    public string DProfile;
    public int DReadLevel;
    public int DAllowMail;
    public int DAllowDownload;
    public int DAllowLog;
    public int DAllowType1;
    public int DAllowType2;
    public int DAllowType3;
    public int DAllowType4;
    public int DAllowType5;
    public int AEditLevelWorkshop;
    public int AReadLevelStructure;
    public int containerID;
    public string TContainerProfile;
    public int TAllowUpload;
    public int TAllowDownload;
    public int TAllowDeleteOwnFiles;
    public int TAllowDeleteAllFiles;
    public int TAllowAccessOwnFilesWithoutPassword;
    public int TAllowAccessAllFilesWithoutPassword;
    public int TAllowCreateFolder;
    public int TAllowDeleteOwnFolder;
    public int TAllowDeleteAllFolder;
    public int TAllowResetPassword;
    public int TAllowTakeOwnership;
    public string TTaskProfile;
    public int TAllowTaskRead;
    public int TAllowTaskEdit;
    public int TAllowTaskFileUpload;
    public int TAllowTaskExport;
    public int TAllowTaskImport;
    public string TContactProfile;
    public int TAllowContactRead;
    public int TAllowContactEdit;
    public int TAllowContactExport;
    public int TAllowContactImport;
    public int AAllowAccountTransfer;
    public int AAllowAccountManagement;
    public string groupID;
    public string TTranslateProfile;
    public int TAllowTranslateRead;
    public int TAllowTranslateEdit;
    public int TAllowTranslateRelease;
    public int TAllowTranslateNew;
    public int TAllowTranslateDelete;
    public int roleID;
    public int translateLanaguageID;

    public TMultiRights(ArrayList aIDList)
    {
        IDList = aIDList;
    }
    protected string checkUpdate(string aValue, string aField, string aUpdate, ref TParameterList parameterList, bool crypt)
    {
        if (crypt)
        {
            parameterList.addParameter(aField, "string", TCrypt.StringVerschluesseln(aValue));
        }
        else
        {
            parameterList.addParameter(aField, "string", aValue);
        }

        if (aValue == "*")
            return aUpdate;
        else
            if (aUpdate == "")
                return "[" + aField + "]=@" + aField;
            else
                return aUpdate + ", [" + aField + "]=@" + aField;
    }
    protected string checkUpdate(int aValue, string aField, string aUpdate, string aInvalid, ref TParameterList parameterList)
    {
        parameterList.addParameter(aField, "int", aValue.ToString());

        if (aValue.ToString() == aInvalid)
            return aUpdate;
        else
            if (aUpdate == "")
                return aField + "=@" + aField;
            else
                return aUpdate + ", " + aField + "=@" + aField;
    }
    public void updateOrgRight(string aProjectID, TSessionObj aSessionObj)
    {
        SqlDB dataReader;

        if (aSessionObj.filterOrgTableIndex == 0)
        {
            // Update-String konstruieren
            string tempUpdate = "";
            TParameterList parameterList = new TParameterList();
            tempUpdate = checkUpdate(UserID, "UserID", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(OrgID, "OrgID", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(OProfile, "OProfile", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(OReadLevel, "OReadLevel", tempUpdate, "-2", ref parameterList);
            tempUpdate = checkUpdate(OEditLevel, "OEditLevel", tempUpdate, "-2", ref parameterList);
            tempUpdate = checkUpdate(OAllowDelegate, "OAllowDelegate", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowLock, "OAllowLock", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowEmployeeRead, "OAllowEmployeeRead", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowEmployeeEdit, "OAllowEmployeeEdit", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowEmployeeImport, "OAllowEmployeeImport", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowEmployeeExport, "OAllowEmployeeExport", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowUnitAdd, "OAllowUnitAdd", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowUnitMove, "OAllowUnitMove", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowUnitDelete, "OAllowUnitDelete", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowUnitProperty, "OAllowUnitProperty", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowReportRecipient, "OAllowReportRecipient", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowStructureImport, "OAllowStructureImport", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowStructureExport, "OAllowStructureExport", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowBouncesRead, "OAllowBouncesRead", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowBouncesEdit, "OAllowBouncesEdit", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowBouncesDelete, "OAllowBouncesDelete", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowBouncesExport, "OAllowBouncesExport", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowBouncesImport, "OAllowBouncesImport", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowReInvitation, "OAllowReInvitation", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowPostNomination, "OAllowPostNomination", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowPostNominationImport, "OAllowPostNominationImport", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FProfile, "FProfile", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(FReadLevel, "FReadLevel", tempUpdate, "-2", ref parameterList);
            tempUpdate = checkUpdate(FEditLevel, "FEditLevel", tempUpdate, "-2", ref parameterList);
            tempUpdate = checkUpdate(FSumLevel, "FSumLevel", tempUpdate, "-2", ref parameterList);
            tempUpdate = checkUpdate(FAllowDelete, "FAllowDelete", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowExcelExport, "FAllowExcelExport", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowReminderOnMeasure, "FAllowReminderOnMeasure", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowReminderOnUnit, "FAllowReminderOnUnit", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowUpUnitsMeasures, "FAllowUpUnitsMeasures", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowColleaguesMeasures, "FAllowColleaguesMeasures", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowThemes, "FAllowThemes", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowReminderOnThemes, "FAllowReminderOnTheme", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowReminderOnThemeUnit, "FAllowReminderOnThemeUnit", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowTopUnitsThemes, "FAllowTopUnitsThemes", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowColleaguesThemes, "FAllowColleaguesThemes", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowOrders, "FAllowOrders", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowReminderOnOrderUnit, "FAllowReminderOnOrderUnit", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(RProfile, "RProfile", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(RReadLevel, "RReadLevel", tempUpdate, "-2", ref parameterList);
            tempUpdate = checkUpdate(DProfile, "DProfile", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(DReadLevel, "DReadLevel", tempUpdate, "-2", ref parameterList);
            tempUpdate = checkUpdate(DAllowMail, "DAllowMail", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(DAllowDownload, "DAllowDownload", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(DAllowLog, "DAllowLog", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(DAllowType1, "DAllowType1", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(DAllowType2, "DAllowType2", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(DAllowType3, "DAllowType3", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(DAllowType4, "DAllowType4", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(DAllowType5, "DAllowType5", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(AEditLevelWorkshop, "AEditLevelWorkshop", tempUpdate, "-2", ref parameterList);
            tempUpdate = checkUpdate(AReadLevelStructure, "AReadLevelStructure", tempUpdate, "-2", ref parameterList);

            parameterList.addParameter("ID", "string", "");

            // Schleife über alle Berechtigungen
            int index = 0;
            foreach (string tempRightID in aSessionObj.selectedIDList)
            {
                index++;
                TProcessStatus.updateProcessInfo(aSessionObj.ProcessID, 2, index, 0, aSessionObj.selectedIDList.Count, aSessionObj.Project.ProjectID);

                parameterList.changeParameterValue("ID", tempRightID);
                // Daten schreiben
                if (tempUpdate != "")
                {
                    dataReader = new SqlDB(aProjectID);
                    dataReader.execSQLwithParameter("UPDATE mapuserorgid SET " + tempUpdate + " WHERE ID=@ID", parameterList);
                }
            }
        }
        if (aSessionObj.filterOrgTableIndex == 1)
        {
            // Update-String konstruieren
            string tempUpdate = "";
            TParameterList parameterList = new TParameterList();
            tempUpdate = checkUpdate(UserID, "UserID", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(OrgID, "OrgID", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(OProfile, "OProfile", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(OReadLevel, "OReadLevel", tempUpdate, "-2", ref parameterList);
            tempUpdate = checkUpdate(OEditLevel, "OEditLevel", tempUpdate, "-2", ref parameterList);
            tempUpdate = checkUpdate(OAllowDelegate, "OAllowDelegate", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowLock, "OAllowLock", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowEmployeeRead, "OAllowEmployeeRead", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowEmployeeEdit, "OAllowEmployeeEdit", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowEmployeeImport, "OAllowEmployeeImport", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowEmployeeExport, "OAllowEmployeeExport", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowUnitAdd, "OAllowUnitAdd", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowUnitMove, "OAllowUnitMove", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowUnitDelete, "OAllowUnitDelete", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowUnitProperty, "OAllowUnitProperty", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowReportRecipient, "OAllowReportRecipient", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowStructureImport, "OAllowStructureImport", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowStructureExport, "OAllowStructureExport", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FProfile, "FProfile", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(FReadLevel, "FReadLevel", tempUpdate, "-2", ref parameterList);
            tempUpdate = checkUpdate(FEditLevel, "FEditLevel", tempUpdate, "-2", ref parameterList);
            tempUpdate = checkUpdate(FSumLevel, "FSumLevel", tempUpdate, "-2", ref parameterList);
            tempUpdate = checkUpdate(FAllowDelete, "FAllowDelete", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowExcelExport, "FAllowExcelExport", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowReminderOnMeasure, "FAllowReminderOnMeasure", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowReminderOnUnit, "FAllowReminderOnUnit", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowUpUnitsMeasures, "FAllowUpUnitsMeasures", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowColleaguesMeasures, "FAllowColleaguesMeasures", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowThemes, "FAllowThemes", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowReminderOnThemes, "FAllowReminderOnTheme", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowReminderOnThemeUnit, "FAllowReminderOnThemeUnit", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowTopUnitsThemes, "FAllowTopUnitsThemes", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowColleaguesThemes, "FAllowColleaguesThemes", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowOrders, "FAllowOrders", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowReminderOnOrderUnit, "FAllowReminderOnOrderUnit", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(RProfile, "RProfile", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(RReadLevel, "RReadLevel", tempUpdate, "-2", ref parameterList);
            tempUpdate = checkUpdate(DProfile, "DProfile", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(DReadLevel, "DReadLevel", tempUpdate, "-2", ref parameterList);
            tempUpdate = checkUpdate(DAllowMail, "DAllowMail", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(DAllowDownload, "DAllowDownload", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(DAllowLog, "DAllowLog", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(DAllowType1, "DAllowType1", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(DAllowType2, "DAllowType2", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(DAllowType3, "DAllowType3", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(DAllowType4, "DAllowType4", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(DAllowType5, "DAllowType5", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(AEditLevelWorkshop, "AEditLevelWorkshop", tempUpdate, "-2", ref parameterList);
            tempUpdate = checkUpdate(AReadLevelStructure, "AReadLevelStructure", tempUpdate, "-2", ref parameterList);

            parameterList.addParameter("ID", "string", "");

            // Schleife über alle Berechtigungen
            int index = 0;
            foreach (string tempRightID in aSessionObj.selectedIDList)
            {
                index++;
                TProcessStatus.updateProcessInfo(aSessionObj.ProcessID, 2, index, 0, aSessionObj.selectedIDList.Count, aSessionObj.Project.ProjectID);

                parameterList.changeParameterValue("ID", tempRightID);
                // Daten schreiben
                if (tempUpdate != "")
                {
                    dataReader = new SqlDB(aProjectID);
                    dataReader.execSQLwithParameter("UPDATE mapuserorgid1 SET " + tempUpdate + " WHERE ID=@ID", parameterList);
                }
            }
        }
        if (aSessionObj.filterOrgTableIndex == 2)
        {
            // Update-String konstruieren
            string tempUpdate = "";
            TParameterList parameterList = new TParameterList();
            tempUpdate = checkUpdate(UserID, "UserID", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(OrgID, "OrgID", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(OProfile, "OProfile", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(OReadLevel, "OReadLevel", tempUpdate, "-2", ref parameterList);
            tempUpdate = checkUpdate(OEditLevel, "OEditLevel", tempUpdate, "-2", ref parameterList);
            tempUpdate = checkUpdate(OAllowDelegate, "OAllowDelegate", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowLock, "OAllowLock", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowEmployeeRead, "OAllowEmployeeRead", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowEmployeeEdit, "OAllowEmployeeEdit", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowEmployeeImport, "OAllowEmployeeImport", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowEmployeeExport, "OAllowEmployeeExport", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowUnitAdd, "OAllowUnitAdd", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowUnitMove, "OAllowUnitMove", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowUnitDelete, "OAllowUnitDelete", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowUnitProperty, "OAllowUnitProperty", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowReportRecipient, "OAllowReportRecipient", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowStructureImport, "OAllowStructureImport", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowStructureExport, "OAllowStructureExport", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FProfile, "FProfile", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(FReadLevel, "FReadLevel", tempUpdate, "-2", ref parameterList);
            tempUpdate = checkUpdate(FEditLevel, "FEditLevel", tempUpdate, "-2", ref parameterList);
            tempUpdate = checkUpdate(FSumLevel, "FSumLevel", tempUpdate, "-2", ref parameterList);
            tempUpdate = checkUpdate(FAllowDelete, "FAllowDelete", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowExcelExport, "FAllowExcelExport", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowReminderOnMeasure, "FAllowReminderOnMeasure", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowReminderOnUnit, "FAllowReminderOnUnit", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowUpUnitsMeasures, "FAllowUpUnitsMeasures", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowColleaguesMeasures, "FAllowColleaguesMeasures", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowThemes, "FAllowThemes", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowReminderOnThemes, "FAllowReminderOnTheme", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowReminderOnThemeUnit, "FAllowReminderOnThemeUnit", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowTopUnitsThemes, "FAllowTopUnitsThemes", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowColleaguesThemes, "FAllowColleaguesThemes", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowOrders, "FAllowOrders", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowReminderOnOrderUnit, "FAllowReminderOnOrderUnit", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(RProfile, "RProfile", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(RReadLevel, "RReadLevel", tempUpdate, "-2", ref parameterList);
            tempUpdate = checkUpdate(DProfile, "DProfile", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(DReadLevel, "DReadLevel", tempUpdate, "-2", ref parameterList);
            tempUpdate = checkUpdate(DAllowMail, "DAllowMail", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(DAllowDownload, "DAllowDownload", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(DAllowLog, "DAllowLog", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(DAllowType1, "DAllowType1", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(DAllowType2, "DAllowType2", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(DAllowType3, "DAllowType3", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(DAllowType4, "DAllowType4", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(DAllowType5, "DAllowType5", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(AEditLevelWorkshop, "AEditLevelWorkshop", tempUpdate, "-2", ref parameterList);
            tempUpdate = checkUpdate(AReadLevelStructure, "AReadLevelStructure", tempUpdate, "-2", ref parameterList);

            parameterList.addParameter("ID", "string", "");

            // Schleife über alle Berechtigungen
            int index = 0;
            foreach (string tempRightID in aSessionObj.selectedIDList)
            {
                index++;
                TProcessStatus.updateProcessInfo(aSessionObj.ProcessID, 2, index, 0, aSessionObj.selectedIDList.Count, aSessionObj.Project.ProjectID);

                parameterList.changeParameterValue("ID", tempRightID);
                // Daten schreiben
                if (tempUpdate != "")
                {
                    dataReader = new SqlDB(aProjectID);
                    dataReader.execSQLwithParameter("UPDATE mapuserorgid2 SET " + tempUpdate + " WHERE ID=@ID", parameterList);
                }
            }
        }
        if (aSessionObj.filterOrgTableIndex == 3)
        {
            // Update-String konstruieren
            string tempUpdate = "";
            TParameterList parameterList = new TParameterList();
            tempUpdate = checkUpdate(UserID, "UserID", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(OrgID, "OrgID", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(OProfile, "OProfile", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(OReadLevel, "OReadLevel", tempUpdate, "-2", ref parameterList);
            tempUpdate = checkUpdate(OEditLevel, "OEditLevel", tempUpdate, "-2", ref parameterList);
            tempUpdate = checkUpdate(OAllowDelegate, "OAllowDelegate", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowLock, "OAllowLock", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowEmployeeRead, "OAllowEmployeeRead", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowEmployeeEdit, "OAllowEmployeeEdit", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowEmployeeImport, "OAllowEmployeeImport", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowEmployeeExport, "OAllowEmployeeExport", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowUnitAdd, "OAllowUnitAdd", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowUnitMove, "OAllowUnitMove", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowUnitDelete, "OAllowUnitDelete", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowUnitProperty, "OAllowUnitProperty", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowReportRecipient, "OAllowReportRecipient", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowStructureImport, "OAllowStructureImport", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(OAllowStructureExport, "OAllowStructureExport", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FProfile, "FProfile", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(FReadLevel, "FReadLevel", tempUpdate, "-2", ref parameterList);
            tempUpdate = checkUpdate(FEditLevel, "FEditLevel", tempUpdate, "-2", ref parameterList);
            tempUpdate = checkUpdate(FSumLevel, "FSumLevel", tempUpdate, "-2", ref parameterList);
            tempUpdate = checkUpdate(FAllowDelete, "FAllowDelete", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowExcelExport, "FAllowExcelExport", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowReminderOnMeasure, "FAllowReminderOnMeasure", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowReminderOnUnit, "FAllowReminderOnUnit", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowUpUnitsMeasures, "FAllowUpUnitsMeasures", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowColleaguesMeasures, "FAllowColleaguesMeasures", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowThemes, "FAllowThemes", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowReminderOnThemes, "FAllowReminderOnTheme", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowReminderOnThemeUnit, "FAllowReminderOnThemeUnit", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowTopUnitsThemes, "FAllowTopUnitsThemes", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowColleaguesThemes, "FAllowColleaguesThemes", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowOrders, "FAllowOrders", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(FAllowReminderOnOrderUnit, "FAllowReminderOnOrderUnit", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(RProfile, "RProfile", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(RReadLevel, "RReadLevel", tempUpdate, "-2", ref parameterList);
            tempUpdate = checkUpdate(DProfile, "DProfile", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(DReadLevel, "DReadLevel", tempUpdate, "-2", ref parameterList);
            tempUpdate = checkUpdate(DAllowMail, "DAllowMail", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(DAllowDownload, "DAllowDownload", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(DAllowLog, "DAllowLog", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(DAllowType1, "DAllowType1", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(DAllowType2, "DAllowType2", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(DAllowType3, "DAllowType3", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(DAllowType4, "DAllowType4", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(DAllowType5, "DAllowType5", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(AEditLevelWorkshop, "AEditLevelWorkshop", tempUpdate, "-2", ref parameterList);
            tempUpdate = checkUpdate(AReadLevelStructure, "AReadLevelStructure", tempUpdate, "-2", ref parameterList);

            parameterList.addParameter("ID", "string", "");

            // Schleife über alle Berechtigungen
            int index = 0;
            foreach (string tempRightID in aSessionObj.selectedIDList)
            {
                index++;
                TProcessStatus.updateProcessInfo(aSessionObj.ProcessID, 2, index, 0, aSessionObj.selectedIDList.Count, aSessionObj.Project.ProjectID);

                parameterList.changeParameterValue("ID", tempRightID);
                // Daten schreiben
                if (tempUpdate != "")
                {
                    dataReader = new SqlDB(aProjectID);
                    dataReader.execSQLwithParameter("UPDATE mapuserorgid3 SET " + tempUpdate + " WHERE ID=@ID", parameterList);
                }
            }
        }
        if (aSessionObj.filterOrgTableIndex == 4)
        {
            // Update-String konstruieren
            string tempUpdate = "";
            TParameterList parameterList = new TParameterList();
            tempUpdate = checkUpdate(UserID, "UserID", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(OrgID, "OrgID", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(OProfile, "OProfile", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(OReadLevel, "OReadLevel", tempUpdate, "-2", ref parameterList);

            parameterList.addParameter("ID", "string", "");

            // Schleife über alle Berechtigungen
            int index = 0;
            foreach (string tempRightID in aSessionObj.selectedIDList)
            {
                index++;
                TProcessStatus.updateProcessInfo(aSessionObj.ProcessID, 2, index, 0, aSessionObj.selectedIDList.Count, aSessionObj.Project.ProjectID);

                parameterList.changeParameterValue("ID", tempRightID);
                // Daten schreiben
                if (tempUpdate != "")
                {
                    dataReader = new SqlDB(aProjectID);
                    dataReader.execSQLwithParameter("UPDATE mapuserback SET " + tempUpdate + " WHERE ID=@ID", parameterList);
                }
            }
        }
        if (aSessionObj.filterOrgTableIndex == 5)
        {
            // Update-String konstruieren
            string tempUpdate = "";
            TParameterList parameterList = new TParameterList();
            tempUpdate = checkUpdate(UserID, "UserID", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(OrgID, "OrgID", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(OProfile, "OProfile", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(OReadLevel, "OReadLevel", tempUpdate, "-2", ref parameterList);

            parameterList.addParameter("ID", "string", "");

            // Schleife über alle Berechtigungen
            int index = 0;
            foreach (string tempRightID in aSessionObj.selectedIDList)
            {
                index++;
                TProcessStatus.updateProcessInfo(aSessionObj.ProcessID, 2, index, 0, aSessionObj.selectedIDList.Count, aSessionObj.Project.ProjectID);

                parameterList.changeParameterValue("ID", tempRightID);
                // Daten schreiben
                if (tempUpdate != "")
                {
                    dataReader = new SqlDB(aProjectID);
                    dataReader.execSQLwithParameter("UPDATE mapuserback1 SET " + tempUpdate + " WHERE ID=@ID", parameterList);
                }
            }
        }
        if (aSessionObj.filterOrgTableIndex == 6)
        {
            // Update-String konstruieren
            string tempUpdate = "";
            TParameterList parameterList = new TParameterList();
            tempUpdate = checkUpdate(UserID, "UserID", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(OrgID, "OrgID", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(OProfile, "OProfile", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(OReadLevel, "OReadLevel", tempUpdate, "-2", ref parameterList);

            parameterList.addParameter("ID", "string", "");

            // Schleife über alle Berechtigungen
            int index = 0;
            foreach (string tempRightID in aSessionObj.selectedIDList)
            {
                index++;
                TProcessStatus.updateProcessInfo(aSessionObj.ProcessID, 2, index, 0, aSessionObj.selectedIDList.Count, aSessionObj.Project.ProjectID);

                parameterList.changeParameterValue("ID", tempRightID);
                // Daten schreiben
                if (tempUpdate != "")
                {
                    dataReader = new SqlDB(aProjectID);
                    dataReader.execSQLwithParameter("UPDATE mapuserback2 SET " + tempUpdate + " WHERE ID=@ID", parameterList);
                }
            }
        }
        if (aSessionObj.filterOrgTableIndex == 7)
        {
            // Update-String konstruieren
            string tempUpdate = "";
            TParameterList parameterList = new TParameterList();
            tempUpdate = checkUpdate(UserID, "UserID", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(OrgID, "OrgID", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(OProfile, "OProfile", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(OReadLevel, "OReadLevel", tempUpdate, "-2", ref parameterList);

            parameterList.addParameter("ID", "string", "");

            // Schleife über alle Berechtigungen
            int index = 0;
            foreach (string tempRightID in aSessionObj.selectedIDList)
            {
                index++;
                TProcessStatus.updateProcessInfo(aSessionObj.ProcessID, 2, index, 0, aSessionObj.selectedIDList.Count, aSessionObj.Project.ProjectID);

                parameterList.changeParameterValue("ID", tempRightID);
                // Daten schreiben
                if (tempUpdate != "")
                {
                    dataReader = new SqlDB(aProjectID);
                    dataReader.execSQLwithParameter("UPDATE mapuserback3 SET " + tempUpdate + " WHERE ID=@ID", parameterList);
                }
            }
        }
        if (aSessionObj.filterOrgTableIndex == 8)
        {
            // Update-String konstruieren
            string tempUpdate = "";
            TParameterList parameterList = new TParameterList();
            //tempUpdate = checkUpdate(UserID, "UserID", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(containerID, "containerID", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(TContainerProfile, "TContainerProfile", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(TAllowUpload, "allowUpload", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(TAllowDownload, "allowDownload", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(TAllowDeleteOwnFiles, "allowDeleteOwnFiles", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(TAllowDeleteAllFiles, "allowDeleteAllFiles", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(TAllowAccessOwnFilesWithoutPassword, "allowAccessOwnFilesWithoutPassword", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(TAllowAccessAllFilesWithoutPassword, "allowAccessAllFilesWithoutPassword", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(TAllowCreateFolder, "allowCreateFolder", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(TAllowDeleteOwnFolder, "allowDeleteOwnFolder", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(TAllowDeleteAllFolder, "allowDeleteAllFolder", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(TAllowResetPassword, "allowResetPassword", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(TAllowTakeOwnership, "allowTakeOwnership", tempUpdate, "-1", ref parameterList);

            parameterList.addParameter("ID", "string", "");

            // Schleife über alle Berechtigungen
            int index = 0;
            foreach (string tempRightID in aSessionObj.selectedIDList)
            {
                index++;
                TProcessStatus.updateProcessInfo(aSessionObj.ProcessID, 2, index, 0, aSessionObj.selectedIDList.Count, aSessionObj.Project.ProjectID);

                parameterList.changeParameterValue("ID", tempRightID);
                // Daten schreiben
                if (tempUpdate != "")
                {
                    dataReader = new SqlDB(aProjectID);
                    dataReader.execSQLwithParameter("UPDATE mapusercontainer SET " + tempUpdate + " WHERE ID=@ID", parameterList);
                }
            }
        }
        if (aSessionObj.filterOrgTableIndex == 9)
        {
            // Update-String konstruieren
            string tempUpdate = "";
            TParameterList parameterList = new TParameterList();
            //tempUpdate = checkUpdate(UserID, "UserID", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(TTaskProfile, "TTaskProfile", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(TAllowTaskRead, "AllowTaskRead", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(TAllowTaskEdit, "AllowTaskCreateEditDelete", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(TAllowTaskFileUpload, "AllowTaskFileUpload", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(TAllowTaskExport, "AllowTaskExport", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(TAllowTaskImport, "AllowTaskImport", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(TAllowContactRead, "AllowContactRead", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(TAllowContactEdit, "AllowContactCreateEditDelete", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(TAllowContactExport, "AllowContactExport", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(TAllowContactImport, "AllowContactImport", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(TTranslateProfile, "TTranslateProfile", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(TAllowTranslateRead, "AllowTranslateRead", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(TAllowTranslateEdit, "AllowTranslateEdit", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(TAllowTranslateRelease, "AllowTranslateRelease", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(TAllowTranslateNew, "AllowTranslateNew", tempUpdate, "-1", ref parameterList);
            tempUpdate = checkUpdate(TAllowTranslateDelete, "AllowTranslateDelete", tempUpdate, "-1", ref parameterList);

            parameterList.addParameter("ID", "string", "");

            // Schleife über alle Berechtigungen
            int index = 0;
            foreach (string tempRightID in aSessionObj.selectedIDList)
            {
                index++;
                TProcessStatus.updateProcessInfo(aSessionObj.ProcessID, 2, index, 0, aSessionObj.selectedIDList.Count, aSessionObj.Project.ProjectID);

                parameterList.changeParameterValue("ID", tempRightID);
                // Daten schreiben
                if (tempUpdate != "")
                {
                    dataReader = new SqlDB(aProjectID);
                    dataReader.execSQLwithParameter("UPDATE mapuserrights SET " + tempUpdate + " WHERE ID=@ID", parameterList);
                }
            }
        }
        if (aSessionObj.filterOrgTableIndex == 10)
        {
            // Update-String konstruieren
            string tempUpdate = "";
            TParameterList parameterList = new TParameterList();
            //tempUpdate = checkUpdate(UserID, "UserID", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(roleID, "roleID", tempUpdate, "-1", ref parameterList);

            parameterList.addParameter("ID", "string", "");

            // Schleife über alle Berechtigungen
            int index = 0;
            foreach (string tempRightID in aSessionObj.selectedIDList)
            {
                index++;
                TProcessStatus.updateProcessInfo(aSessionObj.ProcessID, 2, index, 0, aSessionObj.selectedIDList.Count, aSessionObj.Project.ProjectID);

                parameterList.changeParameterValue("ID", tempRightID);
                // Daten schreiben
                if (tempUpdate != "")
                {
                    dataReader = new SqlDB(aProjectID);
                    dataReader.execSQLwithParameter("UPDATE mapusertaskrole SET " + tempUpdate + " WHERE ID=@ID", parameterList);
                }
            }
        }
        if (aSessionObj.filterOrgTableIndex == 11)
        {
            // Update-String konstruieren
            string tempUpdate = "";
            TParameterList parameterList = new TParameterList();
            //tempUpdate = checkUpdate(UserID, "UserID", tempUpdate, ref parameterList, false);
            tempUpdate = checkUpdate(translateLanaguageID, "translateLanguageID", tempUpdate, "-1", ref parameterList);

            parameterList.addParameter("ID", "string", "");

            // Schleife über alle Berechtigungen
            int index = 0;
            foreach (string tempRightID in aSessionObj.selectedIDList)
            {
                index++;
                TProcessStatus.updateProcessInfo(aSessionObj.ProcessID, 2, index, 0, aSessionObj.selectedIDList.Count, aSessionObj.Project.ProjectID);

                parameterList.changeParameterValue("ID", tempRightID);
                // Daten schreiben
                if (tempUpdate != "")
                {
                    dataReader = new SqlDB(aProjectID);
                    dataReader.execSQLwithParameter("UPDATE mapusertranslatelanguage SET " + tempUpdate + " WHERE ID=@ID", parameterList);
                }
            }
        }
    }
}

