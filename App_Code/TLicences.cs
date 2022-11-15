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
public class TLicences
{
    public class TLicenceSet
    {
        public int User = 0;
        public int UserOrgManager = 0;
        public int UserOrgManagerRights = 0;
        public int UserResponse = 0;
        public int UserResponseRights = 0;
        public int UserDownload = 0;
        public int UserDownloadRights = 0;
        public int UserFollowUp = 0;
        public int UserFollowUpCommunication = 0;
        public int UserFollowUpCommunicationRights = 0;
        public int UserFollowUpMeasure = 0;
        public int UserFollowUpMeasureRights = 0;
        public int UserFollowUpTheme = 0;
        public int UserFollowUpThemeRights = 0;
        public int UserFollowUpOrder = 0;
        public int UserFollowUpOrderRights = 0;
        public int UserTeamspace = 0;
        public int UserTeamspaceProjectplan = 0;
        public int UserTeamspaceContacts = 0;
        public int UserTeamspaceDataexchange = 0;
        public int UserTeamspaceDataexchangeRights = 0;
        public int UserAccountManager = 0;
        public int UserAccountManagerAccountTransfer = 0;
        public int UserAccountManagerAccountTransferRights = 0;
        public int UserAccountManagerAccountManagement = 0;
        public int UserAccountManagerAccountManagementRights = 0;
        public int UserAccountManagerWorkshop = 0;
        public int UserAccountManagerWorkshopRights = 0;
        public int UserAccountManagerStructure = 0;
        public int UserAccountManagerStructureRights = 0;
        public int UserOnlineReporting = 0;
    }
    public TLicenceSet LicencesUsed;
    public TLicenceSet LicencesAmount;
    public TLicenceSet LicencesAvailable;

    public TLicences(string aProjectID)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        LicencesAmount = new TLicenceSet();
        LicencesUsed = new TLicenceSet();
        LicencesAvailable = new TLicenceSet();

        // verfügbare Lizenzen ermitteln
        SqlDB dataReader = new SqlDB(aProjectID);
        LicencesAmount.User = dataReader.getScalarInt32("SELECT value FROM licences WHERE ID='User'").intValue;
        LicencesAmount.UserOrgManager = dataReader.getScalarInt32("SELECT value FROM licences WHERE ID='UserOrgManager'").intValue;
        LicencesAmount.UserOrgManagerRights = dataReader.getScalarInt32("SELECT value FROM licences WHERE ID='UserOrgManagerRights'").intValue;
        LicencesAmount.UserResponse = dataReader.getScalarInt32("SELECT value FROM licences WHERE ID='UserResponse'").intValue;
        LicencesAmount.UserResponseRights = dataReader.getScalarInt32("SELECT value FROM licences WHERE ID='UserResponseRights'").intValue;
        LicencesAmount.UserDownload = dataReader.getScalarInt32("SELECT value FROM licences WHERE ID='UserDownload'").intValue;
        LicencesAmount.UserDownloadRights = dataReader.getScalarInt32("SELECT value FROM licences WHERE ID='UserDownloadRights'").intValue;
        LicencesAmount.UserFollowUp = dataReader.getScalarInt32("SELECT value FROM licences WHERE ID='UserFollowUp'").intValue;
        LicencesAmount.UserFollowUpCommunication = dataReader.getScalarInt32("SELECT value FROM licences WHERE ID='UserFollowUpCommunication'").intValue;
        LicencesAmount.UserFollowUpCommunicationRights = dataReader.getScalarInt32("SELECT value FROM licences WHERE ID='UserFollowUpCommunicationRights'").intValue;
        LicencesAmount.UserFollowUpMeasure = dataReader.getScalarInt32("SELECT value FROM licences WHERE ID='UserFollowUpMeasure'").intValue;
        LicencesAmount.UserFollowUpMeasureRights = dataReader.getScalarInt32("SELECT value FROM licences WHERE ID='UserFollowUpMeasureRights'").intValue;
        LicencesAmount.UserFollowUpTheme = dataReader.getScalarInt32("SELECT value FROM licences WHERE ID='UserFollowUpTheme'").intValue;
        LicencesAmount.UserFollowUpThemeRights = dataReader.getScalarInt32("SELECT value FROM licences WHERE ID='UserFollowUpThemeRights'").intValue;
        LicencesAmount.UserFollowUpOrder = dataReader.getScalarInt32("SELECT value FROM licences WHERE ID='UserFollowUpOrder'").intValue;
        LicencesAmount.UserFollowUpOrderRights = dataReader.getScalarInt32("SELECT value FROM licences WHERE ID='UserFollowUpOrderRights'").intValue;
        LicencesAmount.UserTeamspace = dataReader.getScalarInt32("SELECT value FROM licences WHERE ID='UserTeamspace'").intValue;
        LicencesAmount.UserTeamspaceProjectplan = dataReader.getScalarInt32("SELECT value FROM licences WHERE ID='UserTeamspaceProjectplan'").intValue;
        LicencesAmount.UserTeamspaceContacts = dataReader.getScalarInt32("SELECT value FROM licences WHERE ID='UserTeamspaceContacts'").intValue;
        LicencesAmount.UserTeamspaceDataexchange = dataReader.getScalarInt32("SELECT value FROM licences WHERE ID='UserTeamspaceDataexchange'").intValue;
        LicencesAmount.UserTeamspaceDataexchangeRights = dataReader.getScalarInt32("SELECT value FROM licences WHERE ID='UserTeamspaceDataexchangeRights'").intValue;
        LicencesAmount.UserAccountManager = dataReader.getScalarInt32("SELECT value FROM licences WHERE ID='UserAccountManager'").intValue;
        LicencesAmount.UserAccountManagerAccountTransfer = dataReader.getScalarInt32("SELECT value FROM licences WHERE ID='UserAccountManagerAccountTransfer'").intValue;
        LicencesAmount.UserAccountManagerAccountTransferRights = dataReader.getScalarInt32("SELECT value FROM licences WHERE ID='UserAccountManagerAccountTransferRights'").intValue;
        LicencesAmount.UserAccountManagerAccountManagement = dataReader.getScalarInt32("SELECT value FROM licences WHERE ID='UserAccountManagerAccountManagement'").intValue;
        LicencesAmount.UserAccountManagerAccountManagementRights = dataReader.getScalarInt32("SELECT value FROM licences WHERE ID='UserAccountManagerAccountManagementRights'").intValue;
        LicencesAmount.UserAccountManagerWorkshop = dataReader.getScalarInt32("SELECT value FROM licences WHERE ID='UserAccountManagerWorkshop'").intValue;
        LicencesAmount.UserAccountManagerWorkshopRights = dataReader.getScalarInt32("SELECT value FROM licences WHERE ID='UserAccountManagerWorkshopRights'").intValue;
        LicencesAmount.UserAccountManagerStructure = dataReader.getScalarInt32("SELECT value FROM licences WHERE ID='UserAccountManagerStructure'").intValue;
        LicencesAmount.UserAccountManagerStructureRights = dataReader.getScalarInt32("SELECT value FROM licences WHERE ID='UserAccountManagerStructureRights'").intValue;
        LicencesAmount.UserOnlineReporting = dataReader.getScalarInt32("SELECT value FROM licences WHERE ID='UserOnlineReporting'").intValue;

        // Lizenznutzung ermitteln
        TTools tools = new TTools(SessionObj.Language, SessionObj.Project.ProjectID);

        // Anzahl User
        dataReader = new SqlDB("SELECT COUNT(userID) FROM loginuser WHERE isinternaluser='0' AND loginuser.isdelegate='0'", SessionObj.Project.ProjectID);
        if (dataReader.read())
        {
            LicencesUsed.User = dataReader.getInt32(0);
        }
        //OrgManager
        if (tools.getActive("orgmanager"))
        {
            // User
            SqlDB dataReader1 = new SqlDB("SELECT COUNT(DISTINCT mapuserorgid.userID) FROM mapuserorgid INNER JOIN loginuser ON loginuser.userID = mapuserorgID.userID WHERE (mapuserorgid.OReadlevel<>'0' OR mapuserorgid.OEditLevel<>'0') AND (mapuserorgid.OProfile='') AND loginuser.isinternaluser='0' AND loginuser.isdelegate='0'", SessionObj.Project.ProjectID);
            if (dataReader1.read())
            {
                int check = dataReader1.getInt32(0);
                LicencesUsed.UserOrgManager = LicencesUsed.UserOrgManager + check;
            }
            dataReader1.close();
            // Berechtigungen
            dataReader1 = new SqlDB("SELECT COUNT(mapuserorgid.orgID) FROM mapuserorgid INNER JOIN loginuser ON loginuser.userID = mapuserorgID.userID WHERE (mapuserorgid.OReadlevel<>'0' OR mapuserorgid.OEditLevel<>'0') AND (mapuserorgid.OProfile='') AND loginuser.isinternaluser='0' AND loginuser.isinternaluser='0'", SessionObj.Project.ProjectID);
            if (dataReader1.read())
            {
                int check = dataReader1.getInt32(0);
                LicencesUsed.UserOrgManagerRights = LicencesUsed.UserOrgManagerRights + check;
            }
            dataReader1.close();
            // Profile User
            dataReader1 = new SqlDB("SELECT profileID FROM orgmanager_profiles WHERE OReadlevel<>'0' OR OEditLevel<>'0'", SessionObj.Project.ProjectID);
            while (dataReader1.read())
            {
                string actProfile = dataReader1.getString(0);
                // User
                SqlDB dataReader2 = new SqlDB("SELECT COUNT(DISTINCT mapuserorgid.userID) FROM mapuserorgid INNER JOIN loginuser ON loginuser.userID = mapuserorgID.userID WHERE (OProfile='" + actProfile + "') AND loginuser.isinternaluser='0' AND loginuser.isinternaluser='0'", SessionObj.Project.ProjectID);
                if (dataReader2.read())
                {
                    int check = dataReader2.getInt32(0);
                    LicencesUsed.UserOrgManager = LicencesUsed.UserOrgManager + check;
                }
                dataReader2.close();
                // Berechtigungen 
                dataReader2 = new SqlDB("SELECT COUNT(mapuserorgid.orgID) FROM mapuserorgid INNER JOIN loginuser ON loginuser.userID = mapuserorgID.userID WHERE (OProfile='" + actProfile + "') AND loginuser.isinternaluser='0' AND loginuser.isinternaluser='0'", SessionObj.Project.ProjectID);
                if (dataReader2.read())
                {
                    int check = dataReader2.getInt32(0);
                    LicencesUsed.UserOrgManagerRights = LicencesUsed.UserOrgManagerRights + check;
                }
                dataReader2.close();
            }
            dataReader1.close();
        }
        //Response
        if (tools.getActive("response"))
        {
            // User
            SqlDB dataReader1 = new SqlDB("SELECT COUNT(DISTINCT mapuserorgid.userID) FROM mapuserorgid INNER JOIN loginuser ON loginuser.userID = mapuserorgID.userID WHERE (mapuserorgid.RReadlevel<>'0') AND (mapuserorgid.RProfile='') AND loginuser.isinternaluser='0' AND loginuser.isdelegate='0'", SessionObj.Project.ProjectID);
            if (dataReader1.read())
            {
                int check = dataReader1.getInt32(0);
                LicencesUsed.UserResponse = LicencesUsed.UserResponse + check;
            }
            dataReader1.close();
            // Berechtigungen
            dataReader1 = new SqlDB("SELECT COUNT(mapuserorgid.orgID) FROM mapuserorgid INNER JOIN loginuser ON loginuser.userID = mapuserorgID.userID WHERE (mapuserorgid.RReadlevel<>'0') AND (mapuserorgid.RProfile='') AND loginuser.isinternaluser='0' AND loginuser.isinternaluser='0'", SessionObj.Project.ProjectID);
            if (dataReader1.read())
            {
                int check = dataReader1.getInt32(0);
                LicencesUsed.UserResponseRights = LicencesUsed.UserResponseRights + check;
            }
            dataReader1.close();
            // Profile User
            dataReader1 = new SqlDB("SELECT profileID FROM response_profiles WHERE RReadlevel<>'0'", SessionObj.Project.ProjectID);
            while (dataReader1.read())
            {
                string actProfile = dataReader1.getString(0);
                // User
                SqlDB dataReader2 = new SqlDB("SELECT COUNT(DISTINCT mapuserorgid.userID) FROM mapuserorgid INNER JOIN loginuser ON loginuser.userID = mapuserorgID.userID WHERE (RProfile='" + actProfile + "') AND loginuser.isinternaluser='0' AND loginuser.isinternaluser='0'", SessionObj.Project.ProjectID);
                if (dataReader2.read())
                {
                    int check = dataReader2.getInt32(0);
                    LicencesUsed.UserResponse = LicencesUsed.UserResponse + check;
                }
                dataReader2.close();
                // Berechtigungen 
                dataReader2 = new SqlDB("SELECT COUNT(mapuserorgid.orgID) FROM mapuserorgid INNER JOIN loginuser ON loginuser.userID = mapuserorgID.userID WHERE (RProfile='" + actProfile + "') AND loginuser.isinternaluser='0' AND loginuser.isinternaluser='0'", SessionObj.Project.ProjectID);
                if (dataReader2.read())
                {
                    int check = dataReader2.getInt32(0);
                    LicencesUsed.UserResponseRights = LicencesUsed.UserResponseRights + check;
                }
                dataReader2.close();
            }
            dataReader1.close();
        }
        //Download
        if (tools.getActive("download"))
        {
            // User
            SqlDB dataReader1 = new SqlDB("SELECT COUNT(DISTINCT mapuserorgid.userID) FROM mapuserorgid INNER JOIN loginuser ON loginuser.userID = mapuserorgID.userID WHERE (mapuserorgid.DReadlevel<>'0') AND (mapuserorgid.DProfile='') AND loginuser.isinternaluser='0' AND loginuser.isdelegate='0'", SessionObj.Project.ProjectID);
            if (dataReader1.read())
            {
                int check = dataReader1.getInt32(0);
                LicencesUsed.UserDownload = LicencesUsed.UserDownload + check;
            }
            dataReader1.close();
            // Berechtigungen
            dataReader1 = new SqlDB("SELECT COUNT(mapuserorgid.orgID) FROM mapuserorgid INNER JOIN loginuser ON loginuser.userID = mapuserorgID.userID WHERE (mapuserorgid.DReadlevel<>'0') AND (mapuserorgid.DProfile='') AND loginuser.isinternaluser='0' AND loginuser.isinternaluser='0'", SessionObj.Project.ProjectID);
            if (dataReader1.read())
            {
                int check = dataReader1.getInt32(0);
                LicencesUsed.UserDownloadRights = LicencesUsed.UserDownloadRights + check;
            }
            dataReader1.close();
            // Profile User
            dataReader1 = new SqlDB("SELECT profileID FROM download_profiles WHERE DReadlevel<>'0'", SessionObj.Project.ProjectID);
            while (dataReader1.read())
            {
                string actProfile = dataReader1.getString(0);
                // User
                SqlDB dataReader2 = new SqlDB("SELECT COUNT(DISTINCT mapuserorgid.userID) FROM mapuserorgid INNER JOIN loginuser ON loginuser.userID = mapuserorgID.userID WHERE (DProfile='" + actProfile + "') AND loginuser.isinternaluser='0' AND loginuser.isinternaluser='0'", SessionObj.Project.ProjectID);
                if (dataReader2.read())
                {
                    int check = dataReader2.getInt32(0);
                    LicencesUsed.UserDownload = LicencesUsed.UserDownload + check;
                }
                dataReader2.close();
                // Berechtigungen 
                dataReader2 = new SqlDB("SELECT COUNT(mapuserorgid.orgID) FROM mapuserorgid INNER JOIN loginuser ON loginuser.userID = mapuserorgID.userID WHERE (DProfile='" + actProfile + "') AND loginuser.isinternaluser='0' AND loginuser.isinternaluser='0'", SessionObj.Project.ProjectID);
                if (dataReader2.read())
                {
                    int check = dataReader2.getInt32(0);
                    LicencesUsed.UserDownloadRights = LicencesUsed.UserDownloadRights + check;
                }
                dataReader2.close();
            }
            dataReader1.close();
        }
        //FollowUp
        if (tools.getActive("followup"))
        {
            // Anzahl User auf FollowUp
            dataReader = new SqlDB("SELECT COUNT(DISTINCT mapuserorgid.userID) FROM mapuserorgid INNER JOIN loginuser ON loginuser.userID = mapuserorgID.userID WHERE (FProfile<>'' OR (FProfile='' AND (FAllowCommunications='1' OR FAllowMeasures='1' OR FAllowThemes='1' OR FAllowOrders='1'))) AND (mapuserorgid.FReadLevel<>'0' OR mapuserorgid.FEditLevel<>'0' OR mapuserorgid.FSumLevel<>'0') AND (mapuserorgid.FProfile='') AND loginuser.isinternaluser='0' AND loginuser.isdelegate='0'", SessionObj.Project.ProjectID);
            if (dataReader.read())
            {
                LicencesUsed.UserFollowUp = dataReader.getInt32(0);
            }
            dataReader.close();
            // User Communication
            SqlDB dataReader1 = new SqlDB("SELECT COUNT(DISTINCT mapuserorgid.userID) FROM mapuserorgid INNER JOIN loginuser ON loginuser.userID = mapuserorgID.userID WHERE FAllowCommunications='1' AND (mapuserorgid.FReadLevel<>'0' OR mapuserorgid.FEditLevel<>'0' OR mapuserorgid.FSumLevel<>'0') AND (mapuserorgid.FProfile='') AND loginuser.isinternaluser='0' AND loginuser.isdelegate='0'", SessionObj.Project.ProjectID);
            if (dataReader1.read())
            {
                int check = dataReader1.getInt32(0);
                LicencesUsed.UserFollowUpCommunication = LicencesUsed.UserFollowUpCommunication + check;
            }
            dataReader1.close();
            // Berechtigungen Communication
            dataReader1 = new SqlDB("SELECT COUNT(mapuserorgid.orgID) FROM mapuserorgid INNER JOIN loginuser ON loginuser.userID = mapuserorgID.userID WHERE FAllowCommunications='1' AND (mapuserorgid.FReadLevel<>'0' OR mapuserorgid.FEditLevel<>'0' OR mapuserorgid.FSumLevel<>'0') AND (mapuserorgid.FProfile='') AND loginuser.isinternaluser='0' AND loginuser.isinternaluser='0'", SessionObj.Project.ProjectID);
            if (dataReader1.read())
            {
                int check = dataReader1.getInt32(0);
                LicencesUsed.UserFollowUpCommunicationRights = LicencesUsed.UserFollowUpCommunicationRights + check;
            }
            dataReader1.close();
            // Profile User
            dataReader1 = new SqlDB("SELECT profileID FROM followup_profiles WHERE FAllowCommunications='1' AND (mapuserorgid.FReadLevel<>'0' OR mapuserorgid.FEditLevel<>'0' OR mapuserorgid.FSumLevel<>'0')", SessionObj.Project.ProjectID);
            while (dataReader1.read())
            {
                string actProfile = dataReader1.getString(0);
                // User
                SqlDB dataReader2 = new SqlDB("SELECT COUNT(DISTINCT mapuserorgid.userID) FROM mapuserorgid INNER JOIN loginuser ON loginuser.userID = mapuserorgID.userID WHERE (FProfile='" + actProfile + "') AND loginuser.isinternaluser='0' AND loginuser.isinternaluser='0'", SessionObj.Project.ProjectID);
                if (dataReader2.read())
                {
                    int check = dataReader2.getInt32(0);
                    LicencesUsed.UserFollowUpCommunication = LicencesUsed.UserFollowUpCommunication + check;
                }
                dataReader2.close();
                // Berechtigungen 
                dataReader2 = new SqlDB("SELECT COUNT(mapuserorgid.orgID) FROM mapuserorgid INNER JOIN loginuser ON loginuser.userID = mapuserorgID.userID WHERE (FProfile='" + actProfile + "') AND loginuser.isinternaluser='0' AND loginuser.isinternaluser='0'", SessionObj.Project.ProjectID);
                if (dataReader2.read())
                {
                    int check = dataReader2.getInt32(0);
                    LicencesUsed.UserFollowUpCommunicationRights = LicencesUsed.UserFollowUpCommunicationRights + check;
                }
                dataReader2.close();
            }
            dataReader1.close();
            // User Measures
            dataReader1 = new SqlDB("SELECT COUNT(DISTINCT mapuserorgid.userID) FROM mapuserorgid INNER JOIN loginuser ON loginuser.userID = mapuserorgID.userID WHERE FAllowMeasures='1' AND (mapuserorgid.FReadLevel<>'0' OR mapuserorgid.FEditLevel<>'0' OR mapuserorgid.FSumLevel<>'0') AND (mapuserorgid.FProfile='') AND loginuser.isinternaluser='0' AND loginuser.isdelegate='0'", SessionObj.Project.ProjectID);
            if (dataReader1.read())
            {
                int check = dataReader1.getInt32(0);
                LicencesUsed.UserFollowUpMeasure = LicencesUsed.UserFollowUpMeasure + check;
            }
            dataReader1.close();
            // Berechtigungen Measures
            dataReader1 = new SqlDB("SELECT COUNT(mapuserorgid.orgID) FROM mapuserorgid INNER JOIN loginuser ON loginuser.userID = mapuserorgID.userID WHERE FAllowMeasures='1' AND (mapuserorgid.FReadLevel<>'0' OR mapuserorgid.FEditLevel<>'0' OR mapuserorgid.FSumLevel<>'0') AND (mapuserorgid.FProfile='') AND loginuser.isinternaluser='0' AND loginuser.isinternaluser='0'", SessionObj.Project.ProjectID);
            if (dataReader1.read())
            {
                int check = dataReader1.getInt32(0);
                LicencesUsed.UserFollowUpMeasureRights = LicencesUsed.UserFollowUpMeasureRights + check;
            }
            dataReader1.close();
            // Profile User
            dataReader1 = new SqlDB("SELECT profileID FROM followup_profiles WHERE FAllowMeasures='1' AND (mapuserorgid.FReadLevel<>'0' OR mapuserorgid.FEditLevel<>'0' OR mapuserorgid.FSumLevel<>'0')", SessionObj.Project.ProjectID);
            while (dataReader1.read())
            {
                string actProfile = dataReader1.getString(0);
                // User
                SqlDB dataReader2 = new SqlDB("SELECT COUNT(DISTINCT mapuserorgid.userID) FROM mapuserorgid INNER JOIN loginuser ON loginuser.userID = mapuserorgID.userID WHERE (FProfile='" + actProfile + "') AND loginuser.isinternaluser='0' AND loginuser.isinternaluser='0'", SessionObj.Project.ProjectID);
                if (dataReader2.read())
                {
                    int check = dataReader2.getInt32(0);
                    LicencesUsed.UserFollowUpMeasure = LicencesUsed.UserFollowUpMeasure + check;
                }
                dataReader2.close();
                // Berechtigungen 
                dataReader2 = new SqlDB("SELECT COUNT(mapuserorgid.orgID) FROM mapuserorgid INNER JOIN loginuser ON loginuser.userID = mapuserorgID.userID WHERE (FProfile='" + actProfile + "') AND loginuser.isinternaluser='0' AND loginuser.isinternaluser='0'", SessionObj.Project.ProjectID);
                if (dataReader2.read())
                {
                    int check = dataReader2.getInt32(0);
                    LicencesUsed.UserFollowUpMeasureRights = LicencesUsed.UserFollowUpMeasureRights + check;
                }
                dataReader2.close();
            }
            dataReader1.close();
            // User Themes
            dataReader1 = new SqlDB("SELECT COUNT(DISTINCT mapuserorgid.userID) FROM mapuserorgid INNER JOIN loginuser ON loginuser.userID = mapuserorgID.userID WHERE FAllowThemes='1' AND (mapuserorgid.FReadLevel<>'0' OR mapuserorgid.FEditLevel<>'0' OR mapuserorgid.FSumLevel<>'0') AND (mapuserorgid.FProfile='') AND loginuser.isinternaluser='0' AND loginuser.isdelegate='0'", SessionObj.Project.ProjectID);
            if (dataReader1.read())
            {
                int check = dataReader1.getInt32(0);
                LicencesUsed.UserFollowUpTheme = LicencesUsed.UserFollowUpTheme + check;
            }
            dataReader1.close();
            // Berechtigungen Themes
            dataReader1 = new SqlDB("SELECT COUNT(mapuserorgid.orgID) FROM mapuserorgid INNER JOIN loginuser ON loginuser.userID = mapuserorgID.userID WHERE FAllowThemes='1' AND (mapuserorgid.FReadLevel<>'0' OR mapuserorgid.FEditLevel<>'0' OR mapuserorgid.FSumLevel<>'0') AND (mapuserorgid.FProfile='') AND loginuser.isinternaluser='0' AND loginuser.isinternaluser='0'", SessionObj.Project.ProjectID);
            if (dataReader1.read())
            {
                int check = dataReader1.getInt32(0);
                LicencesUsed.UserFollowUpThemeRights = LicencesUsed.UserFollowUpThemeRights + check;
            }
            dataReader1.close();
            // Profile User
            dataReader1 = new SqlDB("SELECT profileID FROM followup_profiles WHERE FAllowThemes='1' AND (mapuserorgid.FReadLevel<>'0' OR mapuserorgid.FEditLevel<>'0' OR mapuserorgid.FSumLevel<>'0')", SessionObj.Project.ProjectID);
            while (dataReader1.read())
            {
                string actProfile = dataReader1.getString(0);
                // User
                SqlDB dataReader2 = new SqlDB("SELECT COUNT(DISTINCT mapuserorgid.userID) FROM mapuserorgid INNER JOIN loginuser ON loginuser.userID = mapuserorgID.userID WHERE (FProfile='" + actProfile + "') AND loginuser.isinternaluser='0' AND loginuser.isinternaluser='0'", SessionObj.Project.ProjectID);
                if (dataReader2.read())
                {
                    int check = dataReader2.getInt32(0);
                    LicencesUsed.UserFollowUpTheme = LicencesUsed.UserFollowUpTheme + check;
                }
                dataReader2.close();
                // Berechtigungen 
                dataReader2 = new SqlDB("SELECT COUNT(mapuserorgid.orgID) FROM mapuserorgid INNER JOIN loginuser ON loginuser.userID = mapuserorgID.userID WHERE (FProfile='" + actProfile + "') AND loginuser.isinternaluser='0' AND loginuser.isinternaluser='0'", SessionObj.Project.ProjectID);
                if (dataReader2.read())
                {
                    int check = dataReader2.getInt32(0);
                    LicencesUsed.UserFollowUpThemeRights = LicencesUsed.UserFollowUpThemeRights + check;
                }
                dataReader2.close();
            }
            dataReader1.close();
            // User Themes
            dataReader1 = new SqlDB("SELECT COUNT(DISTINCT mapuserorgid.userID) FROM mapuserorgid INNER JOIN loginuser ON loginuser.userID = mapuserorgID.userID WHERE FAllowOrders='1' AND (mapuserorgid.FReadLevel<>'0' OR mapuserorgid.FEditLevel<>'0' OR mapuserorgid.FSumLevel<>'0') AND (mapuserorgid.FProfile='') AND loginuser.isinternaluser='0' AND loginuser.isdelegate='0'", SessionObj.Project.ProjectID);
            if (dataReader1.read())
            {
                int check = dataReader1.getInt32(0);
                LicencesUsed.UserFollowUpOrder = LicencesUsed.UserFollowUpOrder + check;
            }
            dataReader1.close();
            // Berechtigungen Orders
            dataReader1 = new SqlDB("SELECT COUNT(mapuserorgid.orgID) FROM mapuserorgid INNER JOIN loginuser ON loginuser.userID = mapuserorgID.userID WHERE FAllowOrders='1' AND (mapuserorgid.FReadLevel<>'0' OR mapuserorgid.FEditLevel<>'0' OR mapuserorgid.FSumLevel<>'0') AND (mapuserorgid.FProfile='') AND loginuser.isinternaluser='0' AND loginuser.isinternaluser='0'", SessionObj.Project.ProjectID);
            if (dataReader1.read())
            {
                int check = dataReader1.getInt32(0);
                LicencesUsed.UserFollowUpOrderRights = LicencesUsed.UserFollowUpOrderRights + check;
            }
            dataReader1.close();
            // Profile Orders
            dataReader1 = new SqlDB("SELECT profileID FROM followup_profiles WHERE FAllowOrders='1' AND (mapuserorgid.FReadLevel<>'0' OR mapuserorgid.FEditLevel<>'0' OR mapuserorgid.FSumLevel<>'0')", SessionObj.Project.ProjectID);
            while (dataReader1.read())
            {
                string actProfile = dataReader1.getString(0);
                // User
                SqlDB dataReader2 = new SqlDB("SELECT COUNT(DISTINCT mapuserorgid.userID) FROM mapuserorgid INNER JOIN loginuser ON loginuser.userID = mapuserorgID.userID WHERE (FProfile='" + actProfile + "') AND loginuser.isinternaluser='0' AND loginuser.isinternaluser='0'", SessionObj.Project.ProjectID);
                if (dataReader2.read())
                {
                    int check = dataReader2.getInt32(0);
                    LicencesUsed.UserFollowUpOrder = LicencesUsed.UserFollowUpOrder + check;
                }
                dataReader2.close();
                // Berechtigungen 
                dataReader2 = new SqlDB("SELECT COUNT(mapuserorgid.orgID) FROM mapuserorgid INNER JOIN loginuser ON loginuser.userID = mapuserorgID.userID WHERE (FProfile='" + actProfile + "') AND loginuser.isinternaluser='0' AND loginuser.isinternaluser='0'", SessionObj.Project.ProjectID);
                if (dataReader2.read())
                {
                    int check = dataReader2.getInt32(0);
                    LicencesUsed.UserFollowUpOrderRights = LicencesUsed.UserFollowUpOrderRights + check;
                }
                dataReader2.close();
            }
            dataReader1.close();
        }
        //Teamspace
        if (tools.getActive("teamspace"))
        {
            // Anzahl User auf Teamspace ### das stimmt so nicht, hier ist nur DataExchange berücksichtigt
            //dataReader = new SqlDB("SELECT COUNT(DISTINCT mapuserorgid.userID) FROM mapusercontainer INNER JOIN loginuser ON loginuser.userID = mapusercontainer.userID WHERE loginuser.isinternaluser='0' AND loginuser.isdelegate='0'", SessionObj.Project.ProjectID);
            //if (dataReader.read())
            //{
            //    LicencesUsed.UserTeamspace = dataReader.getInt32(0);
            //}
            //dataReader.close();
            // User Dataexchange
            SqlDB dataReader1 = new SqlDB("SELECT COUNT(DISTINCT mapusercontainer.userID) FROM mapusercontainer INNER JOIN loginuser ON loginuser.userID = mapusercontainer.userID WHERE loginuser.isinternaluser='0' AND loginuser.isdelegate='0'", SessionObj.Project.ProjectID);
            if (dataReader1.read())
            {
                int check = dataReader1.getInt32(0);
                LicencesUsed.UserTeamspaceDataexchange = LicencesUsed.UserTeamspaceDataexchange + check;
            }
            dataReader1.close();
            // Berechtigungen DataExchange
            dataReader1 = new SqlDB("SELECT COUNT(DISTINCT mapusercontainer.containerID) FROM mapusercontainer INNER JOIN loginuser ON loginuser.userID = mapusercontainer.userID WHERE AND loginuser.isinternaluser='0' AND loginuser.isinternaluser='0'", SessionObj.Project.ProjectID);
            if (dataReader1.read())
            {
                int check = dataReader1.getInt32(0);
                LicencesUsed.UserTeamspaceDataexchangeRights = LicencesUsed.UserTeamspaceDataexchangeRights + check;
            }
            dataReader1.close();
            // User ProjectPlan
            dataReader1 = new SqlDB("SELECT COUNT(DISTINCT mapuserrights.userID) FROM mapuserrights INNER JOIN loginuser ON loginuser.userID = mapuserrights.userID WHERE (TTaskProfile<>'' OR AllowTaskRead='1' OR AllowTaskCreateEditDelete='1' OR AllowTaskFileUpload='1' OR AllowTaskExport='1' OR AllowTaskImport='1') AND loginuser.isinternaluser='0' AND loginuser.isdelegate='0'", SessionObj.Project.ProjectID);
            if (dataReader1.read())
            {
                int check = dataReader1.getInt32(0);
                LicencesUsed.UserTeamspaceProjectplan = LicencesUsed.UserTeamspaceProjectplan + check;
            }
            dataReader1.close();
            // User ProjectPlan
            dataReader1 = new SqlDB("SELECT COUNT(DISTINCT mapuserrights.userID) FROM mapuserrights INNER JOIN loginuser ON loginuser.userID = mapuserrights.userID WHERE (TContactProfile<>'' OR AllowContactRead='1' OR AllowContactCreateEditDelete='1' OR AllowContactExport='1' OR AllowContactImport='1') AND loginuser.isinternaluser='0' AND loginuser.isdelegate='0'", SessionObj.Project.ProjectID);
            if (dataReader1.read())
            {
                int check = dataReader1.getInt32(0);
                LicencesUsed.UserTeamspaceContacts = LicencesUsed.UserTeamspaceContacts + check;
            }
            dataReader1.close();
        }
        //AccountManager
        if (tools.getActive("accountmanager"))
        {
            // Workshop User
            SqlDB dataReader1 = new SqlDB("SELECT COUNT(DISTINCT mapuserorgid.userID) FROM mapuserorgid INNER JOIN loginuser ON loginuser.userID = mapuserorgid.userID WHERE mapuserorgid.AEditLevelWorkshop<>'0' AND loginuser.isinternaluser='0' AND loginuser.isdelegate='0'", SessionObj.Project.ProjectID);
            if (dataReader1.read())
            {
                int check = dataReader1.getInt32(0);
                LicencesUsed.UserAccountManagerWorkshop = LicencesUsed.UserAccountManagerWorkshop + check;
            }
            dataReader1.close();
            // Workshop Berechtigungen
            dataReader1 = new SqlDB("SELECT COUNT(orgID) FROM mapuserorgid INNER JOIN loginuser ON loginuser.userID = mapuserorgid.userID WHERE mapuserorgid.AEditLevelWorkshop<>'0' AND loginuser.isinternaluser='0' AND loginuser.isdelegate='0'", SessionObj.Project.ProjectID);
            if (dataReader1.read())
            {
                int check = dataReader1.getInt32(0);
                LicencesUsed.UserAccountManagerWorkshopRights = LicencesUsed.UserAccountManagerWorkshopRights + check;
            }
            dataReader1.close();

            //###







        }

        //// internaluser berücksichtigen
        //dataReader = new SqlDB("SELECT userID FROM loginuser WHERE isinternaluser='0' AND loginuser.isdelegate='0'", SessionObj.Project.ProjectID);
        //while (dataReader.read())
        //{
        //    string actUserID = dataReader.getString(0);

        //    //AccountManager
        //    if (tools.getActive("accountmanager"))
        //    {
        //        bool found = false;

        //        SqlDB dataReader1 = new SqlDB("SELECT COUNT(DISTINCT orgID) FROM mapuserorgid WHERE userID='" + actUserID + "' AND AAllowAccountTransfer='1' AND (mapuserorgid.AEditLevel<>'0')", SessionObj.Project.ProjectID);
        //        if (dataReader1.read())
        //        {
        //            int check = dataReader1.getInt32(0);
        //            if (check > 0)
        //            {
        //                found = true;
        //                LicencesUsed.UserAccountManager++;
        //                LicencesUsed.UserAccountManagerAccountTransfer++;
        //            }
        //            LicencesUsed.UserAccountManagerAccountTransferRights = LicencesUsed.UserAccountManagerAccountTransferRights + dataReader1.getInt32(0);
        //        }
        //        dataReader1.close();

        //        dataReader1 = new SqlDB("SELECT COUNT(DISTINCT orgID) FROM mapuserorgid WHERE userID='" + actUserID + "' AND AAllowAccountManagement='1' AND (mapuserorgid.AEditLevel<>'0')", SessionObj.Project.ProjectID);
        //        if (dataReader1.read())
        //        {
        //            int check = dataReader1.getInt32(0);
        //            if (check > 0)
        //            {
        //                if (!found)
        //                {
        //                    found = true;
        //                    LicencesUsed.UserAccountManager++;
        //                }
        //                LicencesUsed.UserAccountManagerAccountManagement++;
        //            }
        //            LicencesUsed.UserAccountManagerAccountManagementRights = LicencesUsed.UserAccountManagerAccountManagementRights + dataReader1.getInt32(0);
        //        }
        //        dataReader1.close();

        //        dataReader1 = new SqlDB("SELECT COUNT(DISTINCT orgID) FROM mapuserorgid WHERE userID='" + actUserID + "' AND (mapuserorgid.AEditLevelStructure<>'0')", SessionObj.Project.ProjectID);
        //        if (dataReader1.read())
        //        {
        //            int check = dataReader1.getInt32(0);
        //            if (check > 0)
        //            {
        //                if (!found)
        //                {
        //                    found = true;
        //                    LicencesUsed.UserAccountManager++;
        //                }
        //                LicencesUsed.UserAccountManagerStructure++;
        //            }
        //            LicencesUsed.UserAccountManagerStructureRights = LicencesUsed.UserAccountManagerStructureRights + dataReader1.getInt32(0);
        //        }
        //        dataReader1.close();
        //    }
        //    //###
        //    LicencesUsed.UserOnlineReporting = 0;
        //}
        //dataReader.close();

        // verfügbare Lizenzen ermitteln
        LicencesAvailable.User = LicencesAmount.User - LicencesUsed.User;
        LicencesAvailable.UserOrgManager = LicencesAmount.UserOrgManager - LicencesUsed.UserOrgManager;
        LicencesAvailable.UserOrgManagerRights = LicencesAmount.UserOrgManagerRights - LicencesUsed.UserOrgManagerRights;
        LicencesAvailable.UserResponse = LicencesAmount.UserResponse - LicencesUsed.UserResponse;
        LicencesAvailable.UserResponseRights = LicencesAmount.UserResponseRights - LicencesUsed.UserResponseRights;
        LicencesAvailable.UserDownload = LicencesAmount.UserDownload - LicencesUsed.UserDownload;
        LicencesAvailable.UserDownloadRights = LicencesAmount.UserDownloadRights - LicencesUsed.UserDownloadRights;
        LicencesAvailable.UserFollowUp = LicencesAmount.UserFollowUp - LicencesUsed.UserFollowUp;
        LicencesAvailable.UserFollowUpCommunication = LicencesAmount.UserFollowUpCommunication - LicencesUsed.UserFollowUpCommunication;
        LicencesAvailable.UserFollowUpCommunicationRights = LicencesAmount.UserFollowUpCommunicationRights - LicencesUsed.UserFollowUpCommunicationRights;
        LicencesAvailable.UserFollowUpMeasure = LicencesAmount.UserFollowUpMeasure - LicencesUsed.UserFollowUpMeasure;
        LicencesAvailable.UserFollowUpMeasureRights = LicencesAmount.UserFollowUpMeasureRights - LicencesUsed.UserFollowUpMeasureRights;
        LicencesAvailable.UserFollowUpTheme = LicencesAmount.UserFollowUpTheme - LicencesUsed.UserFollowUpTheme;
        LicencesAvailable.UserFollowUpThemeRights = LicencesAmount.UserFollowUpThemeRights - LicencesUsed.UserFollowUpThemeRights;
        LicencesAvailable.UserFollowUpOrder = LicencesAmount.UserFollowUpOrder - LicencesUsed.UserFollowUpOrder;
        LicencesAvailable.UserFollowUpOrderRights = LicencesAmount.UserFollowUpOrderRights - LicencesUsed.UserFollowUpOrderRights;
        LicencesAvailable.UserTeamspace = LicencesAmount.UserTeamspace - LicencesUsed.UserTeamspace;
        LicencesAvailable.UserTeamspaceProjectplan = LicencesAmount.UserTeamspaceProjectplan - LicencesUsed.UserTeamspaceProjectplan;
        LicencesAvailable.UserTeamspaceContacts = LicencesAmount.UserTeamspaceContacts - LicencesUsed.UserTeamspaceContacts;
        LicencesAvailable.UserTeamspaceDataexchange = LicencesAmount.UserTeamspaceDataexchange - LicencesUsed.UserTeamspaceDataexchange;
        LicencesAvailable.UserTeamspaceDataexchangeRights = LicencesAmount.UserTeamspaceDataexchangeRights - LicencesUsed.UserTeamspaceDataexchangeRights;
        LicencesAvailable.UserAccountManager = LicencesAmount.UserAccountManager - LicencesUsed.UserAccountManager;
        LicencesAvailable.UserAccountManagerAccountTransfer = LicencesAmount.UserAccountManagerAccountTransfer - LicencesUsed.UserAccountManagerAccountTransfer;
        LicencesAvailable.UserAccountManagerAccountTransferRights = LicencesAmount.UserAccountManagerAccountTransferRights - LicencesUsed.UserAccountManagerAccountTransferRights;
        LicencesAvailable.UserAccountManagerAccountManagement = LicencesAmount.UserAccountManagerAccountManagement - LicencesUsed.UserAccountManagerAccountManagement;
        LicencesAvailable.UserAccountManagerAccountManagementRights = LicencesAmount.UserAccountManagerAccountManagementRights - LicencesUsed.UserAccountManagerAccountManagementRights;
        LicencesAvailable.UserAccountManagerWorkshop = LicencesAmount.UserAccountManagerWorkshop - LicencesUsed.UserAccountManagerWorkshop;
        LicencesAvailable.UserAccountManagerWorkshopRights = LicencesAmount.UserAccountManagerWorkshopRights - LicencesUsed.UserAccountManagerWorkshopRights;
        LicencesAvailable.UserAccountManagerStructure = LicencesAmount.UserAccountManagerStructure - LicencesUsed.UserAccountManagerStructure;
        LicencesAvailable.UserAccountManagerStructureRights = LicencesAmount.UserAccountManagerStructureRights - LicencesUsed.UserAccountManagerStructureRights;
        LicencesAvailable.UserOnlineReporting = LicencesAmount.UserOnlineReporting - LicencesUsed.UserOnlineReporting;
    }
}