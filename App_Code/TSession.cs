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
using FlexCel.Core;
using FlexCel.XlsAdapter;

public class TExportConfig
{
    public MemoryStream outStream = new MemoryStream();
}
/// <summary>
/// Klasse der Verwaltung einer Session
/// </summary>
/// 
public class TSessionObj
{
    public int AppVersion;              // Version
    public string Domain;
    public string StartApp;
    public string AdminApp;
    public string MailServerAdress;		// Adresse des zu verwendenden Mailservers
    public string MailServerUser;		// User für Login auf MailServer
    public string MailServerPassword;		// Password für Login auf MailServer
    public string MailSenderAdress;     // Absender-Emailadresse für Emails
    public string MailSenderName;       // Absender-Anzeigename für Emails
    public string Language;				// aktuell ausgewählte Sprache
    public TProject Project;
    public TAdminUser User;
    public string ProjectDataPath;      // Pfad zur Basis der Dateiverzeichnisse der Projekte
    public string ProjectToolFolder;    // Name des aktuellen Tool-Verzeichnis
    public int configLanguageIndex;
    public string selectedAddOnFollowUp = "";
    public string selectedCategorieOrgManager = "";
    public string selectedCategoriesLanguage;
    public string selectedFieldLabelsLanguage;
    public int numberOfIngressTargetGroups;
    public bool indirectIngressTargetGroup;
    public string ingressTargetGroup;
    public TAdminStructure Root;
    public TAdminStructureBack RootBack;
    public TAdminStructureBack1 RootBack1;
    public TAdminStructureBack2 RootBack2;
    public TAdminStructureBack3 RootBack3;
    public TAdminStructure1 RootOrg1;
    public TAdminStructure2 RootOrg2;
    public TAdminStructure3 RootOrg3;
    public TAdminStructure ActStructure;
    public TAdminStructureBack ActStructureBack;
    public TAdminStructureBack1 ActStructureBack1;
    public TAdminStructureBack2 ActStructureBack2;
    public TAdminStructureBack3 ActStructureBack3;
    public TAdminStructure1 ActStructure1;
    public TAdminStructure2 ActStructure2;
    public TAdminStructure3 ActStructure3;
    public TAdminStructure RootStructure;
    public TStructureField actStructureField;
    public TEmployeeField actEmployeeField;
    public TCategorie actCategoriesField;
    public TFilter actFilter;
    public TFilterCategories actFilterCategories;
    public int actStructureFilter = 0;
    public string actStructureFilterValue = "";
    public int actStructureTabRows = 10;
    public int actUserToolsFilter = 0;
    public string actUserToolsFilterValue = "";
    public int actUserToolsTabRows = 10;
    public int actUserFilter = 0;
    public string actUserFilterValue = "";
    public int actUserTabRows = 10;
    public ExcelFile importExcelFile;
    public string importFile;           // Pfad und Dateiname der Excel-Import-Datei
    public bool importDictionary;
    public bool importOrgManagerDictionary;
    public bool importResponseDictionary;
    public bool importDownloadDictionary;
    public bool importFollowUpDictionary;
    public bool importAccountManagerDictionary;
    public bool importTeamspaceDictionary;
    public bool importEmployee;
    public int importEmployeeType;
    public bool importStructure;
    public int importStructureType;
    public bool importOStructure;
    public int importOStructureType;
    public bool importRStructure;
    public int importRStructureType;
    public bool importRStructure1;
    public int importRStructure1Type;
    public bool importRStructure2;
    public int importRStructure2Type;
    public bool importRStructure3;
    public int importRStructure3Type;
    public bool importDStructure;
    public int importDStructureType;
    public bool importFStructure;
    public int importFStructureType;
    public bool importFStructure1;
    public int importFStructure1Type;
    public bool importFStructure2;
    public int importFStructure2Type;
    public bool importFStructure3;
    public int importFStructure3Type;
    public bool importStructure1;
    public int importStructure1Type;
    public bool importStructure2;
    public int importStructure2Type;
    public bool importStructure3;
    public int importStructure3Type;
    public bool importStructureBack;
    public int importStructureBackType;
    public bool importStructureBack1;
    public int importStructureBack1Type;
    public bool importStructureBack2;
    public int importStructureBack2Type;
    public bool importStructureBack3;
    public int importStructureBack3Type;
    public bool importStatus;
    public int importStatusType;
    public bool importSperre;
    public int importSperreType;
    public bool importBounces;
    public int importBouncesType;
    public bool importLoginUser;
    public int importLoginUserType;
    public bool importUserOrgID;
    public int importUserOrgIDType;
    public bool importUserOrgID1;
    public int importUserOrgID1Type;
    public bool importUserOrgID2;
    public int importUserOrgID2Type;
    public bool importUserOrgID3;
    public int importUserOrgID3Type;
    public bool importUserBack;
    public int importUserBackType;
    public bool importUserBack1;
    public int importUserBack1Type;
    public bool importUserBack2;
    public int importUserBack2Type;
    public bool importUserBack3;
    public int importUserBack3Type;
    public bool importUserRights;
    public int importUserRightsType;
    public bool importUserGroups;
    public int importUserGroupsType;
    public bool importUserGroup;
    public int importUserGroupType;
    public bool importUserTaskRole;
    public int importUserTaskRoleType;
    public bool importUserContainer;
    public int importUserContainerType;
    public bool importTranslateLanguage;
    public int importTranslateLanguageType;
    public bool importUserLanguage;
    public int importUserLanguageType;
    public bool importOrgManagerProfile;
    public int importOrgManagerProfileType;
    public bool importOrgManagerBackProfile;
    public int importOrgManagerBackProfileType;
    public bool importOrgManager1Profile;
    public int importOrgManager1ProfileType;
    public bool importOrgManager1BackProfile;
    public int importOrgManager1BackProfileType;
    public bool importOrgManager2Profile;
    public int importOrgManager2ProfileType;
    public bool importOrgManager2BackProfile;
    public int importOrgManager2BackProfileType;
    public bool importOrgManager3Profile;
    public int importOrgManager3ProfileType;
    public bool importOrgManager3BackProfile;
    public int importOrgManager3BackProfileType;
    public bool importResponseProfile;
    public int importResponseProfileType;
    public bool importResponse1Profile;
    public int importResponse1ProfileType;
    public bool importResponse2Profile;
    public int importResponse2ProfileType;
    public bool importResponse3Profile;
    public int importResponse3ProfileType;
    public bool importDownloadProfile;
    public int importDownloadProfileType;
    public bool importDownload1Profile;
    public int importDownload1ProfileType;
    public bool importDownload2Profile;
    public int importDownload2ProfileType;
    public bool importDownload3Profile;
    public int importDownload3ProfileType;
    public bool importFollowUpProfile;
    public int importFollowUpProfileType;
    public bool importFollowUp1Profile;
    public int importFollowUp1ProfileType;
    public bool importFollowUp2Profile;
    public int importFollowUp2ProfileType;
    public bool importFollowUp3Profile;
    public int importFollowUp3ProfileType;
    public bool importTeamspaceProfileTasks;
    public int importTeamspaceProfileTasksType;
    public bool importTeamspaceProfileContacts;
    public int importTeamspaceProfileContactsType;
    public bool importTeamspaceProfileContainer;
    public int importTeamspaceProfileContainerType;
    public bool importTeamspaceProfileTranslate;
    public int importTeamspaceProfileTranslateType;
    public bool importLicences;
    public int importLicencesType;

    public bool importCheckPreselect;
    public ArrayList importExcelResult;             // Ergebnisliste beim Excel-Import
    public string backPage;                 // Seite für Rückkehr
    public TResponseRight actResponseRights;
    public TUserLanguage userlanguageList;
    public TRadioList statusList;
    public TCheckList ParticipantList;
    public TRadioList SourecList;
    public TCheckList FollowupAddValue1List;
    public TCheckList FollowupAddValue2List;
    public TRadioList FollowupAddValue3List;
    public TRadioList FollowupAddValue4List;
    public ArrayList selectedIDList;        // Liste von IDs
    public string selectedIDListTable;          // Typ 0: mapuserorgid; Typ 1: mapuserorgid1; ...
    public TTools tools;
    public string Filename;
    public TLanguages availableLanguages;
    public TMultiUser ActMultiUser;         // Zwichenspeicher für Userdaten bei Aufruf von PickStructure
    public TMultiRights ActMultiRights;         // Zwichenspeicher für Berechtigungen bei Aufruf von PickStructure
    public TUser.TOrgRight ActUserActRight; // Zwischenspeicher für aktuelles Recht bei Aufruf von PickStructure
    public TUser.TBackRight ActUserActBackRight; // Zwischenspeicher für aktuelles Recht bei Aufruf von PickStructure
    public TUser.TOrg1Right ActUserActOrg1Right; // Zwischenspeicher für aktuelles Recht bei Aufruf von PickStructure
    public TUser.TBack1Right ActUserActBack1Right; // Zwischenspeicher für aktuelles Recht bei Aufruf von PickStructure
    public TUser.TOrg2Right ActUserActOrg2Right; // Zwischenspeicher für aktuelles Recht bei Aufruf von PickStructure
    public TUser.TBack2Right ActUserActBack2Right; // Zwischenspeicher für aktuelles Recht bei Aufruf von PickStructure
    public TUser.TOrg3Right ActUserActOrg3Right; // Zwischenspeicher für aktuelles Recht bei Aufruf von PickStructure
    public TUser.TBack3Right ActUserActBack3Right; // Zwischenspeicher für aktuelles Recht bei Aufruf von PickStructure
    public TUser.TContainerRight ActUserActContainerRight; // Zwischenspeicher für aktuelles Recht bei Aufruf von PickStructure
    public TUser.TUserRight ActUserActUserRight;
    public ArrayList ActUserActTaskRoles;
    public TUser ActUser;                   // Zwichenspeicher für Userdaten bei Aufruf von PickStructure
    public bool filterStructureActive;
    public bool filterIncludeSubunits;
    public int filterIndex;
    public string filterText;
    public string filterTextUserID;
    public string filterTextOrgID;
    public string filterTextContainerID;
    public string filterTextRoleID;
    public string filterTextGroupID;
    public string filterTextTranslateLanguageID;
    public string filterTextName;
    public string filterTextFirstname;
    public string filterTextEmail;
    public string filterTextLanguage;
    public string filterTextGroup;
    public string filterTextComment;
    public string filterTextProfile;
    public string filterTextDelegate;
    public string filterTextIsInternal;
    public string filterTextActive;
    public string filterUserIDs;
    public string filterContainerIDs;
    public string filterRoleIDs;
    public string filterGroupID;
    public string filterTranslateLanguageIDs;
    public string filterEmails;
    public string filterOrgIDs;
    public string filterGroupIDs;
    public string filterTextRight1;
    public string filterTextRight2;
    public string filterTextRight3;
    public string filterTextRight4;
    public string filterTextRight5;
    public string filterTextRight6;
    public string filterTextRight7;
    public string filterOrgTable;
    public int filterOrgTableIndex;
    public int filterDdlRight1;
    public int filterDdlRight2;
    public int filterDdlRight3;
    public int filterDdlRight4;
    public int filterDdlRight5;
    public int filterDdlRight6;
    public int filterDdlRight7;
    public bool filterOrgManager;
    public bool filterResponse;
    public bool filterDownload;
    public bool filterFollowUp;
    public bool filterAccountManager;
    public bool filterTeamSpace;
    public ArrayList dbErrors;
    public ArrayList folderErrors;
    public ArrayList structureErrors;
    public ArrayList configErrors;
    public ArrayList textErrors;
    public ArrayList userErrors;
    public ArrayList contentErrors;
    public string ProcessIDStructureExport;
    public string ProcessIDUserExport;
    public string ProcessIDIngressExport;
    public string ProcessID;
    public string ProcessMessage;
    public string ProcessReturnPage;
    public bool doSendEmail;
    public string doEmailTemplate;
    public TExportConfig ExportConfig;
    public int mailCounter;
    public int markOrgID;
    public EO.Web.Grid actUserGrid;
    public string actMailTemplate;
    public DateTime actMailTime;
    public int actMailJob;
    public string logExportType;
    public string actStructureFieldsTable;
    public string actCategoriesTable;
    public string actFieldID;
    public string actValue;
    public TUsergroups userGroups;
    public TAccessCode AccessCode;
    public string postnominationExportType;
    public ArrayList postnominationExportGroup;
    public TTranslateLanguage translateLaguage;
    public string syncError;

    public TIngress webService;

    /// <summary>
    /// Objekt erzeugen und Initalisierung durchführen
    /// </summary>
    public TSessionObj()
    {
        filterOrgManager = true;
        filterResponse = true;
        filterDownload = true;
        filterFollowUp = true;
        filterAccountManager = true;
        filterTeamSpace = true;

        Project = new TProject();
        User = new TAdminUser();

        ExportConfig = new TExportConfig();
    }
    /// <summary>
    /// Eigenschaften aus der DB heraus aktualisieren
    /// </summary>
    public void reload()
    {
        userlanguageList = new TUserLanguage(Project.ProjectID);
        statusList = new TRadioList(Project.ProjectID, "orgmanager_status");

        ParticipantList = new TCheckList(Project.ProjectID, "followup_participants");
        SourecList = new TRadioList(Project.ProjectID, "followup_source");
        FollowupAddValue1List = new TCheckList(Project.ProjectID, "followup_addvalue1");
        FollowupAddValue2List = new TCheckList(Project.ProjectID, "followup_addavlue2");
        FollowupAddValue3List = new TRadioList(Project.ProjectID, "followup_addvalue3");
        FollowupAddValue4List = new TRadioList(Project.ProjectID, "followup_addvalue4");

        Root = null;
    }
}
