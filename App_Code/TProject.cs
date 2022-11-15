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
/// Klasse zur Verwaltung der Eigenschaften eines Projektes
/// </summary>
public class TProject
{
    public class TOrgmanagerConfig
    {
        public bool ShowStartPage;

        public bool ShowOrgEmployeeInputCount;
        public bool ShowOrgEmployeeInputCountTotal;
        public bool ShowOrgEmployeeCount;
        public bool ShowOrgEmployeeCountTotal;
        public bool ShowOrgBackComparison;
        public bool ShowOrgLanguageUser;

        public bool ShowOrg1Number;
        public bool ShowOrg1NumberTotal;
        public bool ShowOrg1EmployeeCount;
        public bool ShowOrg1EmployeeCountTotal;
        public bool ShowOrg1BackComparison;
        public bool ShowOrg1LanguageUser;

        public bool ShowOrg2Number;
        public bool ShowOrg2NumberTotal;
        public bool ShowOrg2EmployeeCount;
        public bool ShowOrg2EmployeeCountTotal;
        public bool ShowOrg2BackComparison;
        public bool ShowOrg2LanguageUser;

        public bool ShowOrg3Number;
        public bool ShowOrg3NumberTotal;
        public bool ShowOrg3EmployeeCount;
        public bool ShowOrg3EmployeeCountTotal;
        public bool ShowOrg3BackComparison;
        public bool ShowOrg3LanguageUser;

        public bool ShowDelegateReadLevel;
        public bool ShowDelegateEditLevel;
        public bool ShowDelegateAllowDelegate;
        public bool ShowDelegateAllowLock;
        public bool ShowDelegateAllowEmployeeRead;
        public bool ShowDelegateAllowEmployeeEdit;
        public bool ShowDelegateAllowEmployeeImport;
        public bool ShowDelegateAllowEmployeeExport;
        public bool ShowDelegateAllowUnitAdd;
        public bool ShowDelegateAllowUnitMove;
        public bool ShowDelegateAllowUnitDelete;
        public bool ShowDelegateAllowUnitProperty;
        public bool ShowDelegateAllowUnitImport;
        public bool ShowDelegateAllowUnitExport;
        public bool ShowDelegateAllowReportRecipient;
        public bool ShowDelegateGroup;
        public bool ShowDelegateTitle;
        public bool ShowDelegateStreet;
        public bool ShowDelegateZipCode;
        public bool ShowDelegateCity;
        public bool ShowDelegateState;
        public bool ShowDelegateDisplayname;
        public bool ShowDelegateAllowOrgBack;
        public bool ShowDelegateAllowOrg1;
        public bool ShowDelegateAllowOrg1Back;
        public bool ShowDelegateAllowOrg2;
        public bool ShowDelegateAllowOrg2Back;
        public bool ShowDelegateAllowOrg3;
        public bool ShowDelegateAllowOrg3Back;

        public bool ShowOrg1DelegateReadLevel;
        public bool ShowOrg1DelegateEditLevel;
        public bool ShowOrg1DelegateAllowLock;
        public bool ShowOrg1DelegateAllowEmployeeRead;
        public bool ShowOrg1DelegateAllowEmployeeEdit;
        public bool ShowOrg1DelegateAllowEmployeeImport;
        public bool ShowOrg1DelegateAllowEmployeeExport;
        public bool ShowOrg1DelegateAllowUnitAdd;
        public bool ShowOrg1DelegateAllowUnitMove;
        public bool ShowOrg1DelegateAllowUnitDelete;
        public bool ShowOrg1DelegateAllowUnitProperty;
        public bool ShowOrg1DelegateAllowUnitImport;
        public bool ShowOrg1DelegateAllowUnitExport;
        public bool ShowOrg1DelegateAllowReportRecipient;

        public bool ShowOrg2DelegateReadLevel;
        public bool ShowOrg2DelegateEditLevel;
        public bool ShowOrg2DelegateAllowLock;
        public bool ShowOrg2DelegateAllowEmployeeRead;
        public bool ShowOrg2DelegateAllowEmployeeEdit;
        public bool ShowOrg2DelegateAllowEmployeeImport;
        public bool ShowOrg2DelegateAllowEmployeeExport;
        public bool ShowOrg2DelegateAllowUnitAdd;
        public bool ShowOrg2DelegateAllowUnitMove;
        public bool ShowOrg2DelegateAllowUnitDelete;
        public bool ShowOrg2DelegateAllowUnitProperty;
        public bool ShowOrg2DelegateAllowUnitImport;
        public bool ShowOrg2DelegateAllowUnitExport;
        public bool ShowOrg2DelegateAllowReportRecipient;

        public bool ShowOrg3DelegateReadLevel;
        public bool ShowOrg3DelegateEditLevel;
        public bool ShowOrg3DelegateAllowLock;
        public bool ShowOrg3DelegateAllowEmployeeRead;
        public bool ShowOrg3DelegateAllowEmployeeEdit;
        public bool ShowOrg3DelegateAllowEmployeeImport;
        public bool ShowOrg3DelegateAllowEmployeeExport;
        public bool ShowOrg3DelegateAllowUnitAdd;
        public bool ShowOrg3DelegateAllowUnitMove;
        public bool ShowOrg3DelegateAllowUnitDelete;
        public bool ShowOrg3DelegateAllowUnitProperty;
        public bool ShowOrg3DelegateAllowUnitImport;
        public bool ShowOrg3DelegateAllowUnitExport;
        public bool ShowOrg3DelegateAllowReportRecipient;

        public bool ShowEmployeeOrgID1;
        public bool ShowEmployeeOrgID2;
        public bool ShowEmployeeOrgID3;
        public bool ShowEmployeeName;
        public bool ShowEmployeeFirstname;
        public bool ShowEmployeeTitle;
        public bool ShowEmployeeStreet;
        public bool ShowEmployeeZipcode;
        public bool ShowEmployeeCity;
        public bool ShowEmployeeState;
        public bool ShowEmployeeEmail;
        public bool ShowEmployeeLanguage;
        public bool ShowEmployeeGroupnumber;
        public bool ShowEmployeeGender;
        public bool ShowEmployeeRace;
        public bool ShowEmployeeAge;
        public bool ShowEmployeeTenure;
        public bool ShowEmployeeJoblevel;
        public bool ShowEmployeeJobfunction;
        public bool ShowEmployeeJobprofile;
        public bool ShowEmployeeCostcenter;
        public bool ShowEmployeeJobstatus;
        public bool ShowEmployeeSurveyType;
        public bool ShowEmployeeRadio1;
        public bool ShowEmployeeRadio2;
        public bool ShowEmployeeCheck1;
        public bool ShowEmployeeCheck2;
        public bool ShowEmployeeText1;
        public bool ShowEmployeeText2;
        public bool ShowEmployeeTrustee;
        public bool ShowEmployeeSupervisor;

        public TOrgmanagerConfig()
        {
        }
        public TOrgmanagerConfig(string aProjectID)
        {
            SqlDB dataReader = new SqlDB(aProjectID);
            ShowStartPage = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowStartPage'").intValue == 1;

            ShowOrgEmployeeInputCount = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrgEmployeeInputCount'").intValue == 1;
            ShowOrgEmployeeInputCountTotal = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrgEmployeeInputCountTotal'").intValue == 1;
            ShowOrg1Number = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg1EmployeeInputCount'").intValue == 1;
            ShowOrg1NumberTotal = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg1EmployeeInputCountTotal'").intValue == 1;
            ShowOrg2Number = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg2EmployeeInputCount'").intValue == 1;
            ShowOrg2NumberTotal = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg2EmployeeInputCountTotal'").intValue == 1;
            ShowOrg3Number = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg3EmployeeInputCount'").intValue == 1;
            ShowOrg3NumberTotal = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg3EmployeeInputCountTotal'").intValue == 1;

            ShowDelegateReadLevel = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowDelegateReadLevel'").intValue == 1;
            ShowDelegateEditLevel = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowDelegateEditLevel'").intValue == 1;
            ShowDelegateAllowDelegate = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowDelegateAllowDelegate'").intValue == 1;
            ShowDelegateAllowLock = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowDelegateAllowLock'").intValue == 1;
            ShowDelegateAllowEmployeeRead = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowDelegateAllowEmployeeRead'").intValue == 1;
            ShowDelegateAllowEmployeeEdit = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowDelegateAllowEmployeeEdit'").intValue == 1;
            ShowDelegateAllowEmployeeImport = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowDelegateAllowEmployeeImport'").intValue == 1;
            ShowDelegateAllowEmployeeExport = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowDelegateAllowEmployeeExport'").intValue == 1;
            ShowDelegateAllowUnitAdd = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowDelegateAllowUnitAdd'").intValue == 1;
            ShowDelegateAllowUnitMove = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowDelegateAllowUnitMove'").intValue == 1;
            ShowDelegateAllowUnitDelete = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowDelegateAllowUnitDelete'").intValue == 1;
            ShowDelegateAllowUnitProperty = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowDelegateAllowUnitProperty'").intValue == 1;
            ShowDelegateAllowReportRecipient = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowDelegateAllowReportRecipient'").intValue == 1;
            ShowDelegateAllowUnitImport = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowDelegateAllowUnitImport'").intValue == 1;
            ShowDelegateAllowUnitExport = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowDelegateAllowUnitExport'").intValue == 1;
            ShowDelegateTitle = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowDelegateTitle'").intValue == 1;
            ShowDelegateStreet = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowDelegateStreet'").intValue == 1;
            ShowDelegateZipCode = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowDelegateZipcode'").intValue == 1;
            ShowDelegateCity = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowDelegateCity'").intValue == 1;
            ShowDelegateState = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowDelegateState'").intValue == 1;
            ShowDelegateGroup = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowDelegateGroup'").intValue == 1;
            ShowDelegateDisplayname = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowDelegateDisplayname'").intValue == 1;

            ShowOrg1DelegateReadLevel = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg1DelegateReadLevel'").intValue == 1;
            ShowOrg1DelegateEditLevel = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg1DelegateEditLevel'").intValue == 1;
            ShowOrg1DelegateAllowLock = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg1DelegateAllowLock'").intValue == 1;
            ShowOrg1DelegateAllowEmployeeRead = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg1DelegateAllowEmployeeRead'").intValue == 1;
            ShowOrg1DelegateAllowEmployeeEdit = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg1DelegateAllowEmployeeEdit'").intValue == 1;
            ShowOrg1DelegateAllowEmployeeImport = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg1DelegateAllowEmployeeImport'").intValue == 1;
            ShowOrg1DelegateAllowEmployeeExport = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg1DelegateAllowEmployeeExport'").intValue == 1;
            ShowOrg1DelegateAllowUnitAdd = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg1DelegateAllowUnitAdd'").intValue == 1;
            ShowOrg1DelegateAllowUnitMove = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg1DelegateAllowUnitMove'").intValue == 1;
            ShowOrg1DelegateAllowUnitDelete = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg1DelegateAllowUnitDelete'").intValue == 1;
            ShowOrg1DelegateAllowUnitProperty = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg1DelegateAllowUnitProperty'").intValue == 1;
            ShowOrg1DelegateAllowReportRecipient = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg1DelegateAllowReportRecipient'").intValue == 1;
            ShowOrg1DelegateAllowUnitImport = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg1DelegateAllowUnitImport'").intValue == 1;
            ShowOrg1DelegateAllowUnitExport = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg1DelegateAllowUnitExport'").intValue == 1;

            ShowOrg2DelegateReadLevel = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg2DelegateReadLevel'").intValue == 1;
            ShowOrg2DelegateEditLevel = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg2DelegateEditLevel'").intValue == 1;
            ShowOrg2DelegateAllowLock = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg2DelegateAllowLock'").intValue == 1;
            ShowOrg2DelegateAllowEmployeeRead = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg2DelegateAllowEmployeeRead'").intValue == 1;
            ShowOrg2DelegateAllowEmployeeEdit = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg2DelegateAllowEmployeeEdit'").intValue == 1;
            ShowOrg2DelegateAllowEmployeeImport = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg2DelegateAllowEmployeeImport'").intValue == 1;
            ShowOrg2DelegateAllowEmployeeExport = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg2DelegateAllowEmployeeExport'").intValue == 1;
            ShowOrg2DelegateAllowUnitAdd = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg2DelegateAllowUnitAdd'").intValue == 1;
            ShowOrg2DelegateAllowUnitMove = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg2DelegateAllowUnitMove'").intValue == 1;
            ShowOrg2DelegateAllowUnitDelete = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg2DelegateAllowUnitDelete'").intValue == 1;
            ShowOrg2DelegateAllowUnitProperty = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg2DelegateAllowUnitProperty'").intValue == 1;
            ShowOrg2DelegateAllowReportRecipient = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg2DelegateAllowReportRecipient'").intValue == 1;
            ShowOrg2DelegateAllowUnitImport = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg2DelegateAllowUnitImport'").intValue == 1;
            ShowOrg2DelegateAllowUnitExport = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg2DelegateAllowUnitExport'").intValue == 1;

            ShowOrg3DelegateReadLevel = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg3DelegateReadLevel'").intValue == 1;
            ShowOrg3DelegateEditLevel = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg3DelegateEditLevel'").intValue == 1;
            ShowOrg3DelegateAllowLock = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg3DelegateAllowLock'").intValue == 1;
            ShowOrg3DelegateAllowEmployeeRead = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg3DelegateAllowEmployeeRead'").intValue == 1;
            ShowOrg3DelegateAllowEmployeeEdit = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg3DelegateAllowEmployeeEdit'").intValue == 1;
            ShowOrg3DelegateAllowEmployeeImport = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg3DelegateAllowEmployeeImport'").intValue == 1;
            ShowOrg3DelegateAllowEmployeeExport = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg3DelegateAllowEmployeeExport'").intValue == 1;
            ShowOrg3DelegateAllowUnitAdd = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg3DelegateAllowUnitAdd'").intValue == 1;
            ShowOrg3DelegateAllowUnitMove = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg3DelegateAllowUnitMove'").intValue == 1;
            ShowOrg3DelegateAllowUnitDelete = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg3DelegateAllowUnitDelete'").intValue == 1;
            ShowOrg3DelegateAllowUnitProperty = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg3DelegateAllowUnitProperty'").intValue == 1;
            ShowOrg3DelegateAllowReportRecipient = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg3DelegateAllowReportRecipient'").intValue == 1;
            ShowOrg3DelegateAllowUnitImport = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg3DelegateAllowUnitImport'").intValue == 1;
            ShowOrg3DelegateAllowUnitExport = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowOrg3DelegateAllowUnitExport'").intValue == 1;

            ShowDelegateAllowOrgBack = ShowOrgBackComparison;
            ShowDelegateAllowOrg1 = ShowEmployeeOrgID1;
            ShowDelegateAllowOrg1Back = ShowOrg1BackComparison;
            ShowDelegateAllowOrg2 = ShowEmployeeOrgID2;
            ShowDelegateAllowOrg2Back = ShowOrg2BackComparison;
            ShowDelegateAllowOrg3 = ShowEmployeeOrgID3;
            ShowDelegateAllowOrg3Back = ShowOrg3BackComparison;

            //ShowEmployeeOrgID1 = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowEmployeeOrgID1'").intValue == 1;
            //ShowEmployeeOrgID2 = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowEmployeeOrgID2'").intValue == 1;
            //ShowEmployeeOrgID3 = dataReader.getScalarInt32("SELECT value FROM orgmanager_config WHERE ID='ShowEmployeeOrgID3'").intValue == 1;
        }
    }
    public class TFollowUpConfig
    {
        public bool ShowMeasures;               // Info, ob Maßnahmen erhoben werden solln
        public bool ShowCommunication;          // Info, on Ergebniskommunikation erhoben werden soll
        public bool ShowMeasuresTemplates;		// Info, ob Maßnahmenvorlagen verfügbar sein sollen
        public bool ShowMeasuresAdd;			// Info, ob eigene Maßnahmen hinzugefügt werden dürfen
        public bool ShowMeasuresRecom;			// Info, ob Maßnahmenempfehlungen verfügbar sein sollen
        public bool ShowMeasuresColl;			// Info, ob Maßnahmen von anderen Einheiten verfügbar sein sollen
        public bool ShowAddActions;
        public bool ShowPeople;
        public bool ShowMoney;
        public bool ShowTime;
        public bool ShowAddCheck1;
        public bool ShowAddCheck2;
        public bool ShowAddRadio1;
        public bool ShowAddRadio2;
        public bool ShowAddText1;
        public bool ShowAddText2;
        public bool ShowSupervisor;
        public bool ShowSupervisorEmail;
        public bool ShowSource;
        public bool ShowMeasuresPerUnit;
        public bool ShowActions;
        public bool ShowMeasureStart;
        public bool ShowMeasureEnd;
        public bool ShowMeasureDuration;
        public bool ShowEditHintPage;
        public bool ShowSummaryHintPage;
        public bool ShowSolution;
        public bool ShowSolutionStateConnection;
        public bool ShowReminderOnMeasure;
        public bool ShowReminderOnUnit;
        public bool ShowMeasuresToSubunits;
        public bool ShowMeasuresToColleagues;

        public bool ShowThemes;
        public bool ShowReminderOnTheme;
        public bool ShowReminderOnThemeUnit;
        public bool ShowThemesToColleagues;
        public bool ShowThemesToSubunits;
        public bool ShowActionsTheme;
        public bool ShowAddCheck1Theme;
        public bool ShowAddCheck2Theme;
        public bool ShowAddRadio1Theme;
        public bool ShowAddRadio2Theme;
        public bool ShowAddText1Theme;
        public bool ShowAddText2Theme;

        public bool ShowOrders;
        public bool ShowReminderOnOrderUnit;

        public TFollowUpConfig()
        {
        }
        public TFollowUpConfig(string aProjectID)
        {
            SqlDB dataReader = new SqlDB(aProjectID);
            ShowMeasures = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowMeasures'").intValue == 1;
            ShowCommunication = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowCommunication'").intValue == 1;
            ShowMeasuresTemplates = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowMeasureTemplates'").intValue == 1;
            ShowMeasuresAdd = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowAddMeasures'").intValue == 1;
            ShowMeasuresRecom = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowRecommendation'").intValue == 1;
            ShowMeasuresColl = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowPublish'").intValue == 1;
            ShowAddActions = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowAddActions'").intValue == 1;
            ShowPeople = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowPeople'").intValue == 1;
            ShowMoney = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowMoney'").intValue == 1;
            ShowTime = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowTime'").intValue == 1;
            ShowAddCheck1 = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowAddCheck1'").intValue == 1;
            ShowAddCheck2 = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowAddCheck2'").intValue == 1;
            ShowAddRadio1 = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowAddRadio1'").intValue == 1;
            ShowAddRadio2 = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowAddRadio2'").intValue == 1;
            ShowSupervisor = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowSupervisor'").intValue == 1;
            ShowSupervisorEmail = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowSupervisorEmail'").intValue == 1;
            ShowSource = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowSource'").intValue == 1;
            ShowMeasuresPerUnit = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowMeasuresPerUnit'").intValue == 1;
            ShowActions = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowActions'").intValue == 1;
            ShowAddText1 = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowAddText1'").intValue == 1;
            ShowAddText2 = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowAddText2'").intValue == 1;
            ShowMeasureStart = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowMeasureStart'").intValue == 1;
            ShowMeasureEnd = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowMeasureEnd'").intValue == 1;
            ShowMeasureDuration = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowMeasureDuration'").intValue == 1;
            ShowEditHintPage = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowEditHintPage'").intValue == 1;
            ShowSummaryHintPage = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowSummaryHintPage'").intValue == 1;
            ShowSolution = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowSolution'").intValue == 1;
            ShowSolutionStateConnection = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowSolutionStateConnection'").intValue == 1;
            ShowReminderOnMeasure = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowReminderOnMeasure'").intValue == 1;
            ShowReminderOnUnit = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowReminderOnUnit'").intValue == 1;
            ShowMeasuresToSubunits = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowMeasuresToSubunits'").intValue == 1;
            ShowMeasuresToColleagues = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowMeasuresToColleagues'").intValue == 1;

            ShowThemes = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowThemes'").intValue == 1;
            ShowReminderOnTheme = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowReminderOnTheme'").intValue == 1;
            ShowReminderOnThemeUnit = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowReminderOnThemeUnit'").intValue == 1;
            ShowThemesToColleagues = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowThemesToColleagues'").intValue == 1;
            ShowThemesToSubunits = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowThemesToSubunits'").intValue == 1;

            ShowActionsTheme = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowThemeActions'").intValue == 1;
            ShowAddCheck1Theme = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowThemeAddCheck1'").intValue == 1;
            ShowAddCheck1Theme = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowThemeAddCheck2'").intValue == 1;
            ShowAddRadio1Theme = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowThemeAddRadio1'").intValue == 1;
            ShowAddRadio1Theme = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowThemeAddRadio2'").intValue == 1;
            ShowAddText1Theme = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowThemeAddText1'").intValue == 1;
            ShowAddText1Theme = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowThemeAddText2'").intValue == 1;

            ShowOrders = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowOrders'").intValue == 1;
            ShowReminderOnOrderUnit = dataReader.getScalarInt32("SELECT value FROM followup_config WHERE ID='ShowReminderOnOrderUnit'").intValue == 1;
        }
    }
    public class TDownloadConfig
    {
        public bool ShowStartPage;

        public TDownloadConfig()
        {
        }
        public TDownloadConfig(string aProjectID)
        {
            SqlDB dataReader = new SqlDB(aProjectID);
            ShowStartPage = dataReader.getScalarInt32("SELECT value FROM download_config WHERE ID='ShowStartPage'").intValue == 1;

        }
    }
    public class TResponseConfig
    {
        public bool showDirectResponse;
        public bool showTotalResponse;
        public int cutoff;
        public int targetResponse;

        public TResponseConfig()
        {
        }
        public TResponseConfig(string aProjectID)
        {
            SqlDB dataReader = new SqlDB(aProjectID);
            showDirectResponse = dataReader.getScalarInt32("SELECT value FROM response_config WHERE ID='showDirectResponse'").intValue == 1;
            showTotalResponse = dataReader.getScalarInt32("SELECT value FROM response_config WHERE ID='showTotalResponse'").intValue == 1;
            cutoff = dataReader.getScalarInt32("SELECT value FROM response_config WHERE ID='cutoff'").intValue;
            targetResponse = dataReader.getScalarInt32("SELECT value FROM response_config WHERE ID='targetResponse'").intValue;
        }
    }
    public class TAccountManagerConfig
    {
        public bool ShowOrgManager;
        public bool ShowResponse;
        public bool ShowDownload;
        public bool ShowFollowUp;

        public bool AllowEditODelegate;
        public bool AllowEditOLock;
        public bool AllowEditOOrg1;
        public bool AllowEditOOrg2;
        public bool AllowEditOEmployee;
        public bool AllowEditOEmployeeImport;
        public bool AllowEditOEmployeeExport;
        public bool AllowEditOOrg;
        public bool AllowEditOReadLevel;
        public bool AllowEditOEditLevel;
        public bool AllowEditFOrg;
        public bool AllowEditFReadLevel;
        public bool AllowEditFEditLevel;
        public bool AllowEditFSumLevel;
        public bool AllowEditFDeleteMeasure;
        public bool AllowEditFExcelExport;
        public bool AllowEditFReminderMeasure;
        public bool AllowEditFReminderUnit;
        public bool AllowEditFUpUnitsMeasures;
        public bool AllowEditFColleagesMeasures;
        public bool AllowEditROrg;
        public bool AllowEditRReadLevel;
        public bool AllowEditRList;
        public bool AllowEditRStructure;
        public bool AllowEditRChart;
        public bool AllowEditDOrg;
        public bool AllowEditDReadLevel;
        public bool AllowEditDMail;
        public bool AllowEditDDownload;
        public bool AllowEditDProtocol;

        public TAccountManagerConfig()
        {
        }
        public TAccountManagerConfig(string aProjectID)
        {
            SqlDB dataReader = new SqlDB(aProjectID);
            ShowOrgManager = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='ShowOrgManager'").intValue == 1;
            ShowResponse = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='ShowResponse'").intValue == 1;
            ShowDownload = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='ShowDownload'").intValue == 1;
            ShowFollowUp = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='ShowFollowUp'").intValue == 1;

            AllowEditODelegate = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='AllowEditODelegate'").intValue == 1;
            AllowEditOLock = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='AllowEditOLock'").intValue == 1;
            AllowEditOOrg1 = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='AllowEditOOrg1'").intValue == 1;
            AllowEditOOrg2 = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='AllowEditOOrg2'").intValue == 1;
            AllowEditOEmployee = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='AllowEditOEmployee'").intValue == 1;
            AllowEditOEmployeeImport = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='AllowEditOEmployeeImport'").intValue == 1;
            AllowEditOEmployeeExport = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='ShowDelegateReadLevel'").intValue == 1;
            AllowEditOOrg = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='AllowEditOOrg'").intValue == 1;
            AllowEditOReadLevel = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='AllowEditOReadLevel'").intValue == 1;
            AllowEditOEditLevel = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='AllowEditOEditLevel'").intValue == 1;

            AllowEditFOrg = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='AllowEditFOrg'").intValue == 1;
            AllowEditFReadLevel = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='AllowEditFReadLevel'").intValue == 1;
            AllowEditFEditLevel = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='AllowEditFEditLevel'").intValue == 1;
            AllowEditFSumLevel = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='AllowEditFSumLevel'").intValue == 1;
            AllowEditFDeleteMeasure = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='AllowEditFDeleteMeasure'").intValue == 1;
            AllowEditFExcelExport = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='AllowEditFExcelExport'").intValue == 1;
            AllowEditFReminderMeasure = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='AllowEditFReminderMeasure'").intValue == 1;
            AllowEditFReminderUnit = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='AllowEditFReminderUnit'").intValue == 1;
            AllowEditFUpUnitsMeasures = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='AllowEditFUpUnitsMeasures'").intValue == 1;
            AllowEditFColleagesMeasures = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='AllowEditFColleaguesMeasures'").intValue == 1;

            AllowEditROrg = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='AllowEditROrg'").intValue == 1;
            AllowEditRReadLevel = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='AllowEditRReadLevel'").intValue == 1;
            AllowEditRList = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='AllowEditRList'").intValue == 1;
            AllowEditRStructure = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='AllowEditRStructure'").intValue == 1;
            AllowEditRChart = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='AllowEditRChart'").intValue == 1;

            AllowEditDOrg = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='AllowEditDOrg'").intValue == 1;
            AllowEditDReadLevel = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='AllowEditDReadLevel'").intValue == 1;
            AllowEditDMail = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='AllowEditDMail'").intValue == 1;
            AllowEditDDownload = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='AllowEditDDownload'").intValue == 1;
            AllowEditDProtocol = dataReader.getScalarInt32("SELECT value FROM accountmanager_config WHERE ID='AllowEditDProtocol'").intValue == 1;
        }
    }

    public string ProjectID;			// eindeutige Kennung des aktuellen Projektes
    public string ProjectDisplayName;	// Anzeigename des aktuellen Projektes
    public string ProjectReportDB;	// Anzeigename des aktuellen Projektes
    public bool Active;
    public bool Exists;
    public int sessionTimeout;          // Anzahl Minuten für den SessionTimeout
    public TUserPassword userPassword;
    //public bool showCaptcha;            // Verwendnung eines Captchas bei Anforderung Benutzername/Kennwort
    public string ingressEmail = "";
    public string ingressMobile = "";
    public string ingressEmployeeID = "";
    public string ingressOrgID = "";
    public string ingressOrgID1 = "";
    public string ingressOrgID2 = "";
    public string ingressOrgID3 = "";
    public string ingressProjectName = "";
    public string ingressQuestionaireName = "";
    public string ingressSurveyName = "";
    public string ingressPostNominationGroup = "";
    public TOrgmanagerConfig OrgManagerConfig;
    public TResponseConfig ResponseConfig;
    public TDownloadConfig DownloadConfig;
    public TFollowUpConfig FollowUpConfig;
    public TAccountManagerConfig AccountManagerConfig;
    public string masterLanguage;

    public TStructureFieldsList structureFieldsList;
    public TStructureFieldsFilter structureFieldsFilterList;
    public TStructureCategoriesFilter structureFieldsCategoriesFilterList;
    public TStructureFieldsList structure1FieldsList;
    public TStructureFieldsFilter structure1FieldsFilterList;
    public TStructureCategoriesFilter structure1FieldsCategoriesFilterList;
    public TStructureFieldsList structure2FieldsList;
    public TStructureFieldsFilter structure2FieldsFilterList;
    public TStructureCategoriesFilter structure2FieldsCategoriesFilterList;
    public TStructureFieldsList structure3FieldsList;
    public TStructureFieldsFilter structure3FieldsFilterList;
    public TStructureCategoriesFilter structure3FieldsCategoriesFilterList;
    public TEmployeeFieldsList employeeFieldsList;
    public TEmployeeFieldsFilter employeeFieldsFilterList;
    public TEmployeeCategoriesFilter employeeFieldsCategoriesFilterList;

    /// <summary>
    /// Leeres Projekt erzeugen
    /// </summary>
    public TProject()
    {
        OrgManagerConfig = new TOrgmanagerConfig();
        ResponseConfig = new TResponseConfig();
        DownloadConfig = new TDownloadConfig();
        FollowUpConfig = new TFollowUpConfig();
        AccountManagerConfig = new TAccountManagerConfig();
        // Default setzen
        sessionTimeout = 20;
        userPassword = null;
        ingressEmail = "";
        ingressMobile = "";
        ingressEmployeeID = "";
        ingressOrgID = "";
        ingressOrgID1 = "";
        ingressOrgID2 = "";
        ingressOrgID3 = "";
        ingressProjectName = "";
        ingressQuestionaireName = "";
        ingressSurveyName = "";
}
/// <summary>
/// Projekteigenschaften aus der DB lesen
/// </summary>
/// <param name="aProjectID">ID des Projektes</param>
public TProject(string aProjectID)
    {
        TSessionObj SessionObj = (TSessionObj)HttpContext.Current.Session["SessionObj"];

        Exists = false;
        ProjectID = aProjectID;
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("projectID", "string", aProjectID);
        SqlDB dataReader = new SqlDB("SELECT displayName, reportDB, active FROM projects WHERE projectID=@projectID", parameterList, "");
        if (dataReader.read())
        {
            ProjectDisplayName = dataReader.getString(0);
            ProjectReportDB = dataReader.getString(1);
            Active = dataReader.getBool(2);
            Exists = true;
        }
        dataReader.close();
        dataReader = new SqlDB(aProjectID);
        sessionTimeout = dataReader.getScalarInt32("SELECT value FROM config WHERE ID='sessionTimeout'").intValue;
        userPassword = new TUserPassword(aProjectID);
        ingressEmail = dataReader.getScalarString("SELECT text FROM config WHERE ID='ingressEmail'").stringValue;
        ingressMobile = dataReader.getScalarString("SELECT text FROM config WHERE ID='ingressMobile'").stringValue;
        ingressEmployeeID = dataReader.getScalarString("SELECT text FROM config WHERE ID='ingressEmployeeID'").stringValue;
        ingressOrgID = dataReader.getScalarString("SELECT text FROM config WHERE ID='ingressOrgID'").stringValue;
        ingressOrgID1 = dataReader.getScalarString("SELECT text FROM config WHERE ID='ingressOrgID1'").stringValue;
        ingressOrgID2 = dataReader.getScalarString("SELECT text FROM config WHERE ID='ingressOrgID2'").stringValue;
        ingressOrgID3 = dataReader.getScalarString("SELECT text FROM config WHERE ID='ingressOrgID3'").stringValue;
        ingressProjectName = dataReader.getScalarString("SELECT text FROM config WHERE ID='ingressProjectName'").stringValue;
        ingressQuestionaireName = dataReader.getScalarString("SELECT text FROM config WHERE ID='ingressQuestionaireName'").stringValue;
        ingressSurveyName = dataReader.getScalarString("SELECT text FROM config WHERE ID='ingressSurveyName'").stringValue;
        ingressPostNominationGroup = dataReader.getScalarString("SELECT text FROM config WHERE ID='ingressPostNominationGroup'").stringValue;
        masterLanguage = dataReader.getScalarString("SELECT text FROM config WHERE ID='masterLanguage'").stringValue;

        structureFieldsList = new TStructureFieldsList(ref SessionObj.Project.OrgManagerConfig.ShowOrgEmployeeCount, ref SessionObj.Project.OrgManagerConfig.ShowOrgEmployeeCountTotal, ref SessionObj.Project.OrgManagerConfig.ShowOrgBackComparison, ref SessionObj.Project.OrgManagerConfig.ShowOrgLanguageUser, "orgmanager_structurefields", aProjectID);
        structure1FieldsList = new TStructureFieldsList(ref SessionObj.Project.OrgManagerConfig.ShowOrg1EmployeeCount, ref SessionObj.Project.OrgManagerConfig.ShowOrg1EmployeeCountTotal, ref SessionObj.Project.OrgManagerConfig.ShowOrg1BackComparison, ref SessionObj.Project.OrgManagerConfig.ShowOrg1LanguageUser, "orgmanager_structure1fields", aProjectID);
        structure2FieldsList = new TStructureFieldsList(ref SessionObj.Project.OrgManagerConfig.ShowOrg2EmployeeCount, ref SessionObj.Project.OrgManagerConfig.ShowOrg2EmployeeCountTotal, ref SessionObj.Project.OrgManagerConfig.ShowOrg2BackComparison, ref SessionObj.Project.OrgManagerConfig.ShowOrg2LanguageUser, "orgmanager_structure2fields", aProjectID);
        structure3FieldsList = new TStructureFieldsList(ref SessionObj.Project.OrgManagerConfig.ShowOrg3EmployeeCount, ref SessionObj.Project.OrgManagerConfig.ShowOrg3EmployeeCountTotal, ref SessionObj.Project.OrgManagerConfig.ShowOrg3BackComparison, ref SessionObj.Project.OrgManagerConfig.ShowOrg3LanguageUser, "orgmanager_structure3fields", aProjectID);
        structureFieldsFilterList = new TStructureFieldsFilter("orgmanager_structurefieldsfilter", aProjectID);
        structure1FieldsFilterList = new TStructureFieldsFilter("orgmanager_structure1fieldsfilter", aProjectID);
        structure2FieldsFilterList = new TStructureFieldsFilter("orgmanager_structure2fieldsfilter", aProjectID);
        structure3FieldsFilterList = new TStructureFieldsFilter("orgmanager_structure3fieldsfilter", aProjectID);
        structureFieldsCategoriesFilterList = new TStructureCategoriesFilter("orgmanager_structurefieldscategoriesfilter", aProjectID);
        structure1FieldsCategoriesFilterList = new TStructureCategoriesFilter("orgmanager_structure1fieldscategoriesfilter", aProjectID);
        structure2FieldsCategoriesFilterList = new TStructureCategoriesFilter("orgmanager_structure2fieldscategoriesfilter", aProjectID);
        structure3FieldsCategoriesFilterList = new TStructureCategoriesFilter("orgmanager_structure3fieldscategoriesfilter", aProjectID);
        employeeFieldsList = new TEmployeeFieldsList(aProjectID);
        employeeFieldsFilterList = new TEmployeeFieldsFilter(aProjectID);
        employeeFieldsCategoriesFilterList = new TEmployeeCategoriesFilter(aProjectID);

        OrgManagerConfig = new TOrgmanagerConfig(aProjectID);
        ResponseConfig = new TResponseConfig(aProjectID);
        DownloadConfig = new TDownloadConfig(aProjectID);
        FollowUpConfig = new TFollowUpConfig(aProjectID);
        AccountManagerConfig = new TAccountManagerConfig(aProjectID);
    }
    protected void setValue(string aValue, string aKey, string aProjectID)
    {
        SqlDB dataReader;
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQL("DELETE FROM config WHERE ID='" + aKey + "'");
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQL("INSERT INTO config (ID, value) VALUES ('" + aKey + "', '" + aValue + "')");
    }
    protected void setString(string aString, string aKey, string aProjectID)
    {
        SqlDB dataReader;
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQL("DELETE FROM config WHERE ID='" + aKey + "'");
        dataReader = new SqlDB(aProjectID);
        dataReader.execSQL("INSERT INTO config (ID, text) VALUES ('" + aKey + "', '" + aString + "')");
    }
    /// <summary>
    /// Projekteigenschaften (Anzeiegname und Aktiv-Status) in der DB aktualisieren
    /// </summary>
    /// <param name="aDisplayName">zu setzender Anzeigename des Projektes</param>
    /// <param name="aActive">zu setzender Aktiv-Status des Projektes</param>
    public void update(string aDisplayName, string aReportDB, bool aActive, string aSessionTimeout, string aIngressProjectName, string aIngressQuestionaireName, string aIngressSurveyName, string aIngressPostNominationGroup, string aIngressOrgID, string aIngressOrgID1, string aIngressOrgID2, string aIngressOrgID3, string aIngressEmployeeID, string aIngressEmail, string aIngressMobile)
    {
        ProjectDisplayName = aDisplayName;
        ProjectReportDB = aReportDB;
        Active = aActive;
        string tempBool = "0";
        if (Active)
            tempBool = "1";
        TParameterList parameterList = new TParameterList();
        parameterList.addParameter("projectID", "string", ProjectID);
        parameterList.addParameter("displayName", "string", ProjectDisplayName);
        parameterList.addParameter("reportDB", "string", ProjectReportDB);
        parameterList.addParameter("active", "int", tempBool);
        SqlDB dataReader = new SqlDB("UPDATE projects SET displayName=@displayName, reportDB=@reportDB, active=@active WHERE projectID=@projectID", parameterList, "");

        sessionTimeout = Convert.ToInt32(aSessionTimeout);

        ingressProjectName = aIngressProjectName;
        ingressQuestionaireName = aIngressQuestionaireName;
        ingressSurveyName = aIngressSurveyName;
        ingressPostNominationGroup = aIngressPostNominationGroup;
        ingressOrgID = aIngressOrgID;
        ingressOrgID1 = aIngressOrgID1;
        ingressOrgID2 = aIngressOrgID2;
        ingressOrgID3 = aIngressOrgID3;
        ingressEmployeeID = aIngressEmployeeID;
        ingressEmail = aIngressEmail;
        ingressMobile = aIngressMobile;
        setValue(sessionTimeout.ToString(), "sessionTimeout", ProjectID);
        setString(ingressEmail, "ingressEmail", ProjectID);
        setString(ingressMobile, "ingressMobile", ProjectID);
        setString(ingressEmployeeID, "ingressEmployeeID", ProjectID);
        setString(ingressOrgID, "ingressOrgID", ProjectID);
        setString(ingressOrgID1, "ingressOrgID1", ProjectID);
        setString(ingressOrgID2, "ingressOrgID2", ProjectID);
        setString(ingressOrgID3, "ingressOrgID3", ProjectID);
        setString(ingressProjectName, "ingressProjectName", ProjectID);
        setString(ingressQuestionaireName, "ingressQuestionaireName", ProjectID);
        setString(ingressSurveyName, "ingressSurveyName", ProjectID);
        setString(ingressPostNominationGroup, "ingressPostNominationGroup", ProjectID);
    }
}
