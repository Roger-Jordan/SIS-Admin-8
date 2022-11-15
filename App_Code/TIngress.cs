using System;
using System.Collections;
using System.Collections.Generic;
using System.Web;
using System.Globalization;
public class TIngress
{
    public class TEmployeeData
    {
        public int employeeID;
        public Dictionary<string, string> values;
    }

    public webservice_keyingress webService;
    public string projectName;
    public int projectID;
    public string surveyName;
    public int surveyID;
    public string questionaireName;
    public int questionaireID;
    public string employeeIDField;
    public string mailField;
    public string mobileField;
    public string orgIDField;
    public string orgID1Field;
    public string orgID2Field;
    public string orgID3Field;
    public string postNominationField;
    public int postNominationFieldID;

    public TIngress(string aProjectID)
    {
        webService = new webservice_keyingress();
        // Authentifizierung im Webservice
        AuthenticationHeader authenticationHeader = new AuthenticationHeader();
        authenticationHeader.username = "WebServiceUser";
        authenticationHeader.password = "webservice#ipsos#1";
        webService.AuthenticationHeaderValue = authenticationHeader;

        // Konfiguration einlesen
        SqlDB dataReader = new SqlDB(aProjectID);
        // Zugriffsparameter ermimtteln
        projectName = dataReader.getScalarString("SELECT text FROM config WHERE ID='ingressProjectName'").stringValue;
        projectID = Convert.ToInt32(webService.get_system_id(0, false, projectName, "project"));
        surveyName = dataReader.getScalarString("SELECT text FROM config WHERE ID='ingressSurveyName'").stringValue;
        surveyID = Convert.ToInt32(webService.get_system_id(projectID, true, surveyName, "survey"));
        questionaireName = dataReader.getScalarString("SELECT text FROM config WHERE ID='ingressQuestionaireName'").stringValue;
        questionaireID = Convert.ToInt32(webService.get_system_id(projectID, true, questionaireName, "questionnaire"));
        employeeIDField = dataReader.getScalarString("SELECT text FROM config WHERE ID='ingressEmployeeID'").stringValue;
        mailField = dataReader.getScalarString("SELECT text FROM config WHERE ID='ingressEmail'").stringValue;
        mobileField = dataReader.getScalarString("SELECT text FROM config WHERE ID='ingressMobile'").stringValue;
        orgIDField = dataReader.getScalarString("SELECT text FROM config WHERE ID='ingressOrgID'").stringValue;
        orgID1Field = dataReader.getScalarString("SELECT text FROM config WHERE ID='ingressOrgID1'").stringValue;
        orgID2Field = dataReader.getScalarString("SELECT text FROM config WHERE ID='ingressOrgID2'").stringValue;
        orgID3Field = dataReader.getScalarString("SELECT text FROM config WHERE ID='ingressOrgID3'").stringValue;
        postNominationField = dataReader.getScalarString("SELECT text FROM config WHERE ID='ingressPostNominationGroup'").stringValue;
        postNominationFieldID = Convert.ToInt32(webService.get_system_id(projectID, true, postNominationField, "group"));
    }
    public int updateParticipantEmailAndMobile(int aEmployeeID, string aNewEmail, string aNewMobile)
    {
        // Prozess auf Ingress auslösen
        // Teilnehmer E-Mail aktualisieren
        string participants = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><participant_import><series><" + employeeIDField + " unique=\"true\">" + aEmployeeID + "</" + employeeIDField + "><" + mailField + ">" + aNewEmail + "</" + mailField + "><" + mobileField + ">" + aNewMobile + "</" + mobileField + "></series></participant_import>";
        int import_count;
        bool import_countSpecified;
        string participants_ids;
        string extra;
        bool success = webService.add_participant(projectID, participants, out import_count, out import_countSpecified, out participants_ids, out extra);

        if (success)
            return Convert.ToInt32(participants_ids);
        else
            return 0;
    }
    public bool resetParticipantInvitation(int aTID)
    {
        return webService.reset_survey_participant(projectID, surveyID, aTID.ToString(), "email_invitation", questionaireID, true);
    }
    public bool sendParticipantInvitation(int aTID)
    {
        return webService.create_survey_email_invitation(projectID, surveyID, aTID.ToString(), DateTime.Now, true);
    }
    public int getTID(int aEmloyeeID)
    {
        string participants = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><participant_import><series><" + employeeIDField + " unique=\"true\">" + aEmloyeeID + "</" + employeeIDField + "></series></participant_import>";
        int import_count;
        bool import_countSpecified;
        string participants_ids;
        string extra;
        bool success = webService.add_participant(projectID, participants, out import_count, out import_countSpecified, out participants_ids, out extra);
        if (participants_ids.Length > 0)
            return Convert.ToInt32(participants_ids);
        else
            return 0;

    }
    public bool deleteParticipant(int aEmployeeID, string aProjectID)
    {
        int TID = getTID(aEmployeeID);

        if (TID > 0)
        {
            webService.delete_participants(projectID, TID.ToString(), true, false);
            return true;
        }
        else
        {
            return false;
        }
    }
    public string updateParticipant(string aXMLData)
    {
        // Prozess auf Ingress auslösen
        int import_count;
        bool import_countSpecified;
        string participants_ids;
        string extra;
        bool success = webService.add_participant(projectID, aXMLData, out import_count, out import_countSpecified, out participants_ids, out extra);

        if (success)
            return participants_ids;
        else
            return "";
    }
    public string updateParticipantToGroup(string aXMLData, string aTargetGroup)
    {
        // Teilnehmer der Zielgruppe zuordnen
        int import_count;
        bool import_countSpecified;
        string extra;
        string participants_ids = "";
        bool added_to_group = false;
        bool added_to_group_specified = false;
        bool group_created = false;
        bool group_created_specified = false;
        int group_id = 0;
        bool group_isSpecified = false;
        webService.add_participant_to_group(projectID, aXMLData, aTargetGroup, out import_count, out import_countSpecified, out participants_ids, out added_to_group, out added_to_group_specified, out group_created, out group_created_specified, out group_id, out group_isSpecified, out extra);

        return participants_ids;
    }
    public string buildXML(int aEmployeeID, ref string aTargetGroupResult, string aTargetGroup, bool aIndirectTargetGroup, ArrayList aMapping, TAccessCode aAccessCode, string aProjectID)
    {
        string XMLData = "";

        SqlDB dataReader = new SqlDB("Select employeeID, name, firstname, email, orgID, orgIDAdd1, orgIDAdd2, orgIDAdd3, issupervisor, issupervisor1, issupervisor2, issupervisor3, mobile, surveytype, accesscode  FROM orgmanager_employee WHERE employeeID='" + aEmployeeID + "'", aProjectID);
        if (dataReader.read())
        {
            // neuer Mitarbeiter
            TEmployeeData actEmployee = new TEmployeeData();
            actEmployee.values = new Dictionary<string, string>();

            // Levelvariablen ermitteln
            int actOrgID = dataReader.getInt32(4);
            ArrayList levels = new ArrayList();
            TAdminStructure.getTopOrgIDList(actOrgID, ref levels, aProjectID);
            int actOrg1ID = dataReader.getInt32(5);
            ArrayList levels1 = new ArrayList();
            TAdminStructure1.getTopOrgIDList(actOrg1ID, ref levels1, aProjectID);
            int actOrg2ID = dataReader.getInt32(6);
            ArrayList levels2 = new ArrayList();
            TAdminStructure2.getTopOrgIDList(actOrg2ID, ref levels2, aProjectID);
            int actOrg3ID = dataReader.getInt32(7);
            ArrayList levels3 = new ArrayList();
            TAdminStructure3.getTopOrgIDList(actOrg3ID, ref levels3, aProjectID);

            actEmployee.employeeID = dataReader.getInt32(0);

            XMLData = XMLData + "<series>";
            int actCol = 1;
            // Schleife über alle Variablen
            foreach (TIngressMapping.TMappingElement actElement in aMapping)
            {
                // Paper-Teilnehmer werden nicht nach Ingress exportiert
                int actSurveyType = dataReader.getInt32(13);
                if (((actSurveyType == 0) || (actSurveyType == 2) || (actSurveyType == 3)) && (actElement.method == 0))
                {
                    string targetname = ((TIngressMapping.TMappingElement)aMapping[actCol - 1]).targetname;
                    switch (actElement.source)
                    {
                        case "Employee":
                            #region Employee
                            // Daten kommen aus der Mitarbeiterliste
                            if (actElement.sourcecolumn == "EmployeeID")
                            {
                                XMLData = XMLData + "<" + targetname + " unique=\"true\">" + actEmployee.employeeID.ToString() + "</" + targetname + ">";
                                actEmployee.values.Add(actElement.targetname, actEmployee.employeeID.ToString());
                            }
                            else
                            {
                                if (actElement.sourcecolumn == "AccessCode")
                                {
                                    string actAccesscode = dataReader.getString(14).ToString();
                                    if (actAccesscode == "")
                                    {
                                        actAccesscode = aAccessCode.createAndSaveAccessCode(dataReader.getInt32(0));
                                    }
                                    XMLData = XMLData + "<" + targetname + ">" + actAccesscode + "</" + targetname + ">";
                                    actEmployee.values.Add("survey_mail_password", actAccesscode);
                                }
                                else
                                {
                                    if (actElement.sourcecolumn == "OrgID")
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + actOrgID.ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, actOrgID.ToString());
                                    }
                                    else
                                    {
                                        // Datenfeldtyp ermitteln
                                        string actFieldtype = "";
                                        SqlDB dataReader2;
                                        dataReader2 = new SqlDB("Select fieldtype FROM orgmanager_employeefields WHERE fieldID='" + actElement.sourcecolumn + "'", aProjectID);
                                        if (dataReader2.read())
                                            actFieldtype = dataReader2.getString(0);
                                        dataReader2.close();

                                        // Wert ermitteln
                                        switch (actFieldtype)
                                        {
                                            case "NameEmployee":
                                                XMLData = XMLData + "<" + targetname + ">" + dataReader.getString(1).ToString() + "</" + targetname + ">";
                                                actEmployee.values.Add(actElement.targetname, dataReader.getString(1).ToString());
                                                break;
                                            case "FirstnameEmployee":
                                                XMLData = XMLData + "<" + targetname + ">" + dataReader.getString(2).ToString() + "</" + targetname + ">";
                                                actEmployee.values.Add(actElement.targetname, dataReader.getString(2).ToString());
                                                break;
                                            case "EmailEmployee":
                                                if (actElement.method == 3)
                                                {
                                                    XMLData = XMLData + "<" + targetname + ">" + dataReader.getString(12).ToString() + "@sms.messagebird.com</" + targetname + ">";
                                                    actEmployee.values.Add(actElement.targetname, dataReader.getString(12).ToString() + "@sms.messagebird.com");
                                                }
                                                else
                                                {
                                                    XMLData = XMLData + "<" + targetname + ">" + dataReader.getString(3).ToString() + "</" + targetname + ">";
                                                    actEmployee.values.Add(actElement.targetname, dataReader.getString(3).ToString());
                                                }

                                                break;
                                            case "MobileEmployee":
                                                XMLData = XMLData + "<" + targetname + ">" + dataReader.getString(12).ToString() + "</" + targetname + ">";
                                                actEmployee.values.Add(actElement.targetname, dataReader.getString(12).ToString());
                                                break;
                                            case "UnitManager":
                                                XMLData = XMLData + "<" + targetname + ">" + dataReader.getInt32(8).ToString() + "</" + targetname + ">";
                                                actEmployee.values.Add(actElement.targetname, dataReader.getInt32(8).ToString());
                                                break;
                                            case "UnitManager1":
                                                XMLData = XMLData + "<" + targetname + ">" + dataReader.getInt32(9).ToString() + "</" + targetname + ">";
                                                actEmployee.values.Add(actElement.targetname, dataReader.getInt32(9).ToString());
                                                break;
                                            case "UnitManager2":
                                                XMLData = XMLData + "<" + targetname + ">" + dataReader.getInt32(10).ToString() + "</" + targetname + ">";
                                                actEmployee.values.Add(actElement.targetname, dataReader.getInt32(10).ToString());
                                                break;
                                            case "UnitManager3":
                                                XMLData = XMLData + "<" + targetname + ">" + dataReader.getInt32(11).ToString() + "</" + targetname + ">";
                                                actEmployee.values.Add(actElement.targetname, dataReader.getInt32(11).ToString());
                                                break;
                                            case "SurveyType":
                                                XMLData = XMLData + "<" + targetname + ">" + dataReader.getInt32(13).ToString() + "</" + targetname + ">";
                                                actEmployee.values.Add(actElement.targetname, dataReader.getInt32(13).ToString());
                                                break;
                                            case "Matrix1":
                                                XMLData = XMLData + "<" + targetname + ">" + dataReader.getInt32(5).ToString() + "</" + targetname + ">";
                                                actEmployee.values.Add(actElement.targetname, dataReader.getInt32(5).ToString());
                                                break;
                                            case "Matrix2":
                                                XMLData = XMLData + "<" + targetname + ">" + dataReader.getInt32(6).ToString() + "</" + targetname + ">";
                                                actEmployee.values.Add(actElement.targetname, dataReader.getInt32(6).ToString());
                                                break;
                                            case "Matrix3":
                                                XMLData = XMLData + "<" + targetname + ">" + dataReader.getInt32(7).ToString() + "</" + targetname + ">";
                                                actEmployee.values.Add(actElement.targetname, dataReader.getInt32(7).ToString());
                                                break;
                                            default:
                                                dataReader2 = new SqlDB("Select value FROM orgmanager_employeefieldsvalues WHERE employeeID='" + actEmployee.employeeID + "' AND fieldID='" + actElement.sourcecolumn + "'", aProjectID);
                                                if (dataReader2.read())
                                                {
                                                    XMLData = XMLData + "<" + targetname + ">" + dataReader2.getString(0) + "</" + targetname + ">";
                                                    actEmployee.values.Add(actElement.targetname, dataReader2.getString(0));
                                                }
                                                dataReader2.close();
                                                break;
                                        }
                                    }

                                }
                            }
                            break;
                        #endregion Employee
                        case "Structure":
                            #region Structure
                            // Daten kommen aus der Struktur
                            switch (actElement.sourcecolumn)
                            {
                                case "Level0":
                                    if (levels.Count > 0)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels[0]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[0]).ToString());

                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level1":
                                    if (levels.Count > 1)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels[1]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[1]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level2":
                                    if (levels.Count > 2)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels[2]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[2]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level3":
                                    if (levels.Count > 3)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels[3]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[3]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level4":
                                    if (levels.Count > 4)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels[4]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[4]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level5":
                                    if (levels.Count > 5)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels[5]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[5]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level6":
                                    if (levels.Count > 6)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels[6]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[6]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level7":
                                    if (levels.Count > 7)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels[7]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[7]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level8":
                                    if (levels.Count > 8)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels[8]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[8]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level9":
                                    if (levels.Count > 9)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels[9]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[9]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level10":
                                    if (levels.Count > 10)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels[10]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[10]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level11":
                                    if (levels.Count > 11)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels[11]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[11]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level12":
                                    if (levels.Count > 12)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels[12]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[12]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level13":
                                    if (levels.Count > 13)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels[13]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[13]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level14":
                                    if (levels.Count > 14)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels[14]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[14]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level15":
                                    if (levels.Count > 15)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels[15]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[15]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level16":
                                    if (levels.Count > 16)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels[16]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[16]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level17":
                                    if (levels.Count > 17)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels[17]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[17]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level18":
                                    if (levels.Count > 18)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels[18]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[18]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level19":
                                    if (levels.Count > 19)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels[19]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[19]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level20":
                                    if (levels.Count > 20)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels[20]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels[20]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                default:
                                    // Datenfeldtyp ermitteln
                                    string actFieldtype = "";
                                    SqlDB dataReader2;
                                    dataReader2 = new SqlDB("Select fieldtype FROM orgmanager_structurefields WHERE fieldID='" + actElement.sourcecolumn + "'", aProjectID);
                                    if (dataReader2.read())
                                        actFieldtype = dataReader2.getString(0);
                                    dataReader2.close();

                                    dataReader2 = new SqlDB("Select displayname, displaynameShort FROM structure WHERE orgID='" + actOrgID + "'", aProjectID);
                                    if (dataReader2.read())
                                    {
                                        SqlDB dataReader3;
                                        dataReader3 = new SqlDB("Select number, numberTotal, backComparison FROM orgmanager_structure WHERE orgID='" + actOrgID + "'", aProjectID);
                                        if (dataReader3.read())
                                        {
                                            // Wert ermitteln
                                            switch (actFieldtype)
                                            {
                                                case "UnitName":
                                                    XMLData = XMLData + "<" + targetname + ">" + dataReader2.getString(0) + "</" + targetname + ">";
                                                    actEmployee.values.Add(actElement.targetname, dataReader2.getString(0));
                                                    break;
                                                case "UnitNameShort":
                                                    XMLData = XMLData + "<" + targetname + ">" + dataReader2.getString(1) + "</" + targetname + ">";
                                                    actEmployee.values.Add(actElement.targetname, dataReader2.getString(1));
                                                    break;
                                                case "EmployeeCount":
                                                    XMLData = XMLData + "<" + targetname + ">" + dataReader3.getInt32(0).ToString() + "</" + targetname + ">";
                                                    actEmployee.values.Add(actElement.targetname, dataReader3.getInt32(0).ToString());
                                                    break;
                                                case "EmployeeCountTotal":
                                                    XMLData = XMLData + "<" + targetname + ">" + dataReader3.getInt32(1).ToString() + "</" + targetname + ">";
                                                    actEmployee.values.Add(actElement.targetname, dataReader3.getInt32(1).ToString());
                                                    break;
                                                case "BackComparision":
                                                    XMLData = XMLData + "<" + targetname + ">" + dataReader3.getInt32(2).ToString() + "</" + targetname + ">";
                                                    actEmployee.values.Add(actElement.targetname, dataReader3.getInt32(2).ToString());
                                                    break;
                                                default:
                                                    SqlDB dataReader4 = new SqlDB("Select value FROM orgmanager_structurefieldsvalues WHERE orgID='" + actOrgID + "' AND fieldID='" + actElement.sourcecolumn + "'", aProjectID);
                                                    if (dataReader4.read())
                                                    {
                                                        XMLData = XMLData + "<" + targetname + ">" + dataReader4.getString(0) + "</" + targetname + ">";
                                                        actEmployee.values.Add(actElement.targetname, dataReader4.getString(0));
                                                    }
                                                    dataReader4.close();
                                                    break;
                                            }
                                        }
                                        dataReader3.close();
                                    }
                                    dataReader2.close();
                                    break;
                            }
                            break;
                        #endregion Structure
                        case "Structure1":
                            #region Structure1
                            // Daten kommen aus der Struktur1
                            switch (actElement.sourcecolumn)
                            {
                                case "Level0":
                                    if (levels1.Count > 0)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels1[0]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[0]).ToString());

                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level1":
                                    if (levels1.Count > 1)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels1[1]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[1]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level2":
                                    if (levels1.Count > 2)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels1[2]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[2]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level3":
                                    if (levels1.Count > 3)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels1[3]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[3]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level4":
                                    if (levels1.Count > 4)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels1[4]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[4]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level5":
                                    if (levels1.Count > 5)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels1[5]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[5]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level6":
                                    if (levels1.Count > 6)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels1[6]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[6]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level7":
                                    if (levels1.Count > 7)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels1[7]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[7]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level8":
                                    if (levels1.Count > 8)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels1[8]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[8]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level9":
                                    if (levels1.Count > 9)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels1[9]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[9]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level10":
                                    if (levels1.Count > 10)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels1[10]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[10]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level11":
                                    if (levels1.Count > 11)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels1[11]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[11]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level12":
                                    if (levels1.Count > 12)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels1[12]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[12]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level13":
                                    if (levels1.Count > 13)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels1[13]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[13]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level14":
                                    if (levels1.Count > 14)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels1[14]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[14]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level15":
                                    if (levels1.Count > 15)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels1[15]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[15]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level16":
                                    if (levels1.Count > 16)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels1[16]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[16]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level17":
                                    if (levels1.Count > 17)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels1[17]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[17]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level18":
                                    if (levels1.Count > 18)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels1[18]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[18]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level19":
                                    if (levels1.Count > 19)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels1[19]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[19]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level20":
                                    if (levels1.Count > 20)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels1[20]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels1[20]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                default:
                                    // Datenfeldtyp ermitteln
                                    string actFieldtype = "";
                                    SqlDB dataReader2;
                                    dataReader2 = new SqlDB("Select fieldtype FROM orgmanager_structure1fields WHERE fieldID='" + actElement.sourcecolumn + "'", aProjectID);
                                    if (dataReader2.read())
                                        actFieldtype = dataReader2.getString(0);
                                    dataReader2.close();

                                    dataReader2 = new SqlDB("Select displayname, displaynameShort FROM structure1 WHERE orgID='" + actOrg1ID + "'", aProjectID);
                                    if (dataReader2.read())
                                    {
                                        SqlDB dataReader3;
                                        dataReader3 = new SqlDB("Select number, numberTotal, backComparison FROM orgmanager_structure1 WHERE orgID='" + actOrg1ID + "'", aProjectID);
                                        if (dataReader3.read())
                                        {
                                            // Wert ermitteln
                                            switch (actFieldtype)
                                            {
                                                case "UnitName":
                                                    XMLData = XMLData + "<" + targetname + ">" + dataReader2.getString(0) + "</" + targetname + ">";
                                                    actEmployee.values.Add(actElement.targetname, dataReader2.getString(0));
                                                    break;
                                                case "UnitNameShort":
                                                    XMLData = XMLData + "<" + targetname + ">" + dataReader2.getString(1) + "</" + targetname + ">";
                                                    actEmployee.values.Add(actElement.targetname, dataReader2.getString(1));
                                                    break;
                                                case "EmployeeCount":
                                                    XMLData = XMLData + "<" + targetname + ">" + dataReader3.getInt32(0).ToString() + "</" + targetname + ">";
                                                    actEmployee.values.Add(actElement.targetname, dataReader3.getInt32(0).ToString());
                                                    break;
                                                case "EmployeeCountTotal":
                                                    XMLData = XMLData + "<" + targetname + ">" + dataReader3.getInt32(1).ToString() + "</" + targetname + ">";
                                                    actEmployee.values.Add(actElement.targetname, dataReader3.getInt32(1).ToString());
                                                    break;
                                                case "BackComparision":
                                                    XMLData = XMLData + "<" + targetname + ">" + dataReader3.getInt32(2).ToString() + "</" + targetname + ">";
                                                    actEmployee.values.Add(actElement.targetname, dataReader3.getInt32(2).ToString());
                                                    break;
                                                default:
                                                    SqlDB dataReader4 = new SqlDB("Select value FROM orgmanager_structure1fieldsvalues WHERE orgID='" + actOrg1ID + "' AND fieldID='" + actElement.sourcecolumn + "'", aProjectID);
                                                    if (dataReader4.read())
                                                    {
                                                        XMLData = XMLData + "<" + targetname + ">" + dataReader4.getString(0) + "</" + targetname + ">";
                                                        actEmployee.values.Add(actElement.targetname, dataReader4.getString(0));
                                                    }
                                                    dataReader4.close();
                                                    break;
                                            }
                                        }
                                        dataReader3.close();
                                    }
                                    dataReader2.close();
                                    break;
                            }
                            break;
                        #endregion Structure1
                        case "Structure2":
                            #region Structure2
                            // Daten kommen aus der Struktur2
                            switch (actElement.sourcecolumn)
                            {
                                case "Level0":
                                    if (levels2.Count > 0)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels2[0]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[0]).ToString());

                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level1":
                                    if (levels2.Count > 1)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels2[1]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[1]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level2":
                                    if (levels2.Count > 2)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels2[2]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[2]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level3":
                                    if (levels2.Count > 3)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels2[3]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[3]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level4":
                                    if (levels2.Count > 4)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels2[4]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[4]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level5":
                                    if (levels2.Count > 5)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels2[5]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[5]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level6":
                                    if (levels2.Count > 6)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels2[6]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[6]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level7":
                                    if (levels2.Count > 7)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels2[7]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[7]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level8":
                                    if (levels2.Count > 8)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels2[8]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[8]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level9":
                                    if (levels2.Count > 9)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels2[9]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[9]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level10":
                                    if (levels2.Count > 10)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels2[10]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[10]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level11":
                                    if (levels2.Count > 11)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels2[11]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[11]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level12":
                                    if (levels2.Count > 12)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels2[12]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[12]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level13":
                                    if (levels2.Count > 13)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels2[13]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[13]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level14":
                                    if (levels2.Count > 14)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels2[14]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[14]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level15":
                                    if (levels2.Count > 15)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels2[15]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[15]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level16":
                                    if (levels2.Count > 16)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels2[16]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[16]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level17":
                                    if (levels2.Count > 17)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels2[17]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[17]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level18":
                                    if (levels2.Count > 18)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels2[18]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[18]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level19":
                                    if (levels2.Count > 19)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels2[19]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[19]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level20":
                                    if (levels2.Count > 20)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels2[20]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels2[20]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                default:
                                    // Datenfeldtyp ermitteln
                                    string actFieldtype = "";
                                    SqlDB dataReader2;
                                    dataReader2 = new SqlDB("Select fieldtype FROM orgmanager_structure2fields WHERE fieldID='" + actElement.sourcecolumn + "'", aProjectID);
                                    if (dataReader2.read())
                                        actFieldtype = dataReader2.getString(0);
                                    dataReader2.close();

                                    dataReader2 = new SqlDB("Select displayname, displaynameShort FROM structure2 WHERE orgID='" + actOrg2ID + "'", aProjectID);
                                    if (dataReader2.read())
                                    {
                                        SqlDB dataReader3;
                                        dataReader3 = new SqlDB("Select number, numberTotal, backComparison FROM orgmanager_structure2 WHERE orgID='" + actOrg2ID + "'", aProjectID);
                                        if (dataReader3.read())
                                        {
                                            // Wert ermitteln
                                            switch (actFieldtype)
                                            {
                                                case "UnitName":
                                                    XMLData = XMLData + "<" + targetname + ">" + dataReader2.getString(0) + "</" + targetname + ">";
                                                    actEmployee.values.Add(actElement.targetname, dataReader2.getString(0));
                                                    break;
                                                case "UnitNameShort":
                                                    XMLData = XMLData + "<" + targetname + ">" + dataReader2.getString(1) + "</" + targetname + ">";
                                                    actEmployee.values.Add(actElement.targetname, dataReader2.getString(1));
                                                    break;
                                                case "EmployeeCount":
                                                    XMLData = XMLData + "<" + targetname + ">" + dataReader3.getInt32(0).ToString() + "</" + targetname + ">";
                                                    actEmployee.values.Add(actElement.targetname, dataReader3.getInt32(0).ToString());
                                                    break;
                                                case "EmployeeCountTotal":
                                                    XMLData = XMLData + "<" + targetname + ">" + dataReader3.getInt32(1).ToString() + "</" + targetname + ">";
                                                    actEmployee.values.Add(actElement.targetname, dataReader3.getInt32(1).ToString());
                                                    break;
                                                case "BackComparision":
                                                    XMLData = XMLData + "<" + targetname + ">" + dataReader3.getInt32(2).ToString() + "</" + targetname + ">";
                                                    actEmployee.values.Add(actElement.targetname, dataReader3.getInt32(2).ToString());
                                                    break;
                                                default:
                                                    SqlDB dataReader4 = new SqlDB("Select value FROM orgmanager_structure2fieldsvalues WHERE orgID='" + actOrg2ID + "' AND fieldID='" + actElement.sourcecolumn + "'", aProjectID);
                                                    if (dataReader4.read())
                                                    {
                                                        XMLData = XMLData + "<" + targetname + ">" + dataReader4.getString(0) + "</" + targetname + ">";
                                                        actEmployee.values.Add(actElement.targetname, dataReader4.getString(0));
                                                    }
                                                    dataReader4.close();
                                                    break;
                                            }
                                        }
                                        dataReader3.close();
                                    }
                                    dataReader2.close();
                                    break;
                            }
                            break;
                        #endregion Structure2
                        case "Structure3":
                            #region Structure3
                            // Daten kommen aus der Struktur3
                            switch (actElement.sourcecolumn)
                            {
                                case "Level0":
                                    if (levels3.Count > 0)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels3[0]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[0]).ToString());

                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level1":
                                    if (levels3.Count > 1)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels3[1]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[1]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level2":
                                    if (levels3.Count > 2)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels3[2]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[2]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level3":
                                    if (levels3.Count > 3)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels3[3]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[3]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level4":
                                    if (levels3.Count > 4)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels3[4]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[4]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level5":
                                    if (levels3.Count > 5)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels3[5]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[5]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level6":
                                    if (levels3.Count > 6)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels3[6]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[6]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level7":
                                    if (levels3.Count > 7)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels3[7]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[7]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level8":
                                    if (levels3.Count > 8)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels3[8]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[8]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level9":
                                    if (levels3.Count > 9)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels3[9]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[9]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level10":
                                    if (levels3.Count > 10)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels3[10]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[10]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level11":
                                    if (levels3.Count > 11)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels3[11]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[11]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level12":
                                    if (levels3.Count > 12)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels3[12]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[12]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level13":
                                    if (levels3.Count > 13)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels3[13]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[13]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level14":
                                    if (levels3.Count > 14)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels3[14]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[14]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level15":
                                    if (levels3.Count > 15)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels3[15]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[15]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level16":
                                    if (levels3.Count > 16)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels3[16]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[16]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level17":
                                    if (levels3.Count > 17)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels3[17]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[17]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level18":
                                    if (levels3.Count > 18)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels3[18]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[18]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level19":
                                    if (levels3.Count > 19)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels3[19]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[19]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                case "Level20":
                                    if (levels3.Count > 20)
                                    {
                                        XMLData = XMLData + "<" + targetname + ">" + ((int)levels3[20]).ToString() + "</" + targetname + ">";
                                        actEmployee.values.Add(actElement.targetname, ((int)levels3[20]).ToString());
                                    }
                                    else
                                        actEmployee.values.Add(actElement.targetname, "");
                                    break;
                                default:
                                    // Datenfeldtyp ermitteln
                                    string actFieldtype = "";
                                    SqlDB dataReader2;
                                    dataReader2 = new SqlDB("Select fieldtype FROM orgmanager_structure3fields WHERE fieldID='" + actElement.sourcecolumn + "'", aProjectID);
                                    if (dataReader2.read())
                                        actFieldtype = dataReader2.getString(0);
                                    dataReader2.close();

                                    dataReader2 = new SqlDB("Select displayname, displaynameShort FROM structure3 WHERE orgID='" + actOrg3ID + "'", aProjectID);
                                    if (dataReader2.read())
                                    {
                                        SqlDB dataReader3;
                                        dataReader3 = new SqlDB("Select number, numberTotal, backComparison FROM orgmanager_structure3 WHERE orgID='" + actOrg3ID + "'", aProjectID);
                                        if (dataReader3.read())
                                        {
                                            // Wert ermitteln
                                            switch (actFieldtype)
                                            {
                                                case "UnitName":
                                                    XMLData = XMLData + "<" + targetname + ">" + dataReader2.getString(0) + "</" + targetname + ">";
                                                    actEmployee.values.Add(actElement.targetname, dataReader2.getString(0));
                                                    break;
                                                case "UnitNameShort":
                                                    XMLData = XMLData + "<" + targetname + ">" + dataReader2.getString(1) + "</" + targetname + ">";
                                                    actEmployee.values.Add(actElement.targetname, dataReader2.getString(1));
                                                    break;
                                                case "EmployeeCount":
                                                    XMLData = XMLData + "<" + targetname + ">" + dataReader3.getInt32(0).ToString() + "</" + targetname + ">";
                                                    actEmployee.values.Add(actElement.targetname, dataReader3.getInt32(0).ToString());
                                                    break;
                                                case "EmployeeCountTotal":
                                                    XMLData = XMLData + "<" + targetname + ">" + dataReader3.getInt32(1).ToString() + "</" + targetname + ">";
                                                    actEmployee.values.Add(actElement.targetname, dataReader3.getInt32(1).ToString());
                                                    break;
                                                case "BackComparision":
                                                    XMLData = XMLData + "<" + targetname + ">" + dataReader3.getInt32(2).ToString() + "</" + targetname + ">";
                                                    actEmployee.values.Add(actElement.targetname, dataReader3.getInt32(2).ToString());
                                                    break;
                                                default:
                                                    SqlDB dataReader4 = new SqlDB("Select value FROM orgmanager_structure3fieldsvalues WHERE orgID='" + actOrg3ID + "' AND fieldID='" + actElement.sourcecolumn + "'", aProjectID);
                                                    if (dataReader4.read())
                                                    {
                                                        XMLData = XMLData + "<" + targetname + ">" + dataReader4.getString(0) + "</" + targetname + ">";
                                                        actEmployee.values.Add(actElement.targetname, dataReader4.getString(0));
                                                    }
                                                    dataReader4.close();
                                                    break;
                                            }
                                        }
                                        dataReader3.close();
                                    }
                                    dataReader2.close();
                                    break;
                            }
                            break;
                        #endregion Structure3
                        default:
                            #region Zusatzwerte
                            // Daten kommen aus den Zusatzwerten
                            // Schlüsselwerte ermitteln
                            string key1value = "";
                            if (actElement.key1 != "")
                            {
                                key1value = actEmployee.values[actElement.key1];
                            }
                            string key2value = "";
                            if (actElement.key2 != "")
                            {
                                key2value = actEmployee.values[actElement.key2];
                            }
                            string key3value = "";
                            if (actElement.key3 != "")
                            {
                                key3value = actEmployee.values[actElement.key3];
                            }
                            string key4value = "";
                            if (actElement.key4 != "")
                            {
                                key4value = actEmployee.values[actElement.key4];
                            }
                            string key5value = "";
                            if (actElement.key5 != "")
                            {
                                key5value = actEmployee.values[actElement.key5];
                            }
                            // Wert ermitteln
                            SqlDB dataReader1;
                            dataReader1 = new SqlDB("Select value FROM orgmanager_ingressvalues WHERE source='" + actElement.source + "' AND key1='" + key1value + "' AND key2='" + key2value + "' AND key3='" + key3value + "' AND key4='" + key4value + "' AND key5='" + key5value + "' AND valuecolumn='" + actElement.sourcecolumn + "'", aProjectID);
                            if (dataReader1.read())
                            {
                                XMLData = XMLData + "<" + targetname + ">" + dataReader1.getString(0) + "</" + targetname + ">";
                                actEmployee.values.Add(actElement.targetname, dataReader1.getString(0));
                            }
                            else
                            {
                                XMLData = XMLData + "<" + targetname + ">" + "</" + targetname + ">";
                                actEmployee.values.Add(actElement.targetname, "");
                            }

                            dataReader1.close();
                            break;
                            #endregion Zusatzwerte
                    }
                }
                actCol++;
            }
            XMLData = XMLData + "</series>";

            // Zielgruppe bestimmen
            if (aIndirectTargetGroup)
                aTargetGroupResult = actEmployee.values[aTargetGroup];
            else
                aTargetGroupResult = aTargetGroup;

        }
        dataReader.close();
        return XMLData;
    }
    public void addGroupToSurvey(string aGroupName)
    {
        // GroupID ermitteln
        //### hier kommt u.U. eie Komma-separierte Auflistung zurück, dann müsste der höchste Wert genommen werden
        int groupID = Convert.ToInt32(webService.get_system_id(projectID, true, aGroupName, "group"));

        string extra;
        webService.add_group_to_survey(projectID, surveyID, groupID, out extra);
    }
    public static string startSync(string aProjectID)
    {
        string errorMessage = "[Sync] ";
        try
        {
            TIngress webService = new TIngress(aProjectID);

            SqlDB dataReader = new SqlDB(aProjectID);

            // Ingress-ID ermitteln
            //int projectID = Convert.ToInt32(myWebserviceClient.get_system_id(0, false, webService.projectName, "project"));
            int surveyID = webService.surveyID;
            int groupID = 0;
            //int questionaireID = Convert.ToInt32(myWebserviceClient.get_system_id(projectID, true, webService.questionaireName, "questionnaire"));

            System.Collections.Generic.Dictionary<int, int> responseList = new System.Collections.Generic.Dictionary<int, int>();
            System.Collections.Generic.Dictionary<int, int> response1List = new System.Collections.Generic.Dictionary<int, int>();
            System.Collections.Generic.Dictionary<int, int> response2List = new System.Collections.Generic.Dictionary<int, int>();
            System.Collections.Generic.Dictionary<int, int> response3List = new System.Collections.Generic.Dictionary<int, int>();

            string actDate = DateTime.Now.AddDays(1).ToShortDateString();
            // TANs der abgeschlossenen Fragebögen zu ermitteln
            //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Start TANS der abgeschlossneen FB ermitteln";
            string tans = webService.webService.get_questionnaire_export_info(webService.questionaireID, webService.projectID, surveyID, false, groupID, false, 5, true, "01.01.2017", actDate);
            //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Ende TANS der abgeschlossneen FB ermitteln";

            if (tans != "")
            {
                // TIDs der zugehörigen Teilnehmer ermitteln
                // immer je 1000, da get_questionnaire_export nur maximal 1000 Ergebnisdatensätze liefern kann
                string result = "";
                string[] tanList = tans.Split(',');
                int count = tanList.Length;
                int start = 0;
                string actTans = "";
                while (start <= count)
                {
                    //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Start Ergebnisdatensatz der Teilnehmer ermitteln";
                    // maximal 1000 Ergebnisdatensätze laden
                    actTans = "";
                    int end = start + Math.Min(1000, count - start);
                    for (int i = start; i < end; i++)
                        actTans = actTans + ", " + tanList[i];
                    actTans = actTans.Remove(0, 1);

                    result = webService.webService.get_questionnaire_export(webService.projectID, webService.questionaireID, actTans, "");
                    //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Ende Ergebnisdatensatz der Teilnehmer ermitteln";

                    //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Start TIDs der Teilnehmer ermitteln";
                    // XMLladen und TIDs ermitteln
                    System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
                    doc.LoadXml(result);
                    string TIDs = "";
                    System.Xml.XmlNodeList xmlNode = doc.GetElementsByTagName("TID");
                    for (int i = 0; i < xmlNode.Count; i++)
                    {
                        if (TIDs == "")
                            TIDs = xmlNode.Item(i).InnerText;
                        else
                            TIDs = TIDs + ", " + xmlNode.Item(i).InnerText;
                    }
                    //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Ende TIDs der Teilnehmer ermitteln";
                    //LblTIDs.Text = "<br/><br/>TIDs:<br/>" + TIDs;

                    //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Start OrgIDs der Teilnehmer ermitteln";
                    // OrgIDs laden und in Ergebnistabelle schreiben
                    result = webService.webService.get_participant_export(webService.projectID, TIDs, groupID, true, false, true);
                    doc.LoadXml(result);
                    //string OrgIDs = "";
                    // OrgID
                    if (webService.orgIDField != "")
                    {
                        xmlNode = doc.GetElementsByTagName(webService.orgIDField);
                        for (int i = 0; i < xmlNode.Count; i++)
                        {
                            // OrgIDs zusammenfassen
                            int actOrgID;
                            if (Int32.TryParse(xmlNode.Item(i).InnerText, out actOrgID))
                            {
                                int value;
                                if (responseList.TryGetValue(actOrgID, out value))
                                    responseList[actOrgID] += 1;
                                else
                                    responseList.Add(actOrgID, 1);

                                //if (OrgIDs == "")
                                //    OrgIDs = xmlNode.Item(i).InnerText;
                                //else
                                //    OrgIDs = OrgIDs + "<br/>" + xmlNode.Item(i).InnerText;
                            }
                            else
                            {
                                errorMessage += "#" + xmlNode.Item(i).ToString() + "#" + xmlNode.Item(i).InnerText + "#";
                            }
                        }
                    }
                    //LblTIDs.Text = LblTIDs.Text + "<br/><br/>OrgIDs:<br/>" + OrgIDs;
                    //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Ende OrgIDs der Teilnehmer ermitteln";
                    // OrgID1
                    if (webService.orgID1Field != "")
                    {
                        xmlNode = doc.GetElementsByTagName(webService.orgID1Field);
                        for (int i = 0; i < xmlNode.Count; i++)
                        {
                            // OrgIDs zusammenfassen
                            int actOrgID;
                            if (Int32.TryParse(xmlNode.Item(i).InnerText, out actOrgID))
                            {
                                int value;
                                if (response1List.TryGetValue(actOrgID, out value))
                                    response1List[actOrgID] += 1;
                                else
                                    response1List.Add(actOrgID, 1);

                                //if (OrgIDs == "")
                                //    OrgIDs = xmlNode.Item(i).InnerText;
                                //else
                                //    OrgIDs = OrgIDs + "<br/>" + xmlNode.Item(i).InnerText;
                            }
                            else
                            {
                                errorMessage += "#" + xmlNode.Item(i).ToString() + "#" + xmlNode.Item(i).InnerText + "#";
                            }
                        }
                    }
                    //LblTIDs.Text = LblTIDs.Text + "<br/><br/>OrgIDs:<br/>" + OrgIDs;
                    //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Ende OrgIDs der Teilnehmer ermitteln";
                    // OrgID2
                    if (webService.orgID2Field != "")
                    {
                        xmlNode = doc.GetElementsByTagName(webService.orgID2Field);
                        for (int i = 0; i < xmlNode.Count; i++)
                        {
                            // OrgIDs zusammenfassen
                            int actOrgID;
                            if (Int32.TryParse(xmlNode.Item(i).InnerText, out actOrgID))
                            {
                                int value;
                                if (response2List.TryGetValue(actOrgID, out value))
                                    response2List[actOrgID] += 1;
                                else
                                    response2List.Add(actOrgID, 1);

                                //if (OrgIDs == "")
                                //    OrgIDs = xmlNode.Item(i).InnerText;
                                //else
                                //    OrgIDs = OrgIDs + "<br/>" + xmlNode.Item(i).InnerText;
                            }
                            else
                            {
                                errorMessage += "#" + xmlNode.Item(i).ToString() + "#" + xmlNode.Item(i).InnerText + "#";
                            }
                        }
                    }
                    //LblTIDs.Text = LblTIDs.Text + "<br/><br/>OrgIDs:<br/>" + OrgIDs;
                    //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Ende OrgIDs der Teilnehmer ermitteln";
                    // OrgID3
                    if (webService.orgID3Field != "")
                    {
                        xmlNode = doc.GetElementsByTagName(webService.orgID3Field);
                        for (int i = 0; i < xmlNode.Count; i++)
                        {
                            // OrgIDs zusammenfassen
                            int actOrgID;
                            if (Int32.TryParse(xmlNode.Item(i).InnerText, out actOrgID))
                            {
                                int value;
                                if (response3List.TryGetValue(actOrgID, out value))
                                    response3List[actOrgID] += 1;
                                else
                                    response3List.Add(actOrgID, 1);

                                //if (OrgIDs == "")
                                //    OrgIDs = xmlNode.Item(i).InnerText;
                                //else
                                //    OrgIDs = OrgIDs + "<br/>" + xmlNode.Item(i).InnerText;
                            }
                            else
                            {
                                errorMessage += "#" + xmlNode.Item(i).ToString() + "#" + xmlNode.Item(i).InnerText + "#";
                            }
                        }
                    }
                    //LblTIDs.Text = LblTIDs.Text + "<br/><br/>OrgIDs:<br/>" + OrgIDs;
                    //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Ende OrgIDs der Teilnehmer ermitteln";

                    start = start + 1000;
                }
            }

            // OrgID
            if (webService.orgIDField != "")
            {
                // Liste aller OrgIDs aus der Struktur lesen
                //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Start OrgIDs aus Struktur lesen";
                ArrayList orgIDList = new ArrayList();
                dataReader = new SqlDB("select orgID from structure", aProjectID);
                while (dataReader.read())
                {
                    orgIDList.Add(dataReader.getInt32(0));
                }
                dataReader.close();
                //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Ende OrgIDs aus Struktur lesen";

                //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Start Rückläufe in DB schreiben";
                // Rückläufe in DB schreiben
                foreach (int tempOrgID in orgIDList)
                {
                    // Papierrücklauf ermitteln
                    int paperResonse = 0;
                    dataReader = new SqlDB("SELECT numberPaper FROM response_structure WHERE orgID='" + tempOrgID.ToString() + "'", aProjectID);
                    if (dataReader.read())
                    {
                        paperResonse = dataReader.getInt32(0);
                    }
                    dataReader.close();
                    // Onlinerücklauf ermitteln
                    int onlineResonse;
                    if (!responseList.TryGetValue(tempOrgID, out onlineResonse))
                        onlineResonse = 0;
                    // Gesamtrücklauf in DB schreiben
                    dataReader = new SqlDB(aProjectID);
                    dataReader.execSQL("UPDATE response_structure SET number='" + (paperResonse + onlineResonse).ToString() + "' WHERE orgID='" + tempOrgID.ToString() + "'");
                }
                //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Ende Rückläufe in DB schreiben";
                //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Start Summen neu berechnen";
                // Summen neu berechnen
                TAdminStructure.recalc(aProjectID);
                //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Ende Summen neu berechnen";
            }
            // OrgID1
            if (webService.orgID1Field != "")
            {
                // Liste aller OrgIDs aus der Struktur lesen
                //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Start OrgIDs aus Struktur lesen";
                ArrayList orgIDList = new ArrayList();
                dataReader = new SqlDB("select orgID from structure1", aProjectID);
                while (dataReader.read())
                {
                    orgIDList.Add(dataReader.getInt32(0));
                }
                dataReader.close();
                //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Ende OrgIDs aus Struktur lesen";

                //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Start Rückläufe in DB schreiben";
                // Rückläufe in DB schreiben
                foreach (int tempOrgID in orgIDList)
                {
                    // Papierrücklauf ermitteln
                    int paperResonse = 0;
                    dataReader = new SqlDB("SELECT numberPaper FROM response_structure1 WHERE orgID='" + tempOrgID.ToString() + "'", aProjectID);
                    if (dataReader.read())
                    {
                        paperResonse = dataReader.getInt32(0);
                    }
                    dataReader.close();
                    // Onlinerücklauf ermitteln
                    int onlineResonse;
                    if (!response1List.TryGetValue(tempOrgID, out onlineResonse))
                        onlineResonse = 0;
                    // Gesamtrücklauf in DB schreiben
                    dataReader = new SqlDB(aProjectID);
                    dataReader.execSQL("UPDATE response_structure1 SET number='" + (paperResonse + onlineResonse).ToString() + "' WHERE orgID='" + tempOrgID.ToString() + "'");
                }
                //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Ende Rückläufe in DB schreiben";
                //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Start Summen neu berechnen";
                // Summen neu berechnen
                TAdminStructure1.recalc(aProjectID);
                //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Ende Summen neu berechnen";
            }
            // OrgID2
            if (webService.orgID2Field != "")
            {
                // Liste aller OrgIDs aus der Struktur lesen
                //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Start OrgIDs aus Struktur lesen";
                ArrayList orgIDList = new ArrayList();
                dataReader = new SqlDB("select orgID from structure2", aProjectID);
                while (dataReader.read())
                {
                    orgIDList.Add(dataReader.getInt32(0));
                }
                dataReader.close();
                //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Ende OrgIDs aus Struktur lesen";

                //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Start Rückläufe in DB schreiben";
                // Rückläufe in DB schreiben
                foreach (int tempOrgID in orgIDList)
                {
                    // Papierrücklauf ermitteln
                    int paperResonse = 0;
                    dataReader = new SqlDB("SELECT numberPaper FROM response_structure2 WHERE orgID='" + tempOrgID.ToString() + "'", aProjectID);
                    if (dataReader.read())
                    {
                        paperResonse = dataReader.getInt32(0);
                    }
                    dataReader.close();
                    // Onlinerücklauf ermitteln
                    int onlineResonse;
                    if (!response2List.TryGetValue(tempOrgID, out onlineResonse))
                        onlineResonse = 0;
                    // Gesamtrücklauf in DB schreiben
                    dataReader = new SqlDB(aProjectID);
                    dataReader.execSQL("UPDATE response_structure2 SET number='" + (paperResonse + onlineResonse).ToString() + "' WHERE orgID='" + tempOrgID.ToString() + "'");
                }
                //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Ende Rückläufe in DB schreiben";
                //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Start Summen neu berechnen";
                // Summen neu berechnen
                TAdminStructure2.recalc(aProjectID);
                //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Ende Summen neu berechnen";
            }
            // OrgID3
            if (webService.orgID3Field != "")
            {
                // Liste aller OrgIDs aus der Struktur lesen
                //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Start OrgIDs aus Struktur lesen";
                ArrayList orgIDList = new ArrayList();
                dataReader = new SqlDB("select orgID from structure3", aProjectID);
                while (dataReader.read())
                {
                    orgIDList.Add(dataReader.getInt32(0));
                }
                dataReader.close();
                //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Ende OrgIDs aus Struktur lesen";

                //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Start Rückläufe in DB schreiben";
                // Rückläufe in DB schreiben
                foreach (int tempOrgID in orgIDList)
                {
                    // Papierrücklauf ermitteln
                    int paperResonse = 0;
                    dataReader = new SqlDB("SELECT numberPaper FROM response_structure3 WHERE orgID='" + tempOrgID.ToString() + "'", aProjectID);
                    if (dataReader.read())
                    {
                        paperResonse = dataReader.getInt32(0);
                    }
                    dataReader.close();
                    // Onlinerücklauf ermitteln
                    int onlineResonse;
                    if (!response3List.TryGetValue(tempOrgID, out onlineResonse))
                        onlineResonse = 0;
                    // Gesamtrücklauf in DB schreiben
                    dataReader = new SqlDB(aProjectID);
                    dataReader.execSQL("UPDATE response_structure3 SET number='" + (paperResonse + onlineResonse).ToString() + "' WHERE orgID='" + tempOrgID.ToString() + "'");
                }
                //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Ende Rückläufe in DB schreiben";
                //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Start Summen neu berechnen";
                // Summen neu berechnen
                TAdminStructure3.recalc(aProjectID);
                //LblTIDs.Text = LblTIDs.Text + "<br />" + DateTime.Now.ToString() + " Ende Summen neu berechnen";
            }

            // Aktualisierung in DB schreiben
            TParameterList parameterList = new TParameterList();
            parameterList.addParameter("ingressLastAutoUpdate", "string", DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt", new CultureInfo("en-US", false)));
            dataReader = new SqlDB(aProjectID);
            dataReader.execSQLwithParameter("UPDATE response_config SET text=@ingressLastAutoUpdate WHERE ID='ingressLastAutoUpdate'", parameterList);
        }
        catch
        {
        }
        if (errorMessage == "[Sync] ")
            return "[Sync] erfolgreich";
        else
            return errorMessage;
    }
}
